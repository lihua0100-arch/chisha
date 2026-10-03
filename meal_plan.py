#!/usr/bin/env python3
"""一个月家常菜谱生成器（只用标准库）。

  plan   按预算、天数、口味生成 Excel（每日做饭 / 菜谱总览 / 采购清单 / 做法大全 / 单价表），
         并把这份菜谱存成「当前菜谱」，供 today 查询。
  today  查当前菜谱里某一天吃什么（默认今天），连用料、调料、做法一起输出。
  prices 查上海市最新主副食品均价（上海市政府每个工作日发布的价格信息表）。
  dishes 列出菜库。

价格：每次 plan 先抓 shanghai.gov.cn「上海市主要主副食品品种价格信息表」最新一期的 xls，
取全市均价；抓不到就用内置兜底价并在表里注明。菜库与食材见 meal_plan_dishes.py。
菜谱存在 $HERMES_HOME/workspace/meal-plans/（每个微信号一份，互不相通）。
"""
import argparse
import datetime as dt
import hashlib
import json
import math
import os
import random
import re
import struct
import sys
import urllib.request
import zipfile
from xml.sax.saxutils import escape

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import meal_plan_dishes as M  # noqa: E402

TZ = dt.timezone(dt.timedelta(hours=8))
GOV_BASE = "https://www.shanghai.gov.cn"
GOV_LIST = GOV_BASE + "/nw17239/index{}.html"
GOV_TITLE = "上海市主要主副食品品种价格信息表"
CACHE_DIR = "/tmp/meal-plan-cache"
UA = {"User-Agent": "Mozilla/5.0 (meal-plan)"}
WEEK = "一二三四五六日"
CUP_RICE_JIN = 0.3   # 一杯米（电饭煲量杯）约 3 两
SPOON_ML = 15


def now():
    return dt.datetime.now(TZ)


def home_dir():
    home = os.environ.get("HERMES_HOME") or os.path.expanduser("~/.hermes")
    return os.path.join(home, "workspace", "meal-plans")


# ═══════════════════════════ xls（BIFF8）最小读取器 ═══════════════════════════
def _ole_stream(data, want=("Workbook", "Book")):
    if data[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise ValueError("不是 xls 文件")
    ssz = 1 << struct.unpack_from("<H", data, 0x1E)[0]
    msz = 1 << struct.unpack_from("<H", data, 0x20)[0]
    nfat, dir1 = struct.unpack_from("<II", data, 0x2C)
    cutoff, mfat1, nmfat, difat1, ndifat = struct.unpack_from("<IIIII", data, 0x38)

    def sec(i):
        o = 512 + i * ssz
        return data[o:o + ssz]

    difat = list(struct.unpack_from("<109I", data, 0x4C))
    s = difat1
    for _ in range(ndifat):
        if s >= 0xFFFFFFFA:
            break
        vals = struct.unpack("<%dI" % (ssz // 4), sec(s))
        difat += vals[:-1]
        s = vals[-1]
    fat = []
    for i in difat[:nfat]:
        fat += struct.unpack("<%dI" % (ssz // 4), sec(i))

    def chain(start):
        out, s, seen = [], start, 0
        while s < 0xFFFFFFFA and seen < len(fat) + 1:
            out.append(s)
            s = fat[s]
            seen += 1
        return out

    dirdata = b"".join(sec(i) for i in chain(dir1))
    entries = []
    for o in range(0, len(dirdata), 128):
        e = dirdata[o:o + 128]
        nlen = struct.unpack_from("<H", e, 0x40)[0]
        name = e[:max(0, nlen - 2)].decode("utf-16-le", "ignore")
        typ = e[0x42]
        start, size = struct.unpack_from("<II", e, 0x74)
        entries.append((name, typ, start, size))
    root = entries[0]
    for name, typ, start, size in entries:
        if typ == 2 and name in want:
            if size >= cutoff:
                return b"".join(sec(i) for i in chain(start))[:size]
            # 小流：在 mini stream 里
            mfat = []
            for i in chain(mfat1):
                mfat += struct.unpack("<%dI" % (ssz // 4), sec(i))
            ministream = b"".join(sec(i) for i in chain(root[2]))
            out, s = [], start
            while s < 0xFFFFFFFA:
                out.append(ministream[s * msz:(s + 1) * msz])
                s = mfat[s]
            return b"".join(out)[:size]
    raise ValueError("xls 里没有 Workbook 流")


def _rk(v):
    if v & 2:
        n = v >> 2
        if n & 0x20000000:
            n -= 0x40000000
        val = float(n)
    else:
        val = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    return val / 100 if v & 1 else val


class _Seg:
    """跨 CONTINUE 记录读取 SST。"""

    def __init__(self, parts):
        self.parts, self.i, self.p = parts, 0, 0

    def _need(self):
        while self.p >= len(self.parts[self.i]):
            self.i += 1
            self.p = 0

    def raw(self, n):
        out = b""
        while n > 0:
            self._need()
            chunk = self.parts[self.i][self.p:self.p + n]
            out += chunk
            self.p += len(chunk)
            n -= len(chunk)
        return out

    def chars(self, n, high):
        out = []
        while n > 0:
            if self.p >= len(self.parts[self.i]):
                self.i += 1
                self.p = 0
                high = self.parts[self.i][0] & 1   # 断在字符中间：新段开头是一个标志字节
                self.p = 1
            w = 2 if high else 1
            avail = (len(self.parts[self.i]) - self.p) // w
            k = min(n, avail)
            b = self.parts[self.i][self.p:self.p + k * w]
            out.append(b.decode("utf-16-le") if high else b.decode("latin-1"))
            self.p += k * w
            n -= k
        return "".join(out)


def read_xls_first_sheet(data):
    wb = _ole_stream(data)
    recs, o = [], 0
    while o + 4 <= len(wb):
        t, n = struct.unpack_from("<HH", wb, o)
        recs.append((t, wb[o + 4:o + 4 + n]))
        o += 4 + n
    sst, cells, bof, i = [], {}, 0, 0
    while i < len(recs):
        t, b = recs[i]
        if t == 0x00FC:  # SST
            parts = [b[8:]]
            j = i + 1
            while j < len(recs) and recs[j][0] == 0x003C:
                parts.append(recs[j][1])
                j += 1
            total = struct.unpack_from("<I", b, 4)[0]
            seg = _Seg(parts)
            for _ in range(total):
                cch, flag = struct.unpack("<HB", seg.raw(3))
                runs = struct.unpack("<H", seg.raw(2))[0] if flag & 8 else 0
                ext = struct.unpack("<I", seg.raw(4))[0] if flag & 4 else 0
                sst.append(seg.chars(cch, flag & 1))
                if runs:
                    seg.raw(4 * runs)
                if ext:
                    seg.raw(ext)
            i = j
            continue
        if t == 0x0809:
            bof += 1
        elif t == 0x000A and bof >= 2:
            break  # 第一张工作表结束
        elif bof >= 2:
            if t == 0x00FD:
                r, c, _, k = struct.unpack_from("<HHHI", b)
                cells[(r, c)] = sst[k]
            elif t == 0x0203:
                r, c, _, v = struct.unpack_from("<HHHd", b)
                cells[(r, c)] = v
            elif t == 0x027E:
                r, c, _, v = struct.unpack_from("<HHHI", b)
                cells[(r, c)] = _rk(v)
            elif t == 0x00BD:
                r, c0 = struct.unpack_from("<HH", b)
                n = (len(b) - 6) // 6
                for k in range(n):
                    _, v = struct.unpack_from("<HI", b, 4 + 6 * k)
                    cells[(r, c0 + k)] = _rk(v)
            elif t == 0x0006:
                r, c, _ = struct.unpack_from("<HHH", b)
                if b[12:14] != b"\xff\xff":
                    cells[(r, c)] = struct.unpack_from("<d", b, 6)[0]
            elif t == 0x0204:
                r, c, _, cch, flag = struct.unpack_from("<HHHHB", b)
                raw = b[9:]
                cells[(r, c)] = raw[:cch * 2].decode("utf-16-le") if flag & 1 else raw[:cch].decode("latin-1")
        i += 1
    return cells


# ═══════════════════════════ 价格 ═══════════════════════════
def _get(url, timeout=15):
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=timeout) as r:
        return r.read()


def fetch_gov_prices(max_pages=6):
    """返回 {"date": "2026-09-30", "url": 文章地址, "items": {"品种|规格": 均价}}。"""
    art = None
    for k in range(1, max_pages + 1):
        html = _get(GOV_LIST.format("" if k == 1 else "_%d" % k)).decode("utf-8", "ignore")
        for href, title in re.findall(r'href="(/nw17239/\d{8}/[0-9a-f]+\.html)"[^>]*title="([^"]+)"', html):
            if title.startswith(GOV_TITLE):
                art = (href, title)
                break
        if art:
            break
    if not art:
        raise RuntimeError("市政府网站最近几页没找到价格信息表")
    page = _get(GOV_BASE + art[0]).decode("utf-8", "ignore")
    m = re.search(r'href="([^"]+\.xls)"', page)
    if not m:
        raise RuntimeError("价格信息表页面里没有 xls 附件")
    cells = read_xls_first_sheet(_get(GOV_BASE + m.group(1)))
    rows = {}
    for (r, c), v in cells.items():
        if c == 0 and isinstance(v, str) and v.strip() in ("品种", "规格", "单位", "均价"):
            rows[v.strip()] = r
    if not {"品种", "均价"} <= rows.keys():
        raise RuntimeError("价格信息表格式变了，找不到「品种」「均价」两行")
    maxc = max(c for _, c in cells)
    items = {}
    for c in range(2, maxc + 1):
        name = str(cells.get((rows["品种"], c), "")).strip()
        spec = str(cells.get((rows.get("规格", -1), c), "")).strip()
        val = cells.get((rows["均价"], c))
        if name and isinstance(val, float) and val > 0:
            items.setdefault(name + "|" + spec, round(val, 2))
    dm = re.search(r"(\d{4})年(\d{2})月(\d{2})日", art[1])
    date = "-".join(dm.groups()) if dm else ""
    return {"date": date, "url": GOV_BASE + art[0], "items": items}


def load_prices(offline=False):
    os.makedirs(CACHE_DIR, exist_ok=True)
    path = os.path.join(CACHE_DIR, "gov-prices.json")
    cached = None
    try:
        with open(path, encoding="utf-8") as f:
            cached = json.load(f)
    except (OSError, ValueError):
        pass
    if cached and (offline or now().timestamp() - cached.get("fetched", 0) < 6 * 3600):
        return cached, ""
    if offline:
        return None, "离线模式，用内置兜底价"
    try:
        p = fetch_gov_prices()
        p["fetched"] = now().timestamp()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(p, f, ensure_ascii=False)
        return p, ""
    except Exception as exc:  # 抓不到就用旧缓存，再不行用兜底价
        if cached:
            return cached, "这次没抓到最新价（%s），用的是 %s 那期" % (exc, cached.get("date"))
        return None, "没抓到市政府价格表（%s），用内置兜底价" % exc


def gov_lookup(gov, name, spec):
    if not gov:
        return None
    items = gov["items"]
    if spec:
        return items.get(name + "|" + spec)
    for k, v in items.items():
        if k.split("|")[0] == name:
            return v
    return None


def build_price_table(gov):
    """{食材: (单位, 单价, 来源)}，另加 _大米 每斤价 / _米袋价 / _油每升 / _水果。"""
    date = gov["date"] if gov else ""
    src_gov = "上海市政府价格信息表 %s 全市均价" % date
    out = {}
    for ing, (unit, name, spec, factor, fallback, _) in M.ING.items():
        v = gov_lookup(gov, name, spec) if name else None
        if v is not None:
            out[ing] = (unit, round(v * factor, 2), src_gov + ("（按8个/斤折算）" if unit == "个" else ""))
        else:
            out[ing] = (unit, fallback, "估算（政府表没有这一项）" if not name else "兜底价（这次没取到政府价）")
    rice_bag = gov_lookup(gov, "粳米", "10千克袋装") or M.STAPLE["大米"][2]
    rice_jin = gov_lookup(gov, "粳米", "特二散装") or round(rice_bag / 20, 2)
    oil5 = gov_lookup(gov, "大豆油", None) or M.STAPLE["食用油"][2]
    fruit = gov_lookup(gov, M.FRUIT[0], None) or M.FRUIT[1]
    out["_米袋"] = rice_bag
    out["_米斤"] = rice_jin
    out["_油升"] = round(oil5 / 5, 2)
    out["_水果"] = fruit
    return out


PREMIUM = {"牛腩", "牛腱子", "牛里脊", "羊肉", "带鱼", "黄鱼", "基围虾", "排骨"}

# ═══════════════════════════ 口味匹配 ═══════════════════════════
SINGLE = {"鸡": "鸡肉", "牛": "牛肉", "羊": "羊肉", "猪": "猪肉", "鱼": "鱼", "虾": "虾", "蛋": "蛋"}
SPICY_WORDS = {"辣", "辣的", "吃辣", "辣味", "麻辣", "重辣"}


def split_words(s):
    out = []
    for w in re.split(r"[,，、;；\s]+", s or ""):
        w = re.sub(r"^(不吃|不爱吃|爱吃|喜欢吃|喜欢|不要|忌|少吃)", "", w.strip())
        if w:
            out.append(SINGLE.get(w, w))
    return out


def ing_words(ing):
    unit, name, spec, factor, fb, alias = M.ING[ing]
    return {ing} | set(alias.split())


def dish_hits(d, word):
    """这道菜是否命中某个口味词（菜名、食材名、食材别名）。"""
    name, typ, tags, ings = d[0], d[1], d[2], d[3]
    if word in SPICY_WORDS:
        return "spicy" in tags
    if len(word) >= 2 and word in name:
        return True
    return any(word in ing_words(i) for i in ings)


# ═══════════════════════════ 排菜 ═══════════════════════════
TIERS = [  # (每人每天菜钱下限, {餐: [槽位类型]})
    (0, {"午餐": ["main"], "晚餐": ["main"]}),
    (10, {"午餐": ["main", "veg"], "晚餐": ["main", "veg"]}),
    (22, {"午餐": ["main", "veg"], "晚餐": ["main", "side", "veg"]}),
    (38, {"午餐": ["main", "side", "veg"], "晚餐": ["main", "side", "veg", "soup"]}),
    (50, {"午餐": ["main", "side", "veg"], "晚餐": ["main", "main", "veg", "soup"]}),
]
TREAT_QUOTA = [(0, 1), (14, 2), (22, 3), (30, 4), (40, 5)]  # 每人每天菜钱 -> 每周几顿改善
TREAT_ORDER = [5, 6, 2, 4, 0, 3, 1]  # 优先排在周六、周日、周三……
TIER_DESC = ["每顿一个荤菜（荤素一锅出）", "每顿一荤一素", "午餐一荤一素，晚餐一荤一小荤一素",
             "午餐一荤一小荤一素，晚餐再加一个汤", "午餐一荤一小荤一素，晚餐两荤一素一汤"]


class Planner:
    def __init__(self, a, prices):
        self.a, self.P = a, prices
        self.likes, self.dislikes = split_words(a.like), split_words(a.dislike)
        if a.no_spicy and "辣" not in self.dislikes:
            self.dislikes.append("辣")
        self.unmatched_dislike = []
        dishes = list(M.D)
        self.cond_notes = {}
        for w in self.dislikes:
            hit = [d for d in dishes if dish_hits(d, w)]
            if hit:
                dishes = [d for d in dishes if d not in hit]
            elif any(w in d[5] for d in dishes):
                for d in dishes:
                    if w in d[5]:
                        self.cond_notes.setdefault(d[0], []).append(w)
            else:
                self.unmatched_dislike.append(w)
        self.unmatched_like = [w for w in self.likes if not any(dish_hits(d, w) for d in dishes)]
        self.pool = {t: [d for d in dishes if d[1] == t] for t in ("main", "side", "veg", "soup", "breakfast")}
        self.by_name = {d[0]: d for d in M.D}
        self.cost = {d[0]: sum(prices[i][1] * q for i, q in d[3].items()) for d in M.D}
        self.weight = {}
        for d in dishes:
            w = 1.0
            for word in self.likes:
                if dish_hits(d, word):
                    w *= 4
            self.weight[d[0]] = w
        # 爱吃的里有牛羊肉/鱼虾/排骨时，改善的钱优先花在它上面，别的贵菜少排
        self.likes_premium = any(self.weight[d[0]] > 1 and any(i in PREMIUM for i in d[3]) for d in dishes)
        if self.likes_premium:
            for d in dishes:
                if self.weight[d[0]] == 1 and any(i in PREMIUM for i in d[3]):
                    self.weight[d[0]] = 0.25
        seed = a.seed if a.seed is not None else int(hashlib.sha256(
            json.dumps([a.budget, a.days, a.like, a.dislike, a.meals, a.people, str(a.start)]).encode()).hexdigest()[:8], 16)
        self.rng = random.Random(seed)
        self.locked = set()
        prem = [d for d in self.pool["main"] if "stew" not in d[2] or True]
        prem = [d for d in prem if any(i in PREMIUM for i in d[3])]
        liked = [d for d in prem if self.weight.get(d[0], 1) > 1]
        self.premium = liked or prem

    # 每个槽位：(day, meal, slot_type)
    def meals_of(self, tier):
        st = dict(TIERS[tier][1])
        if self.a.meals == 1:
            st = {"晚餐": st["晚餐"]}
        out = {}
        if self.a.meals >= 3:
            out["早餐"] = ["breakfast"]
        out.update(st)
        return out

    def valid(self, d, plan, day, meal, slot_i):
        name, typ, tags, ings = d[0], d[1], d[2], d[3]
        if "stew" in tags and meal != "晚餐":
            return False
        window = 2 if typ == "breakfast" else 4
        for dd in range(max(0, day - window), min(len(plan), day + window + 1)):
            for m, lst in plan[dd].items():
                for k, x in enumerate(lst):
                    if x == name and not (dd == day and m == meal and k == slot_i):
                        return False
        # 同一餐里别有两个菜主料相同
        main_ing = max(ings, key=ings.get) if ings else None
        for k, x in enumerate(plan[day].get(meal, [])):
            if k != slot_i and x and main_ing:
                o = self.by_name[x][3]
                if o and max(o, key=o.get) == main_ing:
                    return False
        return True

    def pick(self, typ, plan, day, meal, slot_i, key=None):
        cand = [d for d in self.pool[typ] if self.valid(d, plan, day, meal, slot_i)]
        if not cand:
            cand = [d for d in self.pool[typ] if "stew" not in d[2] or meal == "晚餐"] or self.pool[typ]
        if key:
            cand = [d for d in cand if key(d)]
            if not cand:
                return None
        ws = [self.weight.get(d[0], 1) for d in cand]
        return self.rng.choices(cand, ws)[0][0]

    def fill(self, tier):
        layout = self.meals_of(tier)
        plan = [{m: [None] * len(s) for m, s in layout.items()} for _ in range(self.a.days)]
        for day in range(self.a.days):
            for m, slots in layout.items():
                for k, typ in enumerate(slots):
                    if not self.pool[typ]:
                        plan[day][m][k] = None
                        continue
                    plan[day][m][k] = self.pick(typ, plan, day, m, k)
        for day in plan:
            for m in day:
                day[m] = [x for x in day[m] if x]
        return plan, layout

    def add_treats(self, plan, per_week, start):
        """每周固定几顿晚餐吃改善菜（爱吃的牛羊肉/鱼虾/排骨优先），锁住不让压预算时换掉。"""
        self.locked = set()
        if not self.premium or per_week <= 0:
            return plan
        for w0 in range(0, len(plan), 7):
            days = list(range(w0, min(w0 + 7, len(plan))))
            days.sort(key=lambda i: TREAT_ORDER.index((start + dt.timedelta(days=i)).weekday()))
            n = max(1, round(per_week * len(days) / 7)) if len(days) >= 3 else (1 if per_week >= 2 else 0)
            for i in days[:n]:
                meal = "晚餐" if "晚餐" in plan[i] else next(iter(plan[i]))
                k = 0
                cand = [d for d in self.premium if self.valid(d, plan, i, meal, k)] or self.premium
                cand = [d for d in cand if "stew" not in d[2] or meal == "晚餐"] or cand
                plan[i][meal][k] = self.rng.choices(cand, [self.weight.get(d[0], 1) for d in cand])[0][0]
                self.locked.add((i, meal, k))
        return plan

    def food_cost(self, plan):
        return sum(self.cost[x] for day in plan for lst in day.values() for x in lst) * self.a.people

    def adjust(self, plan, target, slack):
        """先降到目标以内，再在不超目标的前提下往贵的、爱吃的换。"""
        slots = [(d, m, k) for d in range(len(plan)) for m in plan[d] for k in range(len(plan[d][m]))]
        ppl = self.a.people
        for _ in range(len(slots) * 3):
            c = self.food_cost(plan)
            if c <= target:
                break
            ranked = sorted(slots, key=lambda s: (s in self.locked, -self.cost[plan[s[0]][s[1]][s[2]]]))
            done = False
            for d, m, k in ranked[:12] + [x for x in ranked if x in self.locked]:
                cur = plan[d][m][k]
                typ = self.by_name[cur][1]
                cheaper = [x for x in self.pool[typ] if self.cost[x[0]] < self.cost[cur] - 0.3
                           and self.valid(x, plan, d, m, k)]
                if cheaper:
                    cheaper.sort(key=lambda x: self.cost[x[0]])
                    top = cheaper[:max(3, len(cheaper) // 3)]
                    plan[d][m][k] = self.rng.choices(top, [self.weight.get(x[0], 1) for x in top])[0][0]
                    done = True
                    break
            if not done:
                break
        for _ in range(len(slots) * 2):
            room = target - self.food_cost(plan)
            if room <= slack:
                break
            d, m, k = self.rng.choice(slots)
            if (d, m, k) in self.locked:
                continue
            cur = plan[d][m][k]
            typ = self.by_name[cur][1]
            up = [x for x in self.pool[typ]
                  if 0.5 < (self.cost[x[0]] - self.cost[cur]) * ppl <= room and self.valid(x, plan, d, m, k)]
            if up:
                ws = [self.weight.get(x[0], 1) * (1 + self.cost[x[0]] / 10) for x in up]
                plan[d][m][k] = self.rng.choices(up, ws)[0][0]
        return plan


# ═══════════════════════════ 固定开销 ═══════════════════════════
def parse_amount(s):
    m = re.match(r"(半|\d+(?:\.\d+)?)", s)
    if not m:
        return 0.0
    return 0.5 if m.group(1) == "半" else float(m.group(1))


def seasoning_needs(names):
    need, oil_spoons, aromatics = set(), 0.0, 0
    for n in names:
        for item in re.split(r"[，,]", M_BY[n][5]):
            item = item.strip()
            if not item or "可不放" in item:
                continue
            head = item.split()[0] if " " in item else item
            if head == "油":
                oil_spoons += parse_amount(item.split()[1]) if " " in item and "汤勺" in item else 0.3
            elif head in ("葱", "姜", "蒜"):
                aromatics += 1
            for kw in list(M.SEASON_BASE) + list(M.SEASON_EXTRA):
                if head == kw or head.startswith(kw):
                    need.add(kw)
    return need, oil_spoons, aromatics


M_BY = {d[0]: d for d in M.D}


def fixed_costs(plan, a, P):
    names = [x for day in plan for lst in day.values() for x in lst]
    rows = []
    if a.pantry:
        return rows, set()
    ppl = a.people
    rice_meals = sum(1 for day in plan for m, lst in day.items() if m != "早餐" and lst) * ppl
    rice_meals += sum(0.5 for x in names if x == "白粥鸡蛋") * ppl
    rice = rice_meals * CUP_RICE_JIN
    if rice > 0:
        bulk = math.ceil(rice) * P["_米斤"]
        if bulk <= math.ceil(rice / 20) * P["_米袋"]:
            rows.append(("大米", "%d斤（散装称重）" % math.ceil(rice), round(bulk, 1),
                         "约 %.0f 斤够吃；政府价表散装米均价" % rice))
        else:
            bags = math.ceil(rice / 20)
            rows.append(("大米", "%d袋（10公斤装）" % bags, round(bags * P["_米袋"], 1),
                         "约 %.0f 斤够吃；政府价表10公斤袋装均价" % rice))
    need, spoons, arom = seasoning_needs(names)
    litres = spoons * ppl * SPOON_ML / 1000
    if litres > 0:
        small = math.ceil(litres / 1.8)
        c_small = small * 1.8 * P["_油升"] * 1.15
        n5 = math.ceil(litres / 5)
        c_big = n5 * 5 * P["_油升"]
        if c_small <= c_big:
            rows.append(("食用油", "%d桶（1.8升装）" % small, round(c_small, 1), "约用 %.1f 升；按政府价表5升装均价折算" % litres))
        else:
            rows.append(("食用油", "%d桶（5升装）" % n5, round(c_big, 1), "约用 %.1f 升；政府价表5升装均价" % litres))
    base = [k for k in M.SEASON_BASE]
    extra = sorted(k for k in need if k in M.SEASON_EXTRA)
    sc = sum(M.SEASON_BASE.values()) + sum(M.SEASON_EXTRA[k] for k in extra)
    rows.append(("调料", "、".join(base + extra), round(sc, 1), "第一次买齐，下个月基本不用再买；价格是估算"))
    if arom:
        rows.append(("葱姜蒜", "随用随买", round(max(5, arom * 0.1 * (1 + 0.5 * (ppl - 1))), 1), "估算"))
    return rows, need


# ═══════════════════════════ xlsx 写入器（标准库） ═══════════════════════════
FONT = "微软雅黑"


class Styles:
    def __init__(self):
        self.fonts = ['<font><sz val="11"/><name val="%s"/></font>' % FONT]
        self.fills = ['<fill><patternFill patternType="none"/></fill>',
                      '<fill><patternFill patternType="gray125"/></fill>']
        self.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>',
                        '<border><left style="thin"><color rgb="FFC8C8C8"/></left><right style="thin"><color rgb="FFC8C8C8"/></right>'
                        '<top style="thin"><color rgb="FFC8C8C8"/></top><bottom style="thin"><color rgb="FFC8C8C8"/></bottom><diagonal/></border>',
                        '<border><left style="thin"><color rgb="FFC8C8C8"/></left><right style="thin"><color rgb="FFC8C8C8"/></right>'
                        '<top style="thin"><color rgb="FFC8C8C8"/></top><bottom style="medium"><color rgb="FF4F7942"/></bottom><diagonal/></border>']
        self.numfmts = {"0.0": 164, "0.00": 165}
        self.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>']
        self.cache = {}
        self.fcache, self.flcache = {}, {}

    def font(self, bold=False, color=None, size=11):
        k = (bold, color, size)
        if k not in self.fcache:
            self.fcache[k] = len(self.fonts)
            self.fonts.append('<font>%s<sz val="%s"/>%s<name val="%s"/></font>' % (
                "<b/>" if bold else "", size, '<color rgb="FF%s"/>' % color if color else "", FONT))
        return self.fcache[k]

    def fill(self, rgb):
        if not rgb:
            return 0
        if rgb not in self.flcache:
            self.flcache[rgb] = len(self.fills)
            self.fills.append('<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/><bgColor indexed="64"/></patternFill></fill>' % rgb)
        return self.flcache[rgb]

    def get(self, bold=False, color=None, size=11, fill=None, border=1, fmt=None, align="center", wrap=True):
        k = (bold, color, size, fill, border, fmt, align, wrap)
        if k not in self.cache:
            fid, flid = self.font(bold, color, size), self.fill(fill)
            nf = self.numfmts.get(fmt, 0)
            al = '<alignment horizontal="%s" vertical="center"%s/>' % (align, ' wrapText="1"' if wrap else "")
            self.cache[k] = len(self.xfs)
            self.xfs.append('<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyFont="1" applyFill="1" '
                            'applyBorder="1" applyAlignment="1"%s>%s</xf>' % (nf, fid, flid, border,
                                                                              ' applyNumberFormat="1"' if nf else "", al))
        return self.cache[k]

    def xml(self):
        nf = "".join('<numFmt numFmtId="%d" formatCode="%s"/>' % (i, escape(c)) for c, i in self.numfmts.items())
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<numFmts count="%d">%s</numFmts>' % (len(self.numfmts), nf) +
                '<fonts count="%d">%s</fonts>' % (len(self.fonts), "".join(self.fonts)) +
                '<fills count="%d">%s</fills>' % (len(self.fills), "".join(self.fills)) +
                '<borders count="%d">%s</borders>' % (len(self.borders), "".join(self.borders)) +
                '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                '<cellXfs count="%d">%s</cellXfs>' % (len(self.xfs), "".join(self.xfs)) +
                '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
                '</styleSheet>')


def col_letter(c):
    s = ""
    while c:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s


_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")


class Sheet:
    def __init__(self, name):
        self.name, self.cells, self.merges, self.widths, self.heights, self.freeze = name, {}, [], {}, {}, None

    def set(self, r, c, v, s=0, f=None):
        self.cells[(r, c)] = (v, s, f)

    def merge(self, r1, c1, r2, c2):
        if (r1, c1) != (r2, c2):
            self.merges.append("%s%d:%s%d" % (col_letter(c1), r1, col_letter(c2), r2))

    def xml(self):
        rows = {}
        for (r, c), v in self.cells.items():
            rows.setdefault(r, []).append((c, v))
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
               'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
        out.append("<sheetViews><sheetView workbookViewId=\"0\">")
        if self.freeze:
            fr, fc = self.freeze
            attrs = ""
            if fc > 1:
                attrs += ' xSplit="%d"' % (fc - 1)
            if fr > 1:
                attrs += ' ySplit="%d"' % (fr - 1)
            pane = "bottomRight" if fr > 1 and fc > 1 else ("bottomLeft" if fr > 1 else "topRight")
            out.append('<pane%s topLeftCell="%s%d" activePane="%s" state="frozen"/>' % (
                attrs, col_letter(fc), fr, pane))
        out.append("</sheetView></sheetViews><sheetFormatPr defaultRowHeight=\"16.5\"/>")
        if self.widths:
            out.append("<cols>")
            for c in sorted(self.widths):
                out.append('<col min="%d" max="%d" width="%s" customWidth="1"/>' % (c, c, self.widths[c]))
            out.append("</cols>")
        out.append("<sheetData>")
        for r in sorted(set(rows) | set(self.heights)):
            ht = ' ht="%s" customHeight="1"' % self.heights[r] if r in self.heights else ""
            out.append('<row r="%d"%s>' % (r, ht))
            for c, (v, s, f) in sorted(rows.get(r, [])):
                ref = "%s%d" % (col_letter(c), r)
                if f is not None:
                    out.append('<c r="%s" s="%d"><f>%s</f><v>%s</v></c>' % (ref, s, escape(f), repr(float(v or 0))))
                elif v is None or v == "":
                    out.append('<c r="%s" s="%d"/>' % (ref, s))
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    out.append('<c r="%s" s="%d"><v>%s</v></c>' % (ref, s, repr(float(v)) if isinstance(v, float) else v))
                else:
                    t = escape(_BAD.sub("", str(v)))
                    out.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, t))
            out.append("</row>")
        out.append("</sheetData>")
        if self.merges:
            out.append('<mergeCells count="%d">%s</mergeCells>' % (
                len(self.merges), "".join('<mergeCell ref="%s"/>' % m for m in self.merges)))
        out.append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/></worksheet>')
        return "".join(out)


def save_xlsx(path, sheets, styles):
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          + "".join('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1)
                    for i in range(len(sheets))) + "</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>')
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          '<bookViews><workbookView activeTab="0"/></bookViews><sheets>'
          + "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(s.name), i + 1, i + 1)
                    for i, s in enumerate(sheets)) + '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
    wbrels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              + "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1)
                        for i in range(len(sheets)))
              + '<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % (len(sheets) + 1)
              + "</Relationships>")
    tmp = path + ".part"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wbrels)
        z.writestr("xl/styles.xml", styles.xml())
        for i, s in enumerate(sheets):
            z.writestr("xl/worksheets/sheet%d.xml" % (i + 1), s.xml())
    os.replace(tmp, path)


# ═══════════════════════════ 组装工作簿 ═══════════════════════════
GREEN, WKD, TREAT, TOTAL, GRAY = "4F7942", "FFF4E5", "FDE2E2", "E8F0E3", "808080"


def fit_height(texts_widths):
    lines = 1
    for t, w in texts_widths:
        if not isinstance(t, str):
            continue
        cpl = max(1, int(w / 2.05))
        lines = max(lines, sum(max(1, math.ceil(len(x) / cpl)) for x in t.split("\n")))
    return 16.5 * lines + 4


def is_treat(day):
    return any(i in PREMIUM for lst in day.values() for x in lst
               if M_BY[x][1] == "main" for i in M_BY[x][3])


def thaw_note(plan, i):
    if i + 1 >= len(plan):
        return "最后一天，记得盘点剩下的食材"
    meats, fish = [], []
    for lst in plan[i + 1].values():
        for x in lst:
            for ing in M_BY[x][3]:
                short = M.MEAT_SHORT.get(ing)
                if not short:
                    continue
                target = fish if ing in ("带鱼", "黄鱼", "草鱼", "鲫鱼", "花鲢鱼头", "基围虾") else meats
                if short not in target:
                    target.append(short)
    parts = []
    if meats:
        parts.append("把明天要用的%s从冷冻室拿到冷藏室解冻" % "、".join(meats))
    if fish:
        parts.append("明天有%s，最好当天买新鲜的" % "、".join(fish))
    return "；".join(parts)


def cond_text(name, notes):
    t = M_BY[name][5]
    if name in notes:
        t += "\n（不吃%s就不放）" % "、".join(notes[name])
    return t


def build_workbook(path, a, pl, plan, P, gov, warn, fixed, totals):
    st = Styles()
    H = st.get(bold=True, color="FFFFFF", fill=GREEN)
    N = st.get()
    NL = st.get(align="left")
    B = st.get(bold=True)
    sheets = []
    dates = [a.start + dt.timedelta(days=i) for i in range(a.days)]
    ppl_note = "（一人份）" if a.people == 1 else "（一人份，%d个人就乘%d）" % (a.people, a.people)

    # ── 说明 ──
    s0 = Sheet("说明")
    s0.widths = {1: 100}
    food, fixed_sum, fruit, spare = totals["food"], totals["fixed"], totals["fruit"], totals["spare"]
    total = food + fixed_sum + fruit
    lines = [
        ("%s家常菜谱（%s–%s，共%d天）" % (a.name + "的" if a.name else "", dates[0].strftime("%-m月%-d日"),
                                    dates[-1].strftime("%-m月%-d日"), a.days), st.get(bold=True, size=16, border=0, align="left")),
        ("", None),
        ("这份菜谱的条件", "b"),
        ("• 预算 %s，%d 个人，每天%s。" % (a.budget_text, a.people, {1: "只算晚餐", 2: "午餐和晚餐", 3: "早中晚三餐"}[a.meals]), None),
        ("• 爱吃：%s；不吃：%s。" % ("、".join(pl.likes) or "没说", "、".join(pl.dislikes) or "没说"), None),
        ("• 结构：%s。" % TIER_DESC[totals["tier"]] + ("早餐另算。" if a.meals >= 3 else ""), None),
        ("• 预计花费 %.0f 元（菜钱 %.0f + %s %.0f%s），结余 %.0f 元，平均每天 %.1f 元。" % (
            total, food, "米油调料" if not a.pantry else "固定开销", fixed_sum,
            "，水果 %.0f" % fruit if fruit else "", a.budget - total, total / a.days), None),
    ]
    if a.pantry:
        lines.append(("• 按「家里已经有米、油、调料」算的，没有计入这几项。", None))
    if getattr(a, "shrunk_note", ""):
        lines.append(("• ⚠ " + a.shrunk_note + "。", None))
    if a.shrunk:
        lines.append(("• ⚠ 预算不够排满：原本要 %d 顿，这些钱实际只够 %d 顿（最便宜的一顿约 %.1f 元），所以只排了 %d 顿。" % (
            a.shrunk[0], a.shrunk[1], a.shrunk[2], a.shrunk[1]), None))
    if total > a.budget:
        lines.append(("• ⚠ 预算太紧：已经把菜换到最便宜了还是超出 %.0f 元。可以加预算，或者改成每顿一个菜。" % (total - a.budget), None))
    lines += [
        ("", None),
        ("怎么看", "b"),
        ("• 「每日做饭」：每天每顿吃什么、用多少料、放什么调料、怎么做。最右一列是当晚要提前准备的事。", None),
        ("• 「菜谱总览」：一天一行，带当天菜钱。浅红色是吃牛羊肉、鱼虾、排骨的改善日，浅橙色是周末。", None),
        ("• 「采购清单」：按周列出要买多少，最后算出总花费和结余。", None),
        ("• 「做法大全」：这份菜谱里所有菜的做法；「单价表」：每样食材的单价和出处。", None),
        ("", None),
        ("用量怎么量", "b"),
        ("• 汤勺 = 喝汤的瓷勺；小勺 = 盐罐里配的小勺；少许 = 用手指捏一小撮。", None),
        ("• 麻将块 = 切成麻将牌大小的肉块；「半个巴掌大一块」指一指厚的肉。", None),
        ("• 米饭：每顿每人 1 杯米（电饭煲配的量杯），水加到内胆对应刻度，按煮饭键，大约 40 分钟。", None),
        ("• 肉买回来按一顿的量分袋冷冻，到时候拿一袋就是一顿。鱼虾最好当天买。", None),
        ("• 盐的量偏淡，出锅前尝一下，不够再补。", None),
        ("", None),
        ("价格", "b"),
        ("• " + ("上海市政府「主要主副食品品种价格信息表」%s 那一期的全市均价（60 多个菜场和超市实测）。" % gov["date"]
                 if gov else "这次没取到政府价格，用的是内置兜底价。"), None),
        ("• 鸡胸、鸡腿、豆腐、皮蛋这几样政府表里没有，用的是估算价，单价表里标了「估算」。", None),
        ("• 盒马、叮咚的价格和菜场均价会有差别，以 App 里为准。没有算燃气费和水电费。", None),
    ]
    if warn:
        lines.append(("• 注：" + warn, None))
    for i, (t, f) in enumerate(lines, 1):
        s0.set(i, 1, t, f if isinstance(f, int) else st.get(bold=(f == "b"), border=0, align="left"))
    sheets.append(s0)

    # ── 每日做饭 ──
    s1 = Sheet("每日做饭")
    W = {1: 6, 2: 9, 3: 6, 4: 12, 5: 24, 6: 28, 7: 58, 8: 18}
    s1.widths = W
    for c, h in enumerate(["天", "日期", "餐", "菜名", "用料" + ppl_note, "调料", "做法", "今晚提前准备"], 1):
        s1.set(1, c, h, H)
    r = 2
    for i, day in enumerate(plan):
        d = dates[i]
        fill = TREAT if is_treat(day) else (WKD if d.weekday() >= 5 else None)
        last = None
        start = r
        for meal, lst in day.items():
            mstart = r
            for x in lst:
                dish = M_BY[x]
                s1.set(r, 4, x, st.get(bold=True, fill=fill))
                s1.set(r, 5, dish[4], st.get(fill=fill, align="left"))
                s1.set(r, 6, cond_text(x, pl.cond_notes), st.get(fill=fill, align="left"))
                s1.set(r, 7, dish[6], st.get(fill=fill, align="left"))
                s1.heights[r] = fit_height([(dish[4], W[5]), (cond_text(x, pl.cond_notes), W[6]), (dish[6], W[7])])
                last = r
                r += 1
            if r == mstart:
                continue
            s1.set(mstart, 3, meal, st.get(fill=fill))
            for rr in range(mstart + 1, r):
                s1.set(rr, 3, None, st.get(fill=fill))
            s1.merge(mstart, 3, r - 1, 3)
        if r == start:
            continue
        for c, v in ((1, i + 1), (2, "%d月%d日\n周%s" % (d.month, d.day, WEEK[d.weekday()])), (8, thaw_note(plan, i))):
            s1.set(start, c, v, st.get(bold=(c == 1), fill=fill, align="left" if c == 8 else "center"))
            for rr in range(start + 1, r):
                s1.set(rr, c, None, st.get(fill=fill))
            s1.merge(start, c, r - 1, c)
        for c in range(1, 9):  # 天与天之间的粗线
            v, s, f = s1.cells.get((last, c), (None, 0, None))
            key = [k for k, idx in st.cache.items() if idx == s]
            if key:
                kk = list(key[0])
                kk[4] = 2
                s1.cells[(last, c)] = (v, st.get(*kk), f)
    s1.freeze = (2, 4)
    sheets.append(s1)

    # ── 菜谱总览 ──
    s2 = Sheet("菜谱总览")
    meals = list(plan[0].keys())
    s2.widths = {1: 6, 2: 11, 3: 6, **{4 + k: 30 for k in range(len(meals))}, 4 + len(meals): 12}
    for c, h in enumerate(["天", "日期", "星期"] + meals + ["当天菜钱（元）"], 1):
        s2.set(1, c, h, H)
    for i, day in enumerate(plan):
        d = dates[i]
        fill = TREAT if is_treat(day) else (WKD if d.weekday() >= 5 else None)
        rr = i + 2
        s2.set(rr, 1, i + 1, st.get(fill=fill))
        s2.set(rr, 2, "%d月%d日" % (d.month, d.day), st.get(fill=fill))
        s2.set(rr, 3, "周" + WEEK[d.weekday()], st.get(fill=fill))
        for k, m in enumerate(meals):
            s2.set(rr, 4 + k, "、".join(day.get(m, [])), st.get(fill=fill))
        s2.set(rr, 4 + len(meals), sum(pl.cost[x] for lst in day.values() for x in lst) * a.people,
               st.get(fill=fill, fmt="0.0"))
    tr = len(plan) + 2
    cc = 4 + len(meals)
    s2.set(tr, 1, "合计", st.get(bold=True, fill=TOTAL))
    for c in range(2, cc):
        s2.set(tr, c, None, st.get(fill=TOTAL))
    s2.merge(tr, 1, tr, cc - 1)
    s2.set(tr, cc, food, st.get(bold=True, fill=TOTAL, fmt="0.0"),
           "SUM(%s2:%s%d)" % (col_letter(cc), col_letter(cc), tr - 1))
    s2.freeze = (2, 4)
    sheets.append(s2)

    # ── 采购清单 ──
    s3 = Sheet("采购清单")
    weeks = math.ceil(a.days / 7)
    qty = {}
    for i, day in enumerate(plan):
        w = i // 7
        for lst in day.values():
            for x in lst:
                for ing, q in M_BY[x][3].items():
                    qty.setdefault(ing, [0.0] * weeks)[w] += q * a.people
    order = list(M.ING)
    used = sorted(qty, key=order.index)
    heads = ["食材", "单位"] + ["第%d周" % (w + 1) for w in range(weeks)] + ["合计", "单价（元）", "金额（元）", "怎么买"]
    s3.widths = {1: 12, 2: 6, **{3 + w: 8 for w in range(weeks)}}
    cT, cP, cA, cH = 3 + weeks, 4 + weeks, 5 + weeks, 6 + weeks
    s3.widths.update({cT: 8, cP: 10, cA: 10, cH: 40})
    for c, h in enumerate(heads, 1):
        s3.set(1, c, h, H)
    r = 2
    for ing in used:
        unit, price, src = P[ing]
        s3.set(r, 1, ing, B)
        s3.set(r, 2, unit, N)
        for w in range(weeks):
            v = qty[ing][w]
            s3.set(r, 3 + w, v if v else None, st.get(fmt="0.0"))
        tot = sum(qty[ing])
        s3.set(r, cT, tot, st.get(fmt="0.0"), "SUM(%s%d:%s%d)" % (col_letter(3), r, col_letter(2 + weeks), r))
        s3.set(r, cP, price, st.get(color="0000FF", fmt="0.00"))
        s3.set(r, cA, tot * price, st.get(fmt="0.0"), "%s%d*%s%d" % (col_letter(cT), r, col_letter(cP), r))
        if ing in M.MEAT_SHORT and ing not in ("带鱼", "黄鱼", "草鱼", "鲫鱼", "花鲢鱼头", "基围虾"):
            how = "每周买一次，按一顿的量分袋冷冻"
        elif ing in ("带鱼", "黄鱼", "草鱼", "鲫鱼", "花鲢鱼头", "基围虾"):
            how = "吃的当天买新鲜的"
        elif ing in ("青菜", "鸡毛菜", "菠菜", "生菜", "空心菜", "米苋", "韭菜", "豆芽", "豆腐", "番茄", "黄瓜", "茄子", "丝瓜"):
            how = "放不久，每3–4天买一次"
        else:
            how = "耐放，每周买一次"
        s3.set(r, cH, how, NL)
        r += 1
    ing_end = r - 1
    s3.set(r, 1, "食材小计", st.get(bold=True, fill=TOTAL))
    for c in range(2, cA):
        s3.set(r, c, None, st.get(fill=TOTAL))
    s3.merge(r, 1, r, cA - 1)
    s3.set(r, cA, food, st.get(bold=True, fill=TOTAL, fmt="0.0"), "SUM(%s2:%s%d)" % (col_letter(cA), col_letter(cA), ing_end))
    s3.set(r, cH, None, st.get(fill=TOTAL))
    sub1 = r
    r += 2
    extra_rows = list(fixed)
    if fruit:
        extra_rows.append(("水果", "每周约 %.1f 斤苹果" % (fruit / P["_水果"] / weeks), fruit, "预算有富余，加点水果"))
    sub2 = None
    if extra_rows:
        s3.set(r, 1, "米油调料等", H)
        s3.set(r, 2, "买多少", H)
        for c in range(3, cA):
            s3.set(r, c, None, H)
        s3.merge(r, 2, r, cA - 1)
        s3.set(r, cA, "金额（元）", H)
        s3.set(r, cH, "说明", H)
        r += 1
        fs = r
        for n, q, amt, note in extra_rows:
            s3.set(r, 1, n, B)
            s3.set(r, 2, q, NL)
            for c in range(3, cA):
                s3.set(r, c, None, N)
            s3.merge(r, 2, r, cA - 1)
            s3.set(r, cA, amt, st.get(color="0000FF", fmt="0.0"))
            s3.set(r, cH, note, NL)
            r += 1
        s3.set(r, 1, "小计", st.get(bold=True, fill=TOTAL))
        for c in range(2, cA):
            s3.set(r, c, None, st.get(fill=TOTAL))
        s3.merge(r, 1, r, cA - 1)
        s3.set(r, cA, fixed_sum + fruit, st.get(bold=True, fill=TOTAL, fmt="0.0"),
               "SUM(%s%d:%s%d)" % (col_letter(cA), fs, col_letter(cA), r - 1))
        s3.set(r, cH, None, st.get(fill=TOTAL))
        sub2 = r
        r += 2
    big = st.get(bold=True, size=12)
    for label, val, f in (
            ("总花费", total, "%s%d" % (col_letter(cA), sub1) + ("+%s%d" % (col_letter(cA), sub2) if sub2 else "")),
            ("预算", float(a.budget), None),
            ("结余（负数是超支）", a.budget - total, "%s%d-%s%d" % (col_letter(cA), r + 1, col_letter(cA), r)),
            ("平均每天", total / a.days, "%s%d/%d" % (col_letter(cA), r, a.days))):
        s3.set(r, 1, label, big)
        for c in range(2, cA):
            s3.set(r, c, None, N)
        s3.merge(r, 1, r, cA - 1)
        s3.set(r, cA, val, st.get(bold=True, size=12, fmt="0.0",
                                            color="0000FF" if label == "预算" else ("FF0000" if val < 0 else None)), f)
        r += 1
    s3.freeze = (2, 2)
    sheets.append(s3)

    # ── 做法大全 ──
    s4 = Sheet("做法大全")
    s4.widths = {1: 13, 2: 24, 3: 30, 4: 62}
    for c, h in enumerate(["菜名", "用料" + ppl_note, "调料", "做法"], 1):
        s4.set(1, c, h, H)
    names = []
    for day in plan:
        for lst in day.values():
            for x in lst:
                if x not in names:
                    names.append(x)
    tord = ["breakfast", "main", "side", "veg", "soup"]
    names.sort(key=lambda x: (tord.index(M_BY[x][1]), [d[0] for d in M.D].index(x)))
    for i, x in enumerate(names, 2):
        d = M_BY[x]
        s4.set(i, 1, x, B)
        s4.set(i, 2, d[4], NL)
        s4.set(i, 3, cond_text(x, pl.cond_notes), NL)
        s4.set(i, 4, d[6], NL)
        s4.heights[i] = fit_height([(d[4], 24), (d[5], 30), (d[6], 62)])
    s4.freeze = (2, 2)
    sheets.append(s4)

    # ── 单价表 ──
    s5 = Sheet("单价表")
    s5.widths = {1: 12, 2: 6, 3: 10, 4: 56}
    for c, h in enumerate(["食材", "单位", "单价（元）", "来源"], 1):
        s5.set(1, c, h, H)
    for i, ing in enumerate(used, 2):
        unit, price, src = P[ing]
        s5.set(i, 1, ing, B)
        s5.set(i, 2, unit, N)
        s5.set(i, 3, price, st.get(color="0000FF", fmt="0.00"))
        s5.set(i, 4, src, NL)
    s5.freeze = (2, 1)
    sheets.append(s5)

    save_xlsx(path, sheets, st)


# ═══════════════════════════ 命令 ═══════════════════════════
def parse_date(s, default):
    if not s:
        return default
    t = now().date()
    if s in ("今天",):
        return t
    if s in ("明天",):
        return t + dt.timedelta(days=1)
    if s in ("后天",):
        return t + dt.timedelta(days=2)
    if s in ("昨天",):
        return t - dt.timedelta(days=1)
    m = re.fullmatch(r"(?:(\d{4})[-/.年])?(\d{1,2})[-/.月](\d{1,2})日?", s.strip())
    if not m:
        raise SystemExit("日期看不懂：%s（写成 2026-10-05、10-5 或 今天/明天）" % s)
    y = int(m.group(1) or t.year)
    return dt.date(y, int(m.group(2)), int(m.group(3)))


def cmd_plan(a):
    a.start = parse_date(a.start, now().date() + dt.timedelta(days=1))
    if a.until:
        end = parse_date(a.until, None)
        a.days = (end - a.start).days + 1
    if not a.days or a.days < 1 or a.days > 62:
        raise SystemExit("天数要在 1–62 之间（用 --days 或 --until）")
    if a.budget <= 0:
        raise SystemExit("预算要大于 0")
    a.budget_text = "%g 元" % a.budget
    a.budget_label = {"total": "%g元", "week": "每周%g元", "day": "每天%g元"}[a.per] % a.budget
    if a.per == "week":
        a.budget_text = "每周 %g 元（%d 天合计 %.0f 元）" % (a.budget, a.days, a.budget * a.days / 7)
        a.budget = round(a.budget * a.days / 7, 2)
    elif a.per == "day":
        a.budget_text = "每天 %g 元（%d 天合计 %.0f 元）" % (a.budget, a.days, a.budget * a.days)
        a.budget = round(a.budget * a.days, 2)
    gov, warn = load_prices(a.offline)
    P = build_price_table(gov)
    pl = Planner(a, P)
    if not pl.pool["main"]:
        raise SystemExit("忌口把所有荤菜都排除了，菜库里没有菜可排。")

    # ── 预算够不够：按最便宜的吃法估，够几顿就排几顿，不硬凑 ──
    a.shrunk = None
    rice_meal = CUP_RICE_JIN * P["_米斤"]
    lunch_ok = [d for d in pl.pool["main"] if "stew" not in d[2]]
    cheapest = min(pl.cost[d[0]] for d in (lunch_ok or pl.pool["main"]))
    min_meal = (cheapest + (0 if a.pantry else rice_meal + 0.4)) * a.people
    if a.meals >= 3 and pl.pool["breakfast"]:
        min_bf = min(pl.cost[d[0]] for d in pl.pool["breakfast"]) * a.people
    else:
        min_bf = min_meal
    min_fixed = 0 if a.pantry else sum(M.SEASON_BASE.values()) + 1.8 * P["_油升"] * 1.15
    meals_wanted = a.days * a.meals

    def cost_for(n_meals, fixed):
        return n_meals * min_meal * 1.05 + fixed

    if a.budget < cost_for(meals_wanted, min_fixed):
        cheap_name = min((lunch_ok or pl.pool["main"]), key=lambda d: pl.cost[d[0]])[0]

        def real_fixed(n_meals):   # 按 n 顿最便宜的菜实算一遍米油调料
            dummy = [{"午餐": [cheap_name]} for _ in range(n_meals)]
            return sum(x[2] for x in fixed_costs(dummy, a, P)[0])

        def fit_n():
            n = int(a.budget / (min_meal * 1.05))
            while n >= 1 and n * min_meal * 1.05 + real_fixed(n) > a.budget:
                n -= 1
            return n

        if not a.pantry:
            # 钱不够排满：米油盐酱醋不算进来（默认家里有），钱全用在买菜上
            a.pantry = True
            min_meal = cheapest * a.people
            a.shrunk_note = "钱比较少，米油盐酱醋没算进来（默认家里有），钱全用来买菜"
        n = fit_n()
        if n < 1:
            print("预算 %g 元连一顿饭都不够：最便宜的一顿（%s）%d 个人约 %.1f 元，这还没算米油盐。" % (
                a.budget, cheap_name, a.people, cheapest * a.people))
            print("没有生成菜谱，告诉用户需要加预算。")
            return
        if n < meals_wanted:
            a.shrunk = (meals_wanted, n, min_meal)
            a.meals = min(a.meals, n)
            a.days = max(1, n // a.meals)
    # 先按估计的固定开销定结构，再按真实固定开销微调
    est_fixed = 0 if a.pantry else 60 + a.days * a.people * 0.9
    per = (a.budget - est_fixed) / a.days / a.people
    tier = max([0] + [i for i, (lo, _) in enumerate(TIERS) if per >= lo])
    # 爱吃的是牛羊肉/鱼虾这类贵菜：每顿少一个菜，省下的钱多排爱吃的（用户定的：爱吃的多 > 菜的样数多）
    pl.tier_dropped = pl.likes_premium and 1 < tier < len(TIERS) - 1
    if pl.tier_dropped:
        tier -= 1
    if a.shrunk:
        est_fixed = 0 if a.pantry else min_fixed
        per = (a.budget - est_fixed) / a.days / a.people
        tier = 0
    best = None
    quota0 = max([0] + [q for lo, q in TREAT_QUOTA if per >= lo])
    tries = [(t, q) for t in range(tier, -1, -1) for q in range(quota0, -1, -1)]
    for tier, quota in tries:
        plan, layout = pl.fill(tier)
        plan = pl.add_treats(plan, quota, a.start)
        fixed, _ = fixed_costs(plan, a, P)
        fsum = sum(x[2] for x in fixed)
        target = a.budget * 0.985 - fsum
        plan = pl.adjust(plan, target, slack=a.budget * 0.01)
        food = pl.food_cost(plan)
        fixed, _ = fixed_costs(plan, a, P)   # 菜换过，调料可能变
        fsum = sum(x[2] for x in fixed)
        best = (plan, tier, food, fixed, fsum)
        if food + fsum <= a.budget:
            break
    plan, tier, food, fixed, fsum = best
    while food + fsum > a.budget and a.days > 1:   # 还是超：再少排一天
        a.days -= 1
        plan = plan[:a.days]
        food = pl.food_cost(plan)
        fixed, _ = fixed_costs(plan, a, P)
        fsum = sum(x[2] for x in fixed)
        if not a.shrunk:
            a.shrunk = (meals_wanted, None, min_meal)
    if a.shrunk:
        n_real = sum(1 for day in plan for lst in day.values() if lst)
        a.shrunk = (a.shrunk[0], n_real, a.shrunk[2])
    left = a.budget - food - fsum
    fruit = 0.0
    if left > a.budget * 0.05 and not a.shrunk:
        fruit = round(left * 0.6, 0)
    totals = {"food": food, "fixed": fsum, "fruit": fruit, "spare": left - fruit, "tier": tier}

    out_dir = a.out_dir or home_dir()
    os.makedirs(out_dir, exist_ok=True)
    stamp = now().strftime("%m%d-%H%M%S")
    fname = "%s菜谱-%s-%d天-%s.xlsx" % ((a.name + "的") if a.name else "", a.budget_label, a.days, stamp)
    fname = re.sub(r"[\s/\\:*?\"<>|]", "", fname)
    path = os.path.join(out_dir, fname)
    build_workbook(path, a, pl, plan, P, gov, warn, fixed, totals)

    dates = [a.start + dt.timedelta(days=i) for i in range(a.days)]
    state = {
        "created": now().isoformat(timespec="seconds"), "file": path, "budget": a.budget,
        "people": a.people, "meals": a.meals, "like": pl.likes, "dislike": pl.dislikes,
        "price_date": gov["date"] if gov else "", "cond_notes": pl.cond_notes,
        "days": [{"date": dates[i].isoformat(), "meals": plan[i]} for i in range(a.days)],
    }
    with open(os.path.join(home_dir(), "current.json"), "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=1)

    total = food + fsum + fruit
    names = [x for day in plan for lst in day.values() for x in lst]
    treat = sum(1 for day in plan if is_treat(day))
    print("已生成：%s" % path)
    if getattr(a, "shrunk_note", ""):
        print("⚠ " + a.shrunk_note)
    if a.shrunk:
        print("⚠ 预算不够排满：原本要 %d 顿，这些钱实际只够 %d 顿（最便宜的一顿约 %.1f 元），下面只排了这 %d 顿。" % (
            a.shrunk[0], a.shrunk[1], a.shrunk[2], a.shrunk[1]))
    print("发给用户时在回复里单独一行写：MEDIA:%s" % path)
    print("时间：%s 到 %s，共 %d 天，%d 人，%s" % (dates[0], dates[-1], a.days, a.people,
                                         {1: "只排晚餐", 2: "午餐+晚餐", 3: "早中晚三餐"}[a.meals]))
    print("结构：%s%s" % (TIER_DESC[tier], "（为了多排爱吃的，每顿少排了一个菜）" if pl.tier_dropped and not a.shrunk else ""))
    print("花费：菜钱 %.0f + 米油调料 %.0f%s = %.0f 元，预算 %s，结余 %.0f，平均每天 %.1f 元" % (
        food, fsum, " + 水果 %.0f" % fruit if fruit else "", total, a.budget_text, a.budget - total, total / a.days))
    if total > a.budget:
        print("⚠ 超预算 %.0f 元：已经全换成最便宜的菜了。建议加预算；家里有米油调料的话加 --pantry 不计这几项。" % (total - a.budget))
    print("用到 %d 道不同的菜；改善日（牛羊肉/鱼虾/排骨）%d 天" % (len(set(names)), treat))
    print("价格：%s" % ("上海市政府价格信息表 %s 全市均价" % gov["date"] if gov else "内置兜底价"))
    if warn:
        print("注意：" + warn)
    if pl.unmatched_dislike:
        print("这些忌口在菜库里没有对应的菜，不用排除：%s" % "、".join(pl.unmatched_dislike))
    for w in pl.likes:
        if w in pl.unmatched_like:
            continue
        hits = ["%s%s" % (dates[i].strftime("%-m/%-d"), m[0])
                for i, day in enumerate(plan) for m, lst in day.items()
                if any(dish_hits(M_BY[x], w) for x in lst)]
        print("爱吃的「%s」：%d 顿里排了 %d 顿%s" % (w, sum(len(day) for day in plan), len(hits),
                                           "（%s）" % "、".join(hits) if hits else ""))
    if pl.unmatched_like:
        print("这些爱吃的菜库里没有，没能多排：%s" % "、".join(pl.unmatched_like))
    if pl.cond_notes:
        print("忌口只出现在调料里的，菜照排、调料注明不放：%s" % "、".join(sorted({w for v in pl.cond_notes.values() for w in v})))
    print("第一天（%s）：%s" % (dates[0], "；".join("%s %s" % (m, "、".join(l)) for m, l in plan[0].items())))
    print("提醒用户：今晚" + (thaw_note([{}] + plan, 0) or "不用提前准备"))


def cmd_today(a):
    path = os.path.join(home_dir(), "current.json")
    try:
        with open(path, encoding="utf-8") as f:
            state = json.load(f)
    except (OSError, ValueError):
        print("这个微信号还没有生成过菜谱。先告诉我预算、天数、爱吃和不吃的，我来做一份。")
        return
    day = parse_date(a.date, now().date())
    days = {d["date"]: d for d in state["days"]}
    hit = days.get(day.isoformat())
    first, last = state["days"][0]["date"], state["days"][-1]["date"]
    if not hit:
        print("%s 不在当前菜谱里（菜谱是 %s 到 %s）。" % (day, first, last))
        if day.isoformat() > last:
            print("菜谱已经吃完了，可以再生成一份新的。")
        return
    idx = state["days"].index(hit)
    notes = state.get("cond_notes", {})
    ppl = state.get("people", 1)
    print("%s 周%s（菜谱第 %d 天，共 %d 天）%s" % (day.strftime("%-m月%-d日"), WEEK[day.weekday()], idx + 1,
                                         len(state["days"]), "" if ppl == 1 else "，用量是一人份，%d 个人就乘 %d" % (ppl, ppl)))
    for meal, lst in hit["meals"].items():
        if a.meal and a.meal != meal:
            continue
        print("\n【%s】%s" % (meal, "、".join(lst)))
        if a.brief:
            continue
        for x in lst:
            d = M_BY.get(x)
            if not d:
                continue
            print("\n▶ %s\n用料：%s\n调料：%s" % (x, d[4].replace("\n", "，"), d[5] + (
                "（不吃%s就不放）" % "、".join(notes[x]) if x in notes else "")))
            print(d[6])
    plan = [d["meals"] for d in state["days"]]
    note = thaw_note(plan, idx)
    if note:
        print("\n今晚：" + note)


def cmd_prices(a):
    gov, warn = load_prices(a.offline)
    if not gov:
        print(warn)
        return
    print("上海市主要主副食品品种价格信息表 %s（全市均价，元/斤；米是10公斤一袋，油是5升一桶）" % gov["date"])
    print(gov["url"])
    words = split_words(a.item) if a.item else []
    for k, v in gov["items"].items():
        name, spec = k.split("|")
        if words and not any(w in name or w in spec for w in words):
            continue
        print("%s%s：%.2f" % (name, "（%s）" % spec if spec else "", v))
    if warn:
        print("注意：" + warn)


def cmd_dishes(a):
    label = {"main": "荤菜", "side": "蛋·豆腐", "veg": "素菜", "soup": "汤", "breakfast": "早餐"}
    for t, n in label.items():
        ds = [d[0] + ("🌶" if "spicy" in d[2] else "") for d in M.D if d[1] == t]
        print("%s（%d）：%s" % (n, len(ds), "、".join(ds)))


def main():
    ap = argparse.ArgumentParser(description="家常菜谱生成器（上海菜价）")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("plan", help="生成菜谱 Excel")
    p.add_argument("--budget", type=float, required=True, help="预算（元），默认是整段时间的总额")
    p.add_argument("--per", choices=("total", "week", "day"), default="total",
                   help="预算口径：total=总额（默认），week=每周，day=每天")
    p.add_argument("--days", type=int, help="天数")
    p.add_argument("--start", help="开始日期，默认明天；可写 今天/明天/10-5")
    p.add_argument("--until", help="结束日期（含），和 --days 二选一，如 10-31")
    p.add_argument("--meals", type=int, choices=(1, 2, 3), default=2, help="每天几顿：2=午晚（默认），3=加早餐，1=只晚餐")
    p.add_argument("--people", type=int, default=1, help="几个人吃")
    p.add_argument("--like", default="", help="爱吃的，逗号分隔，如 猪肉,羊肉,茄子")
    p.add_argument("--dislike", default="", help="不吃的，逗号分隔，如 香菜,鱼,辣")
    p.add_argument("--no-spicy", action="store_true", help="不吃辣")
    p.add_argument("--pantry", action="store_true", help="家里已有米、油、调料，不计入")
    p.add_argument("--name", default="", help="表格标题里的称呼")
    p.add_argument("--seed", type=int, help="换一种排法时用不同的数字")
    p.add_argument("--offline", action="store_true", help="不联网，用缓存或兜底价")
    p.add_argument("--out-dir", help=argparse.SUPPRESS)
    t = sub.add_parser("today", help="查某天吃什么")
    t.add_argument("--date", help="日期，默认今天；可写 明天/后天/10-5")
    t.add_argument("--meal", choices=("早餐", "午餐", "晚餐"), help="只看某一顿")
    t.add_argument("--brief", action="store_true", help="只列菜名")
    pr = sub.add_parser("prices", help="上海最新菜价")
    pr.add_argument("--item", help="只看某几样，逗号分隔")
    pr.add_argument("--offline", action="store_true")
    sub.add_parser("dishes", help="列出菜库")
    a = ap.parse_args()
    {"plan": cmd_plan, "today": cmd_today, "prices": cmd_prices, "dishes": cmd_dishes}[a.cmd](a)


if __name__ == "__main__":
    main()
