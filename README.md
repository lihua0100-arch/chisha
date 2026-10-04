# 吃啥

给预算、时间、口味，生成一份能照着做的家常菜谱 Excel；之后随时问「今天吃啥」。

- **纯 Python 标准库**，不用装任何依赖（xls 读取和 xlsx 写入都是自己写的）
- **实时菜价**：每次生成先抓上海市政府「上海市主要主副食品品种价格信息表」最新一期的全市均价，抓不到就用内置兜底价
- **113 道家常菜**：55 荤 / 14 蛋豆腐 / 29 素 / 8 汤 / 7 早餐，每道都有家常用量、调料和做法
- **钱不够不硬凑**：先不算米油盐，再算实际够几顿就排几顿并说明；连一顿都不够就不出表

## 用法

```bash
# 一周 300 块，两个人，爱吃牛肉，不吃辣
python3 meal_plan.py plan --budget 300 --per week --days 7 --people 2 --like 牛肉 --no-spicy

# 今天吃啥 / 明天晚上吃啥
python3 meal_plan.py today
python3 meal_plan.py today --date 明天 --meal 晚餐

# 上海最新菜价
python3 meal_plan.py prices --item 猪肉,青菜

# 菜库
python3 meal_plan.py dishes
```

`plan` 的常用参数：

| 参数 | 说明 |
|---|---|
| `--budget` | 预算（元） |
| `--per total\|week\|day` | 预算是总额 / 每周 / 每天，默认总额 |
| `--days` / `--until` | 天数，或结束日期（如 `10-31`） |
| `--start` | 开始日期，默认明天 |
| `--meals 1\|2\|3` | 每天几顿：1=晚饭，2=午晚（默认），3=加早饭 |
| `--people` | 几个人吃 |
| `--like` / `--dislike` | 爱吃 / 不吃的，逗号分隔 |
| `--no-spicy` | 不吃辣 |
| `--pantry` | 家里有米油调料，不计入 |
| `--seed` | 换一批菜 |
| `--offline` | 不联网，用缓存或兜底价 |

生成的 Excel 有五张表：每日做饭、菜谱总览、采购清单、做法大全、单价表。

## 手机网页版

```bash
python3 meal_plan.py web -o 菜谱.html   # 把当前菜谱生成一个单文件网页，手机浏览器直接打开
```

网页分四页：今天吃啥（按天切换、提前解冻提醒）、菜单、买菜（按周分组打勾）、做法（可搜索，每道菜有分步做法）。
`meal_plan_template.html` 要和 `meal_plan.py` 放在同一个目录。

想让 `plan` 生成后自动发布到自己的网站，写一个配置文件 `~/.config/chisha/publish.json`：

```json
{"base": "https://你的域名", "dirs": ["/网站根目录"]}
```

网页会写到 `<网站根目录>/m/<随机串>/index.html`，链接打印在输出最后，重新生成菜谱链接不变。没有这个配置就不发布。

## 存档位置

菜谱存在 `$HERMES_HOME/workspace/meal-plans/`，没设 `HERMES_HOME` 时是 `~/.hermes/workspace/meal-plans/`。
`today` 查的是最近一次 `plan` 生成的那份。菜价缓存 6 小时，放在 `/tmp/meal-plan-cache`。

## 接到 Hermes 聊天机器人

`skill/SKILL.md` 是给 Hermes 聊天机器人用的技能说明：
把两个 `.py` 放进 Hermes 的 `tools/`，`SKILL.md` 放进 `skills/<分类>/chisha/`，
在微信里发「吃啥」或「一周 500 块的菜谱」就能用。文档里的 `/opt/data` 是 Hermes 容器里的数据目录，按你的部署改。

## 局限

- 价格只有上海的；别的城市按上海均价估算
- 鸡胸、鸡腿、豆腐、皮蛋等价格表里没有的食材用估算价
- 菜库是固定的，要加菜改 `meal_plan_dishes.py`

## 许可证

MIT
