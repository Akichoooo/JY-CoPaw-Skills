# JY 的 CoPaw 技能合集

一个技能一个文件夹，按需取用：要用哪个就把那个文件夹拷进你的 skills 目录
（ZCode/Claude 风格：`C:\Users\92586\.agents\skills\`）。

| 技能 | 目录 | 干什么 | 怎么用 | 触发词 |
|---|---|---|---|---|
| SerpApi 搜索 | `serpapi-search/` | 调 SerpApi 做网页/新闻搜索 | 有 `skill.yaml` + `skill.py` 入口，需要 API key | 联网搜索、查新闻、websearch |
| 摄像机掉线时长台账 | `spd-offline-ledger/` | 南水北调中线摄像机掉线时长/在线率周报台账 | 只有 `SKILL.md` + 脚本，AI 读说明后自己跑脚本 | 掉线台账、在线率台账、剔除后表、运维掉线时长 |

## serpapi-search 配置

```bash
cd serpapi-search
cp .env.example .env      # 填 SERPAPI_API_KEY（申请：https://serpapi.com/manage-api-key）
pip install -r requirements.txt
```

`.env` 已被 `.gitignore` 忽略，不要提交。

## spd-offline-ledger 用法

```bash
# 1) 分公司 CSV → 剔除后表（周界保留；保定/顺平自动拆成两张表）
python spd-offline-ledger/scripts/build_tichu.py <工作目录>

# 2) 邢台：告警导出 → 剔除后表（第三个参数=资源唯一标识清单，可选）
python spd-offline-ledger/scripts/build_xingtai_tichu.py <工作目录> [告警文件] [清单.txt]

# 3) 按运维处反馈备注判断剔除，算 N/O/P
python spd-offline-ledger/scripts/calc_nop.py <工作目录> [新表.xlsx]

# 4) 按 P 拆 Q/R/S/T/U（Q+R+S+T=P，U 为命中太阳能表的子集）
python spd-offline-ledger/scripts/fill_xinbiao.py <工作目录> [新表.xlsx] [太阳能目录]
```

口径细节见 `spd-offline-ledger/SKILL.md`。

## 加新技能

在仓库根目录新建一个同名文件夹，放进 `SKILL.md`（必须）和脚本；如果这个技能需要
Python 入口，再加 `skill.yaml` + `skill.py`。然后在上面表格里补一行。
