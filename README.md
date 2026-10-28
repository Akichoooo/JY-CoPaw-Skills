# JY 的 CoPaw 技能合集

一个技能一个文件夹，按项目归类，按需取用。技能文件夹保持在根目录，复制单个文件夹即可安装。

安装目录：
- Qoder：`C:\Users\92586\.qoder-cn\skills\`
- ZCode/Claude 风格：`C:\Users\92586\.agents\skills\`

找技能：先看「项目」，再看「触发词」；用 AI 时直接说任务，不用背技能名。

| 项目 | 技能 | 目录 | 干什么 | 怎么用 | 触发词 |
|---|---|---|---|---|---|
| 百度一见 | 预警准确率报告 | `alarm-accuracy-report/` | 每日预警导出 → 准确率日报（含误报原因）；按台账汇总月报；月末归档 | 只有 `SKILL.md` + 脚本，AI 读说明后自己跑脚本 | 预警准确率、日报、月报、误报统计 |
| 安防水尺 | 摄像机掉线时长台账 | `spd-offline-ledger/` | 南水北调中线摄像机掉线时长/在线率周报台账 | 只有 `SKILL.md` + 脚本，AI 读说明后自己跑脚本 | 掉线台账、在线率台账、剔除后表、运维掉线时长 |
| 通用 | SerpApi 搜索 | `serpapi-search/` | 调 SerpApi 做网页/新闻搜索 | 有 `skill.yaml` + `skill.py` 入口，需要 API key | 联网搜索、查新闻、websearch |

项目只是分类标签，不单独建目录：安装时拷贝的是 `alarm-accuracy-report/` 这一层，不是项目层。

- **百度一见**：视觉算法预警平台，产出预警准确率日报/月报。
- **安防水尺**：南水北调中线摄像机掉线/在线率台账。
- **通用**：与具体项目无关的工具。

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

## alarm-accuracy-report 用法

```bash
# 日报：导出表 + 各技能误报数 + 误报原因（未提到的技能按 0 误报）
python alarm-accuracy-report/scripts/accuracy_report.py daily <导出.xlsx> \
  --date 2026-10-26 \
  --false "夜间区域入侵=3" --false "区域入侵=1" --false "烟雾识别=all" \
  --reason "夜间区域入侵=算法误识别：把门识别成人" \
  --reason "区域入侵=算法误识别：022镜头把围栏识别成人"

# 月报：读日报自动维护的 台账/准确率台账.csv
python alarm-accuracy-report/scripts/accuracy_report.py monthly \
  --month 2026-10 --ledger <工作目录>/台账/准确率台账.csv

# 手动整理旧日报/月报（daily/monthly 默认会自动整理）
python alarm-accuracy-report/scripts/accuracy_report.py tidy \
  --outdir <工作目录>

# 月末归档：日报/月报 → 归档/2026-10/（台账不动）
python alarm-accuracy-report/scripts/accuracy_report.py archive \
  --month 2026-10 --outdir <工作目录>
```

默认输出：`准确率日报_YYYY-MM-DD.xlsx`（每天一份，不覆盖昨天）和 `准确率月报_YYYY-MM.xlsx`；根目录只留最新一份日报和最新一份月报，旧日报自动进 `归档/YYYY-MM/日报/`、旧月报进 `归档/YYYY-MM/`。台账收在 `台账/` 子目录。误报原因会写入日报和台账；没有填的技能显示“未填写”，脚本也会提示。导出里有测试技能（如 `机器狗测试技能`）时，加 `--exclude "机器狗测试技能"`，不计入总数和准确率。月末归档只移动报告，台账留在 `台账/`，历史月报仍可重跑。

口径细节见 `alarm-accuracy-report/SKILL.md`。

## 加新技能

在仓库根目录新建一个同名文件夹，放进 `SKILL.md`（必须）和脚本；如果这个技能需要
Python 入口，再加 `skill.yaml` + `skill.py`。然后在上面表格里补一行，项目列填
`百度一见` / `安防水尺` / `通用`。
