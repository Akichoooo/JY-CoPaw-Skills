# -*- coding: utf-8 -*-
"""fill_xinbiao.py — 按新口径填写 新表.xlsx 的 Q/R/S/T/U 列

用法: python fill_xinbiao.py <工作目录> [新表.xlsx路径] [太阳能目录]

新口径(2026-09 二次修订):
    Q + R + S + T = P             # 安防 + 闸站 + 周界 + 左排 = 运维掉线时长
    U = (Q∪R∪S∪T 中命中太阳能表的份额)   # U 是 P 内部的一个子集, 不参与上面那个加总
  - P(运维掉线时长)、N(剔除计划性掉线镜头数量)、O(剔除掉线时长) 已由人工填好, 本脚本不再改动;
  - 把 P 对应的行集(运维处反馈表优先, 否则剔除后表)按分组名/镜头名称分成 Q/R/S/T 四类,
    以各类掉线时长为权重, 把 P 按比例拆成 Q/R/S/T(T 取残差), 四项之和严格等于 P;
  - U = P × (命中太阳能表的行权重 / 四类总权重), 即 P 中属于太阳能摄像机的那一部分;
  - 分类优先级(具体 → 宽泛):
        S 周界 : 分组(三~五级)/名称含"周界"
        T 左排 : 分组/名称含"左排"; 无任何类别关键字的行也归 T
        R 闸站 : 三级分组含"闸站"
        Q 安防 : 其余(分组含"安防")
  - U 的判定独立于上面四类: 镜头名称命中 <太阳能摄像机供电时长统计> 表里同一管理处的摄像机名 → 计入 U 权重
  - 基准行 = 批量标记空 且 备注空; 基准为空时退化为按全部行占比。
  - 反馈表通常是旧口径生成的(缺周界行), 会自动从剔除后表补入周界行, 保证 S 有权重。
  - 合计行只重算 Q~U。

列位: L总镜头数量 M总镜头在线时长 N剔除数量 O剔除掉线时长 P运维掉线时长
      Q安防 R闸站 S周界 T左排 U太阳能 V在线率  (L~P 不写)
"""
import sys, os, glob, re, shutil, collections
from openpyxl import load_workbook

BASE = sys.argv[1] if len(sys.argv) > 1 else r"."
XB   = sys.argv[2] if len(sys.argv) > 2 else os.path.join(BASE, "新表.xlsx")
SOLAR_ARG = sys.argv[3] if len(sys.argv) > 3 else None

DEFAULT_SOLAR = r"D:\BaiduSyncdisk\work\掉线时长\太阳能摄像机供电时长统计"


def norm(s):
    return re.sub(r'[\s（）()\[\]【】、,，.。\-—_/+]', '', str(s or '')).lower()


def name_key(s):
    """镜头名称归一：去标点/大小写 + 去前导编号 + 去尾部 (全)/(细)。"""
    k = norm(s)
    k = re.sub(r'^0*\d+', '', k)
    k = re.sub(r'(全|细|全细)$', '', k)
    return k


def solar_dir():
    if SOLAR_ARG:
        return SOLAR_ARG
    parent = os.path.dirname(os.path.abspath(BASE))
    for up in (BASE, parent):
        try:
            for n in os.listdir(up):
                p = os.path.join(up, n)
                if os.path.isdir(p) and '太阳能' in n and '统计' in n:
                    return p
        except OSError:
            pass
    return DEFAULT_SOLAR


def load_solar(root):
    """管理处 -> {name_key,...}；同时兼容 供电电池类型 里的 渠道市电（整表视为太阳能名单）。"""
    idx = collections.defaultdict(set)
    for f in glob.glob(os.path.join(root, '*.xlsx')):
        if os.path.basename(f).startswith('~$'):
            continue
        try:
            wb = load_workbook(f, data_only=True)
        except Exception as e:
            print(f"  [warn] 读太阳能表失败 {os.path.basename(f)}: {e}")
            continue
        ws = wb[wb.sheetnames[0]]
        hr = next((r for r in range(1, 6) if str(ws.cell(row=r, column=1).value).strip() == '分公司'), None)
        if hr is None:
            continue
        mg = None
        for r in range(hr + 1, ws.max_row + 1):
            v = ws.cell(row=r, column=2).value
            if v:
                mg = str(v).strip()
            nm = ws.cell(row=r, column=3).value
            if mg and nm:
                idx[mg].add(name_key(nm))
    return idx


def mg_key(s):
    return re.sub(r'[（(].*?[)）]', '', str(s or '')).strip()


def is_batch(v):
    return v is not None and '批量' in str(v)


def read_xlsx(path):
    """返回 (rows, 剔除计划性掉线时长)。行字段: g2~g5/name/ip/dur/batch/reason。"""
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    hr = next((r for r in range(1, 10) if ws.cell(row=r, column=1).value == '序号'), None)
    if hr is None:
        return [], None
    rows = []
    summ = None
    for r in range(hr + 1, ws.max_row + 1):
        if isinstance(ws.cell(row=r, column=1).value, int):
            c12 = ws.cell(row=r, column=12).value
            c13 = ws.cell(row=r, column=13).value if ws.max_column >= 13 else None
            batch = c12 if is_batch(c12) else (c13 if is_batch(c13) else None)
            reason = (c13 if (c13 not in (None, '') and not is_batch(c13))
                      else (c12 if (c12 not in (None, '') and not is_batch(c12)) else None))
            rows.append(dict(g2=ws.cell(row=r, column=3).value, g3=ws.cell(row=r, column=4).value,
                             g4=ws.cell(row=r, column=5).value, g5=ws.cell(row=r, column=6).value,
                             name=ws.cell(row=r, column=8).value, ip=ws.cell(row=r, column=9).value,
                             dur=ws.cell(row=r, column=10).value or 0,
                             batch=batch, reason=reason))
        elif '剔除计划性掉线时长' in str(ws.cell(row=r, column=1).value or ''):
            summ = ws.cell(row=r, column=2).value
    return rows, summ


def is_solar(x, solar_names):
    name = str(x['name'] or '')
    return bool(name) and name_key(name) in solar_names


def classify(x):
    """P 内部四分类: 周界 S / 左排 T / 闸站 R / 安防 Q。

    以"分组名"为准:
      - 分组(三~五级)含"周界"(如 周界安防视频) → S
      - 分组/名称含"左排"(如 四级分组=左排、名称 -左排) → T
      - 三级分组含"闸站" → R; 闸站组里名字带"周界"的相机(如 屯庄南1#保水堰周界01)仍算闸站,
        徐水/天津就属于这种情况, 不能因为名称里有"周界"就记成 S
      - 安防组里的"…-周界智能"球机(顺平) → S
      - 其余(含"安防") → Q; 无任何关键字 → T
    """
    g = f"{x['g3'] or ''} {x['g4'] or ''} {x['g5'] or ''}"
    name = str(x['name'] or '')
    if '周界' in g:
        return 'S'
    if '左排' in g or '左排' in name:
        return 'T'
    if '闸站' in str(x['g3'] or ''):
        return 'R'
    if '周界' in name:
        return 'S'
    if '安防' in g or '安防' in name:
        return 'Q'
    return 'T'


def weights(rows, solar_names):
    """返回 (四类权重 dict, 命中太阳能表的权重)。太阳能行已含在四类里, 是子集。"""
    w = collections.defaultdict(int)
    wu = 0
    for x in rows:
        w[classify(x)] += x['dur']
        if is_solar(x, solar_names):
            wu += x['dur']
    return w, wu


def split_p(P, w, wu):
    """把 P 拆成 Q/R/S/T(T 取残差, 四项之和 = P); U = P 中太阳能权重占比。"""
    W = sum(w.values())
    if W <= 0 or not isinstance(P, (int, float)) or P <= 0:
        return 0, 0, 0, 0, 0, False
    q = round(P * w.get('Q', 0) / W)
    r = round(P * w.get('R', 0) / W)
    s = round(P * w.get('S', 0) / W)
    t = P - q - r - s
    u = round(P * wu / W)
    return q, r, s, t, u, True


# ---- 源文件索引 ----
fb = {}
for f in glob.glob(os.path.join(BASE, '运维处反馈', '**', '*.xlsx'), recursive=True):
    fb[os.path.basename(f).split('_')[0]] = f
my = {}
for f in glob.glob(os.path.join(BASE, '剔除后表_按运维处', '**', '*.xlsx'), recursive=True):
    my[os.path.basename(f).split('_')[0]] = f

SDIR = solar_dir()
SOLAR = load_solar(SDIR)
print(f"太阳能名单: {SDIR}  ->  {len(SOLAR)} 个管理处")

wb = load_workbook(XB)
ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb[wb.sheetnames[0]]
report = []
for row in range(2, ws.max_row + 1):
    org = ws.cell(row=row, column=4).value
    if not org:
        continue
    P = ws.cell(row=row, column=16).value
    key = mg_key(org).replace('未计算在内', '')
    src = None
    is_fb = False
    if org == '顺平管理处' and '保定管理处' in fb:
        src = fb['保定管理处']; is_fb = True
    elif key in fb:
        src = fb[key]; is_fb = True
    else:
        k = org.replace('分公司', '分公司本部') if str(org).endswith('分公司') else key
        src = my.get(k) or my.get(key)
    if src is None:
        # 无源(如邢台全部超6天被剔除): 保留原有 Q~U 不动
        report.append((org, '无源', P, ws.cell(row=row, column=17).value or 0, ws.cell(row=row, column=18).value or 0,
                       ws.cell(row=row, column=19).value or 0, ws.cell(row=row, column=20).value or 0,
                       ws.cell(row=row, column=21).value or 0))
        continue
    rows, _ = read_xlsx(src)
    if is_fb:
        # 反馈表多为旧口径生成、缺周界行 -> 从剔除后表补周界，保证 S 有权重
        base_path = my.get(key) or my.get(org.replace('分公司', '分公司本部') if str(org).endswith('分公司') else key)
        if base_path:
            have = {(str(x['name']), x['ip'], x['dur']) for x in rows}
            extra, _ = read_xlsx(base_path)
            for x in extra:
                txt = f"{x['g3'] or ''} {x['g4'] or ''} {x['g5'] or ''} {x['name'] or ''}"
                if '周界' in txt and (str(x['name']), x['ip'], x['dur']) not in have:
                    rows.append(x)
    if org == '顺平管理处':
        rows = [x for x in rows if '顺平' in str(x['g4'] or '') + str(x['g5'] or '')]
    if org == '保定管理处':
        rows = [x for x in rows if '顺平' not in str(x['g4'] or '') + str(x['g5'] or '')]

    snames = SOLAR.get(key, set())
    base_rows = [x for x in rows if not x['batch'] and not x['reason']]
    w, wu = weights(base_rows, snames)
    q, r, s, t, u, ok = split_p(P, w, wu)
    if not ok and isinstance(P, (int, float)) and P > 0:
        w, wu = weights(rows, snames)
        q, r, s, t, u, ok = split_p(P, w, wu)
    if isinstance(P, (int, float)) and P > 0 and ok:
        assert q + r + s + t == P, f"{org} Q+R+S+T 不等于 P"
    ws.cell(row=row, column=17).value = q
    ws.cell(row=row, column=18).value = r
    ws.cell(row=row, column=19).value = s
    ws.cell(row=row, column=20).value = t
    ws.cell(row=row, column=21).value = u
    report.append((org, 'FB' if is_fb else 'MY', P, q, r, s, t, u))

# 合计行: 只重算 Q~U(17~21), N/O/P 不动
# 合计行特征: 管理处(D)为空、序号(A)也空、但数量(E)是数字
for row in range(2, ws.max_row + 1):
    if (ws.cell(row=row, column=4).value is None and ws.cell(row=row, column=1).value is None
            and isinstance(ws.cell(row=row, column=5).value, (int, float))):
        tot = collections.defaultdict(int)
        for r2 in range(2, ws.max_row + 1):
            if ws.cell(row=r2, column=4).value is None:
                continue
            for c in (17, 18, 19, 20, 21):
                v = ws.cell(row=r2, column=c).value
                if isinstance(v, (int, float)):
                    tot[c] += v
        for c in (17, 18, 19, 20, 21):
            ws.cell(row=row, column=c).value = tot[c]

bak = XB.replace('.xlsx', '_备份_新口径.xlsx')
if not os.path.exists(bak):
    shutil.copy(XB, bak)
for fn in (XB, XB.replace('.xlsx', '_已填.xlsx')):
    try:
        wb.save(fn); print(f"已保存 -> {fn}"); break
    except PermissionError:
        print(f"{fn} 被占用, 尝试其他文件名")

print(f"\n{'管理处':<14}{'源':>3}{'P':>8}{'Q':>8}{'R':>7}{'S':>7}{'T':>7}{'U':>7}{'校验':>5}")
for x in report:
    P, (q, r, s, t, u) = x[2], x[3:8]
    if not isinstance(P, (int, float)) or P <= 0:
        chk = '-'
    else:
        chk = 'OK' if q + r + s + t == P else 'BAD'
    print(f"{str(x[0]):<14}{str(x[1]):>3}{str(P):>8}{q:>8}{r:>7}{s:>7}{t:>7}{u:>7}{chk:>5}")
