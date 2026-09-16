# -*- coding: utf-8 -*-
"""calc_nop.py — 按"备注判断剔除"口径填写 新表.xlsx 的 N/O/P 三列

用法: python calc_nop.py <工作目录> [新表.xlsx路径]

口径(2026-09-11 起):
    剔除 = 备注里写明了计划性/人工/非真实掉线的行
        (关键词: 计划 停电 施工 维护 检修 更换 拆装 拆除 报备 割接 门禁 未掉线 统计出错 建议删除)
        故障类(电缆故障/电池恶化/电压低/泡水/太阳能供电不足等)不算剔除, 留在 P 里
    N(剔除计划性掉线镜头数量) = 剔除行数 + 剔除后表里"隐藏/超6天/空分组"的剔除镜头数
    O(剔除掉线时长)         = 剔除行时长 + 同上的剔除时长
    P(运维掉线时长)         = 该管理处全部掉线时长 − 剔除行时长
数据源: 运维处反馈优先(可带备注), 否则剔除后表; 保定表按四级分组含"顺平"拆出顺平管理处。
N/O/P 由本脚本填; Q/R/S/T/U 由 fill_xinbiao.py 按 P 拆分。
"""
import sys, os, glob, re, collections
from openpyxl import load_workbook

BASE = sys.argv[1] if len(sys.argv) > 1 else r"."
XB = sys.argv[2] if len(sys.argv) > 2 else os.path.join(BASE, "新表.xlsx")

# 计划性/人工/非真实 → 剔除
PLAN_KW = ['计划', '停电', '施工', '维护', '检修', '更换', '拆装', '拆除',
           '报备', '割接', '门禁', '未掉线', '统计出错', '建议删除',
           '清理', '预置位', '移位', '远程重启', '剔除']


def is_batch(v):
    return v is not None and '批量' in str(v)


def mg_key(s):
    return re.sub(r'[（(].*?[)）]', '', str(s or '')).strip()


def read_xlsx(path):
    wb = load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]
    hr = next((r for r in range(1, 10) if ws.cell(row=r, column=1).value == '序号'), None)
    rows = []
    if hr is None:
        return rows
    for r in range(hr + 1, ws.max_row + 1):
        if isinstance(ws.cell(row=r, column=1).value, int):
            c12 = ws.cell(row=r, column=12).value
            c13 = ws.cell(row=r, column=13).value if ws.max_column >= 13 else None
            notes = [str(v).strip() for v in (c12, c13) if v not in (None, '') and not is_batch(v)]
            rows.append(dict(g4=ws.cell(row=r, column=5).value, g5=ws.cell(row=r, column=6).value,
                             name=ws.cell(row=r, column=8).value,
                             dur=ws.cell(row=r, column=10).value or 0,
                             note=' '.join(notes)))
    return rows


def stats(path):
    """剔除后表统计块: (剔除镜头数, 剔除时长, 保留时长)。"""
    wb = load_workbook(path, data_only=True); ws = wb[wb.sheetnames[0]]
    out = {}
    for r in range(1, ws.max_row + 1):
        v = str(ws.cell(row=r, column=1).value or '').strip()
        if v == '剔除计划性掉线镜头数量':
            out['remN'] = ws.cell(row=r, column=2).value or 0
        elif v == '剔除计划性掉线时长(分钟)':
            out['remD'] = ws.cell(row=r, column=2).value or 0
        elif v == '掉线镜头时长(分钟)':
            out['retD'] = ws.cell(row=r, column=2).value or 0
    return out


def is_plan(note):
    return any(k in note for k in PLAN_KW)


def main():
    fb, my = {}, {}
    for f in glob.glob(os.path.join(BASE, '运维处反馈', '**', '*.xlsx'), recursive=True):
        fb[os.path.basename(f).split('_')[0]] = f
    for f in glob.glob(os.path.join(BASE, '剔除后表_按运维处', '**', '*.xlsx'), recursive=True):
        my[os.path.basename(f).split('_')[0]] = f

    wb = load_workbook(XB); ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb[wb.sheetnames[0]]
    print(f"{'管理处':<18}{'总时长':>9}{'剔除行':>7}{'剔除时长':>9}{'N':>5}{'O':>9}{'P':>9}  剔除原因样例")
    for row in range(2, ws.max_row + 1):
        org = ws.cell(row=row, column=4).value
        if not org:
            continue
        key = mg_key(org).replace('未计算在内', '')
        is_fb = False
        if org == '顺平管理处' and '保定管理处' in fb:
            src, is_fb = fb['保定管理处'], True
        elif key in fb:
            src, is_fb = fb[key], True
        else:
            k = org.replace('分公司', '分公司本部') if str(org).endswith('分公司') else key
            src = my.get(k) or my.get(key)
        if src is None:
            continue
        rows = read_xlsx(src)
        if org == '顺平管理处':
            rows = [x for x in rows if '顺平' in str(x['g4'] or '') + str(x['g5'] or '')]
        if org == '保定管理处':
            rows = [x for x in rows if '顺平' not in str(x['g4'] or '') + str(x['g5'] or '')]

        total = sum(x['dur'] for x in rows)
        plan_rows = [x for x in rows if x['note'] and is_plan(x['note'])]
        planD = sum(x['dur'] for x in plan_rows)
        N, O = len(plan_rows), planD
        # 剔除后表里"隐藏/超6天/空分组"的剔除量也要计入 N/O。
        # 例外: 来源是剔除后表本身且已带备注(如邢台告警表)时, 备注已覆盖, 不重复计。
        if is_fb or not plan_rows:
            base_path = my.get(key) or my.get(org.replace('分公司', '分公司本部') if str(org).endswith('分公司') else key)
            if base_path:
                st = stats(base_path)
                N += st.get('remN', 0); O += st.get('remD', 0)
        P = total - planD
        ws.cell(row=row, column=14).value = N
        ws.cell(row=row, column=15).value = O
        ws.cell(row=row, column=16).value = P
        ex = collections.Counter(x['note'] for x in plan_rows)
        sample = '; '.join(f"{k[:26]}x{v}" for k, v in ex.most_common(2))
        print(f"{org:<18}{total:>9}{len(plan_rows):>7}{planD:>9}{N:>5}{O:>9}{P:>9}  {sample}")

    for fn in (XB, XB.replace('.xlsx', '_已填.xlsx')):
        try:
            wb.save(fn); print(f"\n已保存 -> {fn}"); break
        except PermissionError:
            print(f"{fn} 被占用, 尝试其他文件名")


if __name__ == '__main__':
    main()
