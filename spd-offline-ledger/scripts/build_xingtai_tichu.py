# -*- coding: utf-8 -*-
"""build_xingtai_tichu.py — 把邢台告警导出转成"邢台管理处_剔除后表.xlsx"

用法: python build_xingtai_tichu.py <工作目录> [告警文件] [资源唯一标识清单.txt]

背景: 邢台管理处在分公司 CSV 里没有数据，运维处反馈的是"摄像头离线告警"导出，
      掉线时长 = 事件清除时间 − 事件发生时间(分钟)，按相机(资源唯一标识)汇总。

判定(二选一):
  1) 给了"资源唯一标识清单"文件(每行一个ID): 清单里的相机 = 运维(计入P)，
     清单外的 = 剔除(记入剔除统计)。这是运维处点名确认的口径，优先使用。
  2) 没给清单: 按离线原因关键词，计划性/人工操作(维护、更换、拆装、远程重启、
     调整预置位、清理、门禁…) = 剔除; 故障及无原因 = 运维。

输出: <工作目录>/剔除后表_按运维处/邢台运维处/邢台管理处_剔除后表.xlsx
      表内保留全部相机, 剔除行的备注列写"剔除"; 末列=完整离线原因(仅供查看)。
"""
import sys, os, glob, datetime, collections

BASE = sys.argv[1] if len(sys.argv) > 1 else r"."
ARG = sys.argv[2] if len(sys.argv) > 2 else None
IDS = sys.argv[3] if len(sys.argv) > 3 else None

# 计划性/人工操作关键词 → 剔除 (仅在未提供资源唯一标识清单时使用)
PLAN = ['维护', '更换', '拆装', '远程重启', '调整预置位', '清理', '门禁', '后台维护']


def find_alarm(base):
    if ARG:
        return ARG
    for p in ('运维处反馈/**/*告警*.xls*', '**/邢台*告警*.xls*', '**/*告警*.xlsx'):
        hits = [f for f in glob.glob(os.path.join(base, p), recursive=True)
                if not os.path.basename(f).startswith('~$')]
        if hits:
            return max(hits, key=os.path.getmtime)
    return None


def parse_dt(s):
    s = str(s or '').strip()
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M'):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def read_alarms(path):
    """返回 [dict(rid, name, ip, dur, reason)]，兼容 .xls / .xlsx。"""
    if path.lower().endswith('.xls'):
        import xlrd
        sh = xlrd.open_workbook(path).sheet_by_index(0)
        idx = {str(sh.cell_value(0, c)).strip(): c for c in range(sh.ncols)}
        rng = range(1, sh.nrows)
        cell = lambda r, n: sh.cell_value(r, idx[n]) if n in idx else ''
    else:
        from openpyxl import load_workbook
        ws = load_workbook(path, data_only=True).worksheets[0]
        idx = {str(ws.cell(row=1, column=c).value or '').strip(): c for c in range(1, ws.max_column + 1)}
        rng = range(2, ws.max_row + 1)
        cell = lambda r, n: ws.cell(row=r, column=idx[n]).value if n in idx else ''
    out = []
    for r in rng:
        nm = str(cell(r, '网元名称') or '').strip()
        if not nm:
            continue
        occ = parse_dt(cell(r, '事件发生时间')); clr = parse_dt(cell(r, '事件清除时间'))
        d = int((clr - occ).total_seconds() // 60) if (occ and clr) else 0
        out.append(dict(rid=str(cell(r, '资源唯一标识')).strip(), name=nm,
                        ip=str(cell(r, '网元IP') or '').strip(), dur=d,
                        reason=str(cell(r, '离线原因') or '').strip()))
    return out


def is_plan(reason):
    return any(k in reason for k in PLAN)


def main():
    src = find_alarm(BASE)
    if not src:
        print('未找到邢台告警文件'); return
    print('告警文件:', src)
    events = read_alarms(src)

    if IDS:
        # 按资源唯一标识清单: 清单内=运维, 清单外=剔除; 按相机汇总
        with open(IDS, encoding='utf-8-sig') as fh:
            ids = {l.strip() for l in fh if l.strip()}
        cams = {}
        for e in events:
            c = cams.setdefault(e['rid'], dict(name=e['name'], ip=e['ip'], dur=0, cnt=0,
                                               reasons=[], excl=(e['rid'] not in ids)))
            c['dur'] += e['dur']; c['cnt'] += 1
            if e['reason'] and e['reason'] not in c['reasons']:
                c['reasons'].append(e['reason'])
        rows = [(c['name'], c['ip'], c['dur'], c['cnt'], c['excl'], ' / '.join(c['reasons']))
                for c in cams.values()]
        mode = f'资源唯一标识清单({len(ids)}个ID)'
    else:
        # 按离线原因关键词: 逐条事件判定
        rows = [(e['name'], e['ip'], e['dur'], 1, is_plan(e['reason']), e['reason']) for e in events]
        mode = '离线原因关键词'
    rows.sort(key=lambda x: (x[2], x[0]))

    oper = [x for x in rows if not x[4]]
    excl = [x for x in rows if x[4]]
    retN, retD = len(oper), sum(x[2] for x in oper)
    remN, remD = len(excl), sum(x[2] for x in excl)

    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook(); ws = wb.active; ws.title = '邢台管理处'
    ws['A1'] = '统计区间: 本周  (来源: 邢台告警导出, 掉线时长 = 事件清除时间 − 事件发生时间)'
    ws['A3'] = f'邢台管理处 剔除后表(告警) | 判定方式: {mode} | 备注列写"剔除"的行为剔除, 其余为运维; 末列=完整离线原因'
    hdrs = ['序号', '一级分组', '二级分组', '三级分组', '四级分组', '五级分组', '六级分组',
            '镜头名称', '设备IP地址', '掉线时长', '掉线次数', '批量掉线标记', '备注', '离线原因(全部)']
    for ci, h in enumerate(hdrs, 1):
        c = ws.cell(row=4, column=ci, value=h)
        c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='1B2A4A')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    r = 5
    for i, (nm, ip, d, cnt, is_excl, rs) in enumerate(rows, 1):
        vals = [i, '河北分公司', '邢台管理处', '安防视频', None, None, None,
                nm, ip, d, cnt, None, ('剔除' if is_excl else None), rs]
        for ci, v in enumerate(vals, 1):
            ws.cell(row=r, column=ci, value=v)
        r += 1
    r += 1
    for label, val in [('掉线镜头数量(剔除后)', retN), ('掉线镜头时长(分钟)', retD),
                       ('剔除计划性掉线镜头数量', remN), ('剔除计划性掉线时长(分钟)', remD)]:
        ws.cell(row=r, column=1, value=label).font = Font(size=10)
        ws.cell(row=r, column=2, value=val); r += 1
    for ci, wd in enumerate([5, 10, 10, 12, 12, 12, 12, 34, 15, 9, 8, 22, 14, 40], 1):
        ws.column_dimensions[get_column_letter(ci)].width = wd
    out = os.path.join(BASE, '剔除后表_按运维处', '邢台运维处', '邢台管理处_剔除后表.xlsx')
    os.makedirs(os.path.dirname(out), exist_ok=True)
    wb.save(out)

    print(f'已生成 {out}   (判定方式: {mode})')
    print(f'  运维(保留): {len(oper)} 台 / {retD} 分钟   -> P')
    print(f'  剔除:       {len(excl)} 台 / {remD} 分钟  -> N / O')
    print('  剔除明细:')
    for nm, ip, d, cnt, is_excl, rs in excl:
        print(f'    {d:>6}分 x{cnt:<2} {nm}  | {rs[:50]}')

    return retN, retD, remN, remD


if __name__ == '__main__':
    main()
