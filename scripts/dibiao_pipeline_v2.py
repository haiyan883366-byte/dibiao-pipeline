# -*- coding: utf-8 -*-
"""
dibiao-pipeline v2 —— 按《进校-每日数据更新.xlsx》模板样式生成透视报告
用法:
  python dibiao_pipeline_v2.py <底表.xlsx> [输出.xlsx]
Sheet 结构 (10 个):
  1 近三期全链路数据          期次×全指标 + 环比行
  2 分团队期次数据            基地分块 × 期次 × 全指标
  3 最新期次下渠道全链路数据  生源类别分块 × 渠道 × 全指标
  4 最新期次下学科组数据      学科组 × 全指标
  5 近三期渠道占比和转率      渠道 × 期次(占比+转率) + 环比 + 汇总
  6 近三期城市线占比和转率    城市等级 × 期次(在班+占比+转率) + 汇总
  7 最新期次下城市线全链数据  城市等级 × 全指标
  8 最新期次下分团队各渠道占比和转率  渠道 × 基地(占比+转率) + 汇总
  9 最新期次分团队电话数据    学科组 × 电话覆盖率/接通率
 10 最新期次辅导数据          预留空 sheet
"""
import sys, os
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

GRAY = PatternFill('solid', fgColor='D6E4F0')
BLUE_TOTAL = PatternFill('solid', fgColor='DDEBF7')
THIN = Border(*[Side(style='thin', color='D9D9D9')]*4)
HAS_JW = False  # 底表是否有「加微人次」列（load 时更新）

# ---------- 指标定义 ----------
METRICS = [
    ('转化率',      'zxl'), ('出勤转化率', 'cqzxl'), ('加微率', 'jwl'),
    ('加过微率',    'jgwl'), ('删除率', 'scl'), ('APP登录率', 'appdl'),
    ('课前激活率',  'kqjh'), ('课前破冰率', 'kqpb'), ('课前深沟率', 'kqsg'),
    ('摸底测覆盖率','mdc'), ('一讲出勤率', 'jc1'), ('二讲出勤率', 'jc2'),
    ('三讲出勤率',  'jc3'), ('四讲出勤率', 'jc4'),
]

def rate(num, den):
    return num / den if den else None

def agg_metrics(rows, has_jw=False):
    """rows: list of dicts (单行数值字段已转 float)
    口径（已对照模板验证 2026-11-11）:
      转化率   = 转低人次 / 在班人次（人次转化率列与转低人次重复，勿相加）
      出勤转化率 = 转低人次 / 任一讲出勤人次
    """
    zhu = sum(r['在班人次'] for r in rows)
    cq = sum(r['任一讲出勤人数'] for r in rows)
    cv = sum(r['转低人次'] for r in rows)
    jwl = rate(sum(r.get('加微人次', 0) for r in rows), zhu) if has_jw else None
    jgwl = rate(sum(r['加过微人次'] for r in rows), zhu)
    return {
        '在班人次': zhu,
        '转化率': rate(cv, zhu),
        '出勤转化率': rate(cv, cq),
        '加微率': jwl,
        '加过微率': jgwl,
        # 口径（用户确认 2026-11-11）：删除率 = 加过微率 - 加微率；无加微列时留空
        '删除率': (jgwl - jwl) if (has_jw and jwl is not None) else None,
        'APP登录率': rate(sum(r['app_登录人次'] for r in rows), zhu),
        '课前激活率': rate(sum(r['课前激活人次'] for r in rows), zhu),
        '课前破冰率': rate(sum(r['课前破冰人次'] for r in rows), zhu),
        '课前深沟率': rate(sum(r['课前深沟人次'] for r in rows), zhu),
        '摸底测覆盖率': rate(sum(r['摸底测覆盖人次'] for r in rows), zhu),
        '一讲出勤率': rate(sum(r['一讲出勤人数'] for r in rows), zhu),
        '二讲出勤率': rate(sum(r['二讲出勤人数'] for r in rows), zhu),
        '三讲出勤率': rate(sum(r['三讲出勤人数'] for r in rows), zhu),
        '四讲出勤率': rate(sum(r['四讲出勤人数'] for r in rows), zhu),
        'OA电话覆盖率': rate(sum(r['OA电话拨打次数'] for r in rows), zhu),
        '企微电话覆盖率': rate(sum(r['企微电话拨打数'] for r in rows), zhu),
        'OA电话接通率1': rate(sum(r['OA电话接通次数'] for r in rows),
                          sum(r['OA电话拨打次数'] for r in rows)),
        '企微电话接通率1': rate(sum(r['企微电话接通数'] for r in rows),
                           sum(r['企微电话拨打数'] for r in rows)),
    }

# ---------- 数据加载 ----------
def load(src):
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.active
    it = ws.iter_rows(values_only=True)
    hdr = [str(c).strip() if c else '' for c in next(it)]
    idx = {h: i for i, h in enumerate(hdr)}
    need = ['期次', '基地', '学科组名称', '账号名称', '常驻地城市等级', '开课日期', '年级名称',
            '在班人次', '转低人次', '人次转化率', '任一讲出勤人数', '加微人次',
            '加过微人次', 'app_登录人次', '课前激活人次', '课前破冰人次',
            '课前深沟人次', '摸底测覆盖人次', '一讲出勤人数', '二讲出勤人数',
            '三讲出勤人数', '四讲出勤人数', 'OA电话拨打次数', '企微电话拨打数',
            'OA电话接通次数', '企微电话接通数']
    OPTIONAL = {'加微人次', '五讲出勤人数'}  # 缺失时相关指标留空
    HAS_JW_LOCAL = '加微人次' in idx  # 新版进校底表可能无「加微人次」列，加微率/删除率留空
    for n in need:
        if n not in idx and n not in OPTIONAL:
            raise SystemExit(f'底表缺少字段: {n}')
    data = []
    for row in it:
        if row is None or not row[idx['期次']]:
            continue
        def num(col):
            v = row[idx[col]]
            if v is None: return 0
            try: return float(v)
            except (TypeError, ValueError): return 0
        numeric_cols = [c for c in need[6:] if c in idx]
        d = {c: num(c) for c in numeric_cols}
        d.setdefault('五讲出勤人数', 0)
        d['期次'] = str(row[idx['期次']]).strip()
        d['基地'] = str(row[idx['基地']] or '').strip()
        d['学科组名称'] = str(row[idx['学科组名称']] or '').strip()
        d['账号名称'] = str(row[idx['账号名称']] or '').strip()
        cl = str(row[idx['常驻地城市等级']] or '').strip()
        d['城市等级'] = {'一线新一线': '一线/新一线'}.get(cl, cl)
        kd = row[idx['开课日期']]
        d['开课日期'] = kd
        d['年级'] = str(row[idx['年级名称']] or '').strip()
        d['辅导姓名'] = str(row[idx.get('辅导姓名', 0)] or '').strip() if '辅导姓名' in idx else ''
        data.append(d)
    wb.close()
    return data, HAS_JW_LOCAL

def latest_periods(data, n=3):
    """最新 n 个期次，按开课日期排序；开课日期缺失的期次排最后"""
    kd = {}
    for d in data:
        kd.setdefault(d['期次'], d['开课日期'])
    def key(p):
        v = kd[p]
        try: return (0, v.toordinal()) if hasattr(v, 'toordinal') else (1, 0)
        except Exception: return (1, 0)
    ps = sorted(kd.keys(), key=key)
    return ps[-n:], ps[-1]

# ---------- 写出辅助 ----------
def write_header(ws, row, headers, start_col=1):
    for j, h in enumerate(headers):
        c = ws.cell(row=row, column=start_col + j, value=h)
        c.fill = GRAY; c.font = Font(bold=True); c.border = THIN
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def write_metric_row(ws, row, label, m, start_col=1, bold=False, fill=None):
    c = ws.cell(row=row, column=start_col, value=label)
    c.font = Font(bold=bold); c.border = THIN
    c.alignment = Alignment(vertical='center')
    if fill: c.fill = fill
    # 第2列: 在班人次; 第3列起: 14个比率指标
    c = ws.cell(row=row, column=start_col + 1, value=m.get('在班人次'))
    c.border = THIN
    c.alignment = Alignment(horizontal='center')
    if fill: c.fill = fill
    if bold: c.font = Font(bold=True)
    for j, (name, _) in enumerate(METRICS):
        v = m.get(name)
        c = ws.cell(row=row, column=start_col + 2 + j, value=v)
        c.number_format = '0.00%'; c.border = THIN
        c.alignment = Alignment(horizontal='center')
        if fill: c.fill = fill
        if bold: c.font = Font(bold=True)
    return row + 1

def finish_sheet(ws, hdr_row=1, first_data_row=2, last_row=None, last_col=None,
                 freeze='B2', pct_from_col=2, widths=None):
    """统一收尾：自动列宽、色阶（按列）、居中。pct_from_col 起的列为比率列"""
    last_row = last_row or ws.max_row
    last_col = last_col or ws.max_column
    if widths:
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
    for j in range(1, last_col + 1):
        L = get_column_letter(j)
        if not widths or L not in widths:
            mx = 8
            for i in range(1, min(ws.max_row, 80) + 1):
                v = ws.cell(row=i, column=j).value
                if v is not None:
                    mx = max(mx, min(len(str(v)) + 4, 22))
            ws.column_dimensions[L].width = mx
    # 比率列做色阶（跳过在班人次列，绝对量不做色阶）
    if pct_from_col:
        for j in range(pct_from_col, last_col + 1):
            L = get_column_letter(j)
            # 列内含数字才做
            has_num = any(isinstance(ws.cell(row=i, column=j).value, (int, float))
                          for i in range(first_data_row, last_row + 1))
            if has_num:
                color_scale(ws, f'{L}{first_data_row}:{L}{last_row}')
    if freeze:
        ws.freeze_panes = freeze

def color_scale(ws, rng):
    """红-黄-绿三色色阶（低=红，中=黄，高=绿），叠加不遮挡数值"""
    # 防御：范围不合法（如 G1:G0）时跳过
    try:
        from openpyxl.worksheet.cell_range import CellRange
        cr = CellRange(rng)
        if cr.max_row <= cr.min_row or cr.max_col <= 0:
            return
    except Exception:
        return
    ws.conditional_formatting.add(rng, ColorScaleRule(
        start_type='min', start_color='F8696B',
        mid_type='percentile', mid_value=50, mid_color='FFEB84',
        end_type='max', end_color='63BE7B'))

def color_scale_cols(ws, cols, row1, row2):
    """对若干列分别按列做色阶（不同列口径不同，不能整块一起做）"""
    for c in cols:
        L = get_column_letter(c)
        color_scale(ws, f'{L}{row1}:{L}{row2}')

# ---------- Sheet 实现 ----------
HEAD_FULL = ['在班人次'] + [m for m, _ in METRICS]

def s1_near3(wb, data, periods, latest):
    ws = wb.create_sheet('近三期全链路数据')
    write_header(ws, 1, ['名称'] + HEAD_FULL)
    r = 2
    by = defaultdict(list)
    for d in data:
        if d['期次'] in periods: by[d['期次']].append(d)
    ms = {}
    for p in periods:
        ms[p] = agg_metrics(by[p], HAS_JW)
        r = write_metric_row(ws, r, p, ms[p], bold=(p == latest))
    # 环比行
    prev = periods[-2] if len(periods) >= 2 else None
    if prev:
        row_vals = ['秋8环比秋7'.replace('秋8', latest).replace('秋7', prev)]
        # 通用
        row_vals = [f'{latest}环比{prev}']
        for name, _ in METRICS:
            a, b = ms[latest].get(name), ms[prev].get(name)
            row_vals.append(a - b if (a is not None and b is not None) else None)
        for j, v in enumerate(row_vals):
            c = ws.cell(row=r, column=1 + j, value=v)
            c.font = Font(bold=True, color='C00000'); c.fill = BLUE_TOTAL; c.border = THIN
            c.alignment = Alignment(horizontal='center', vertical='center')
            if j > 0: c.number_format = '0.00%'
    finish_sheet(ws, last_row=r, freeze='B2', pct_from_col=3)

def s2_team_periods(wb, data, periods):
    ws = wb.create_sheet('分团队期次数据')
    bases = []
    for d in data:
        if d['基地'] and d['基地'] not in bases: bases.append(d['基地'])
    write_header(ws, 1, ['期次'] + HEAD_FULL)
    r = 2
    base_rows = []   # (行号, 该基地首行, 该基地末行) 用于基地内色阶
    for b in bases:
        c = ws.cell(row=r, column=1, value=b)
        c.font = Font(bold=True, size=12); c.fill = BLUE_TOTAL
        r += 1
        first = r
        for p in periods:
            rows_ = [d for d in data if d['基地'] == b and d['期次'] == p]
            if not rows_: continue
            m = agg_metrics(rows_, HAS_JW)
            c = ws.cell(row=r, column=1, value=p); c.border = THIN
            c = ws.cell(row=r, column=2, value=m['在班人次']); c.border = THIN
            for j, (name, _) in enumerate(METRICS):
                c = ws.cell(row=r, column=3 + j, value=m.get(name))
                c.number_format = '0.00%'; c.border = THIN
            r += 1
        base_rows.append((first, r - 1))
        r += 1
    # 基地内色阶（各基地口径独立，不跨块）
    for f, t in base_rows:
        if t > f:
            color_scale_cols(ws, range(3, 3 + len(METRICS)), f, t)
    finish_sheet(ws, last_row=r - 1, freeze='B2', pct_from_col=None)

def s3_channel_latest(wb, data, latest):
    ws = wb.create_sheet('最新期次下渠道全链路数据')
    # 「渠道生源类别」列留空，由用户自行补充（2026-11-11 要求）
    # 底表无类别列，所有渠道按行排列，最后加总计
    write_header(ws, 1, ['渠道生源类别', '名称'] + HEAD_FULL)
    chans = sorted({d['账号名称'] for d in data if d['期次'] == latest and d['账号名称']})
    r = 2
    for ch in chans:
        rows_ = [d for d in data if d['期次'] == latest and d['账号名称'] == ch]
        if not rows_: continue
        # 渠道主行
        m = agg_metrics(rows_, HAS_JW)
        ws.cell(row=r, column=1, value=None).border = THIN  # 类别留空
        ws.cell(row=r, column=2, value=ch).border = THIN
        ws.cell(row=r, column=3, value=m['在班人次']).border = THIN
        for j, (name, _) in enumerate(METRICS):
            c = ws.cell(row=r, column=4 + j, value=m.get(name))
            c.number_format = '0.00%'; c.border = THIN
        r += 1
        # 年级交叉行（名称列缩进）
        r = write_grade_cross(ws, r, rows_, ch, 2, HAS_JW, indent=1)
    # 总计
    rows_ = [d for d in data if d['期次'] == latest]
    m = agg_metrics(rows_, HAS_JW)
    ws.cell(row=r, column=1, value='总计').font = Font(bold=True)
    ws.cell(row=r, column=2, value='').font = Font(bold=True)
    ws.cell(row=r, column=3, value=m['在班人次']).font = Font(bold=True)
    for j, (name, _) in enumerate(METRICS):
        c = ws.cell(row=r, column=4 + j, value=m.get(name))
        c.number_format = '0.00%'; c.font = Font(bold=True); c.fill = BLUE_TOTAL
    # 色阶只做渠道主行（年级交叉行口径不同，不做），范围逐行拼接不现实——
    # 简化：主行+总计行同口径，做逐列色阶排除交叉行不可行，改为整列色阶
    # 实际策略：对整列做色阶（交叉行值本身较小，不破坏排序直觉）
    last = r
    finish_sheet(ws, last_row=last, freeze='C2', pct_from_col=4,
                 widths={'A': 22, 'B': 14})

def s4_xkz(wb, data, latest):
    ws = wb.create_sheet('最新期次下学科组数据')
    write_header(ws, 1, ['名称'] + HEAD_FULL)
    xkz = sorted({d['学科组名称'] for d in data if d['期次'] == latest and d['学科组名称']})
    r = 2
    for x in xkz:
        rows_ = [d for d in data if d['期次'] == latest and d['学科组名称'] == x]
        m = agg_metrics(rows_, HAS_JW)
        c = ws.cell(row=r, column=1, value=x); c.border = THIN
        ws.cell(row=r, column=2, value=m['在班人次']).border = THIN
        for j, (name, _) in enumerate(METRICS):
            c = ws.cell(row=r, column=3 + j, value=m.get(name))
            c.number_format = '0.00%'; c.border = THIN
        r += 1
        # 年级交叉行（名称列缩进）
        r = write_grade_cross(ws, r, rows_, x, 1, HAS_JW, indent=1)
    finish_sheet(ws, last_row=r - 1, freeze='B2', pct_from_col=3, widths={'A': 14})

def s5_channel_share(wb, data, periods, latest):
    ws = wb.create_sheet('近三期渠道占比和转率')
    # 结构: C3 账号名称; 每期2列(占比,转率) + 环比2列 + 汇总2列
    chans = sorted({d['账号名称'] for d in data if d['期次'] in periods and d['账号名称']})
    # 统计
    stat = defaultdict(lambda: [0, 0])   # (期次,渠道) -> [在班, 转化]
    tot = defaultdict(lambda: [0, 0])    # 期次 -> [在班, 转化]
    for d in data:
        if d['期次'] not in periods: continue
        z = d['在班人次']; cv = d['转低人次'] + d['人次转化率']
        stat[(d['期次'], d['账号名称'])][0] += z
        stat[(d['期次'], d['账号名称'])][1] += cv
        tot[d['期次']][0] += z; tot[d['期次']][1] += cv
    # 表头两层
    write_header(ws, 1, ['', '', '账号名称'])
    c = ws.cell(row=3, column=3, value='账号名称')
    c.fill = GRAY; c.font = Font(bold=True); c.border = THIN
    col = 4
    for p in periods:
        ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 1)
        c = ws.cell(row=3, column=col, value=p)
        c.fill = GRAY; c.font = Font(bold=True); c.alignment = Alignment(horizontal='center'); c.border = THIN
        for off, sub in enumerate(['占比', '转率']):
            cc = ws.cell(row=4, column=col + off, value=sub)
            cc.fill = GRAY; cc.font = Font(bold=True); cc.border = THIN
        col += 2
    # 环比
    prev = periods[-2] if len(periods) >= 2 else None
    if prev:
        ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 1)
        c = ws.cell(row=3, column=col, value=f'{latest}环比{prev}')
        c.fill = GRAY; c.font = Font(bold=True); c.alignment = Alignment(horizontal='center'); c.border = THIN
        for off, sub in enumerate(['占比', '转率']):
            cc = ws.cell(row=4, column=col + off, value=sub)
            cc.fill = GRAY; cc.font = Font(bold=True); cc.border = THIN
        col += 2
    ws.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 1)
    c = ws.cell(row=3, column=col, value='汇总')
    c.fill = GRAY; c.font = Font(bold=True); c.alignment = Alignment(horizontal='center'); c.border = THIN
    for off, sub in enumerate(['占比', '转率']):
        cc = ws.cell(row=4, column=col + off, value=sub)
        cc.fill = GRAY; cc.font = Font(bold=True); cc.border = THIN
    col_all = col + 2
    r = 5
    sum_z = sum(tot[p][0] for p in periods); sum_c = sum(tot[p][1] for p in periods)
    for ch in chans:
        ws.cell(row=r, column=3, value=ch).border = THIN
        col = 4
        for p in periods:
            z, cv = stat[(p, ch)]
            ws.cell(row=r, column=col, value=rate(z, tot[p][0])).number_format = '0.00%'
            ws.cell(row=r, column=col + 1, value=rate(cv, z) if z else None).number_format = '0.00%'
            col += 2
        if prev:
            z1, _ = stat[(prev, ch)]; z2, _ = stat[(latest, ch)]
            ws.cell(row=r, column=col, value=rate(z2, tot[latest][0]) - rate(z1, tot[prev][0])).number_format = '0.00%'
            c1 = rate(stat[(prev, ch)][1], z1) if z1 else None
            c2 = rate(stat[(latest, ch)][1], z2) if z2 else None
            ws.cell(row=r, column=col + 1,
                    value=(c2 - c1) if (c1 is not None and c2 is not None) else None).number_format = '0.00%'
            col += 2
        ws.cell(row=r, column=col, value=rate(sum(stat[(p, ch)][0] for p in periods), sum_z)).number_format = '0.00%'
        ws.cell(row=r, column=col + 1,
                value=rate(sum(stat[(p, ch)][1] for p in periods),
                           sum(stat[(p, ch)][0] for p in periods))).number_format = '0.00%'
        r += 1
    # 总计行
    ws.cell(row=r, column=3, value='总计').font = Font(bold=True)
    col = 4
    for p in periods:
        ws.cell(row=r, column=col, value=1).number_format = '0.00%'
        ws.cell(row=r, column=col + 1, value=rate(tot[p][1], tot[p][0])).number_format = '0.00%'
        col += 2
    # 色阶：转率列（col+1）整列；占比列（col）不做
    rate_cols = [c for c in range(5, col_all, 2)]
    color_scale_cols(ws, rate_cols, 5, r - 1)
    finish_sheet(ws, hdr_row=1, first_data_row=5, last_row=r, last_col=col_all,
                 freeze='D5', pct_from_col=None, widths={'C': 16})

def s6_city_share(wb, data, periods, latest):
    ws = wb.create_sheet('近三期城市线占比和转率')
    cities = ['一线/新一线', '二线城市', '三线及以下']
    stat = defaultdict(lambda: [0, 0]); tot = defaultdict(lambda: [0, 0])
    for d in data:
        if d['期次'] not in periods: continue
        z = d['在班人次']; cv = d['转低人次'] + d['人次转化率']
        stat[(d['期次'], d['城市等级'])][0] += z
        stat[(d['期次'], d['城市等级'])][1] += cv
        tot[d['期次']][0] += z; tot[d['期次']][1] += cv
    write_header(ws, 1, ['常驻地城市等级'])
    col = 2
    for p in periods + ['汇总']:
        for off, sub in enumerate(['求和项:在班人次', '占比', '求和项:城市转率'] if p != '汇总'
                                  else ['在班人次', '占比', '城市转率']):
            c = ws.cell(row=2 if p != '汇总' else 2, column=col + off, value=p if off == 0 else sub)
            c.fill = GRAY; c.font = Font(bold=True); c.border = THIN
            c.alignment = Alignment(horizontal='center', wrap_text=True)
        # 一级表头合并
        if p != '汇总':
            ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
            c = ws.cell(row=1, column=col, value=p)
            c.fill = GRAY; c.font = Font(bold=True); c.alignment = Alignment(horizontal='center')
        col += 3
    r = 3
    sum_z = sum(tot[p][0] for p in periods)
    for city in cities:
        c = ws.cell(row=r, column=1, value=city); c.border = THIN
        col = 2
        for p in periods + ['汇总']:
            if p == '汇总':
                z = sum(stat[(pp, city)][0] for pp in periods)
                cv = sum(stat[(pp, city)][1] for pp in periods)
            else:
                z, cv = stat[(p, city)]
            tt = sum_z if p == '汇总' else tot[p][0]
            ws.cell(row=r, column=col, value=z).border = THIN
            ws.cell(row=r, column=col + 1, value=rate(z, tt)).number_format = '0.00%'
            ws.cell(row=r, column=col + 2, value=rate(cv, z) if z else None).number_format = '0.00%'
            col += 3
        r += 1
    # 总计
    c = ws.cell(row=r, column=1, value='总计'); c.font = Font(bold=True); c.fill = BLUE_TOTAL
    col = 2
    for p in periods + ['汇总']:
        tt = sum_z if p == '汇总' else tot[p][0]
        tv = sum_z and sum(tot[pp][1] for pp in periods) if p == '汇总' else tot[p][1]
        ws.cell(row=r, column=col, value=tt).font = Font(bold=True)
        ws.cell(row=r, column=col + 1, value=1).number_format = '0.00%'
        ws.cell(row=r, column=col + 2, value=rate(tv, tt)).number_format = '0.00%'
        col += 3
    ws.column_dimensions['A'].width = 16
    for j in range(2, col):
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = 11
    # 色阶：城市转率列（每块第3列）
    rate_cols = [c for c in range(4, col, 3)]
    color_scale_cols(ws, rate_cols, 3, r - 1)
    ws.freeze_panes = 'B3'

def s7_city_latest(wb, data, latest):
    ws = wb.create_sheet('最新期次下城市线全链数据')
    write_header(ws, 1, ['城市线'] + HEAD_FULL)
    cities = ['一线/新一线', '二线城市', '三线及以下']
    r = 2
    for city in cities:
        rows_ = [d for d in data if d['期次'] == latest and d['城市等级'] == city]
        if not rows_: continue
        m = agg_metrics(rows_, HAS_JW)
        c = ws.cell(row=r, column=1, value=city); c.border = THIN
        ws.cell(row=r, column=2, value=m['在班人次']).border = THIN
        for j, (name, _) in enumerate(METRICS):
            c = ws.cell(row=r, column=3 + j, value=m.get(name))
            c.number_format = '0.00%'; c.border = THIN
        r += 1
    finish_sheet(ws, last_row=r - 1, freeze='B2', pct_from_col=3, widths={'A': 14})

def s8_team_channel(wb, data, latest):
    ws = wb.create_sheet('最新期次下分团队各渠道占比和转率')
    bases = []
    for d in data:
        if d['基地'] and d['基地'] not in bases: bases.append(d['基地'])
    chans = sorted({d['账号名称'] for d in data if d['期次'] == latest and d['账号名称']})
    # 两层表头
    col = 1
    c = ws.cell(row=1, column=2, value='账号名称'); c.fill = GRAY; c.font = Font(bold=True); c.border = THIN
    col = 3
    for b in bases:
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 1)
        c = ws.cell(row=1, column=col, value=b)
        c.fill = GRAY; c.font = Font(bold=True); c.alignment = Alignment(horizontal='center'); c.border = THIN
        for off, sub in enumerate(['求和项:在班人次', '求和项:转率']):
            cc = ws.cell(row=2, column=col + off, value=sub)
            cc.fill = GRAY; cc.font = Font(bold=True); cc.border = THIN
        col += 2
    # 汇总两列
    for off, sub in enumerate(['求和项:在班人次汇总', '求和项:转率汇总']):
        cc = ws.cell(row=1 if off == 0 else 1, column=col + off, value=sub)
        cc.fill = GRAY; cc.font = Font(bold=True); cc.border = THIN
        cc.alignment = Alignment(horizontal='center', wrap_text=True)
    # 数据
    stat = defaultdict(lambda: [0, 0]); btot = defaultdict(lambda: [0, 0])
    allz = allcv = 0
    for d in data:
        if d['期次'] != latest: continue
        z = d['在班人次']; cv = d['转低人次'] + d['人次转化率']
        stat[(d['账号名称'], d['基地'])][0] += z
        stat[(d['账号名称'], d['基地'])][1] += cv
        btot[d['基地']][0] += z; btot[d['基地']][1] += cv
        allz += z; allcv += cv
    r = 3
    for ch in chans:
        c = ws.cell(row=r, column=2, value=ch); c.border = THIN
        col = 3
        for b in bases:
            z, cv = stat[(ch, b)]
            ws.cell(row=r, column=col, value=rate(z, btot[b][0])).number_format = '0.00%'
            ws.cell(row=r, column=col + 1, value=rate(cv, z) if z else None).number_format = '0.00%'
            col += 2
        z = sum(stat[(ch, b)][0] for b in bases); cv = sum(stat[(ch, b)][1] for b in bases)
        ws.cell(row=r, column=col, value=rate(z, allz)).number_format = '0.00%'
        ws.cell(row=r, column=col + 1, value=rate(cv, z) if z else None).number_format = '0.00%'
        r += 1
    # 总计
    c = ws.cell(row=r, column=2, value='总计'); c.font = Font(bold=True); c.fill = BLUE_TOTAL
    col = 3
    for b in bases:
        ws.cell(row=r, column=col, value=1).number_format = '0.00%'
        ws.cell(row=r, column=col + 1, value=rate(btot[b][1], btot[b][0])).number_format = '0.00%'
        col += 2
    ws.cell(row=r, column=col, value=1).number_format = '0.00%'
    ws.cell(row=r, column=col + 1, value=rate(allcv, allz)).number_format = '0.00%'
    # 色阶：转率列（每块第2列 + 汇总转率）
    rate_cols = [c for c in range(4, col + 2, 2)]
    color_scale_cols(ws, rate_cols, 3, r - 1)
    finish_sheet(ws, first_data_row=3, last_row=r, last_col=col + 1,
                 freeze='C3', pct_from_col=None, widths={'B': 14})

def s9_phone(wb, data, latest):
    ws = wb.create_sheet('最新期次分团队电话数据')
    write_header(ws, 1, ['', '学科组名称', '求和项:OA电话覆盖率', '求和项:企微电话覆盖率',
                        '求和项:OA电话接通率1', '求和项:企微电话接通率1'], start_col=2)
    xkz = sorted({d['学科组名称'] for d in data if d['期次'] == latest and d['学科组名称']})
    r = 2
    for x in xkz:
        rows_ = [d for d in data if d['期次'] == latest and d['学科组名称'] == x]
        m = agg_metrics(rows_, HAS_JW)
        c = ws.cell(row=r, column=3, value=x); c.border = THIN
        for j, k in enumerate(['OA电话覆盖率', '企微电话覆盖率', 'OA电话接通率1', '企微电话接通率1']):
            c = ws.cell(row=r, column=4 + j, value=m[k])
            c.number_format = '0.00%'; c.border = THIN
        r += 1
        # 年级交叉行
        for g in grades_of(rows_):
            gr = [d for d in rows_ if d['年级'] == g]
            if not gr: continue
            mg = agg_metrics(gr, HAS_JW)
            c = ws.cell(row=r, column=3, value='  ' + g); c.border = THIN
            for j, k in enumerate(['OA电话覆盖率', '企微电话覆盖率', 'OA电话接通率1', '企微电话接通率1']):
                c = ws.cell(row=r, column=4 + j, value=mg[k])
                c.number_format = '0.00%'; c.border = THIN
            r += 1
    for col in ['C', 'D', 'E', 'F', 'G']:
        ws.column_dimensions[col].width = 20
    # 色阶：4个比率列（D-G）
    color_scale_cols(ws, range(4, 8), 2, r - 1)
    ws.freeze_panes = 'D2'

def s10_fudao(wb, data, latest):
    """扁平排名表（2026-11-11 按用户截图改版）:
    辅导老师 × 学科组 × 年级 一行，按转化率降序排名；299绩效等级列留空由用户填写；底部总计行"""
    ws = wb.create_sheet('最新期次辅导数据')
    write_header(ws, 1, ['辅导老师', '学科组（组长）', '年级', '转化率排名', '299绩效等级',
                         '在班人次', '转化率', '出勤转化率', 'APP登录率',
                         '课前激活率', '课前破冰率', '课前深沟率', '摸底测覆盖率',
                         '加过微率', '任一讲出勤率', '一讲出勤率', '二讲出勤率',
                         '三讲出勤率', '四讲出勤率', '五讲出勤率'])
    rows_ = [d for d in data if d['期次'] == latest and d['辅导姓名']]
    groups = defaultdict(list)
    for d in rows_:
        groups[(d['辅导姓名'], d['学科组名称'], d['年级'])].append(d)
    # 排序：按辅导级聚合转化率降序（同辅导的年级行跟在其后），行内年级再按辅导内转化率降序
    def cvrate(rs):
        z = sum(x['在班人次'] for x in rs)
        return (sum(x['转低人次'] for x in rs) / z) if z else -1
    fuds = defaultdict(list)
    for (f, x, g), rs in groups.items():
        fuds[f].append((x, g, rs))
    # 辅导级排名（按该辅导全部年级合计转化率）
    fud_rank = sorted(fuds.keys(),
                      key=lambda f: -cvrate([d for _, _, rs in fuds[f] for d in rs]))
    rank = 0
    r = 2
    for f in fud_rank:
        rank += 1
        # 年级行按转化率降序
        items = sorted(fuds[f], key=lambda t: -cvrate(t[2]))
        for x, g, rs in items:
            m = agg_metrics(rs, HAS_JW)
            vals = [f, x, g, rank, None, m['在班人次'], m['转化率'], m['出勤转化率'],
                    m['APP登录率'], m['课前激活率'], m['课前破冰率'], m['课前深沟率'],
                    m['摸底测覆盖率'], m['加过微率'],
                    m['一讲出勤率'], m['一讲出勤率'], m['二讲出勤率'], m['三讲出勤率'],
                    m['四讲出勤率'], rate(sum(z['五讲出勤人数'] for z in rs), m['在班人次'])]
            # 注意第15、16列：任一讲出勤率 与 一讲出勤率 分列
            vals[14] = m.get('任一讲出勤率', rate(sum(z['任一讲出勤人数'] for z in rs), m['在班人次']))
            for j, v in enumerate(vals, start=1):
                c = ws.cell(row=r, column=j, value=v); c.border = THIN
                if j >= 7:
                    c.number_format = '0.00%'
                    c.alignment = Alignment(horizontal='center')
                if j <= 3:
                    c.alignment = Alignment(horizontal='center')
            r += 1
    last_data = r - 1
    # 总计行
    m = agg_metrics(rows_, HAS_JW)
    vals = ['总计', None, None, None, None, m['在班人次'], m['转化率'], m['出勤转化率'],
            m['APP登录率'], m['课前激活率'], m['课前破冰率'], m['课前深沟率'],
            m['摸底测覆盖率'], m['加过微率'], m['一讲出勤率'], m['一讲出勤率'],
            m['二讲出勤率'], m['三讲出勤率'], m['四讲出勤率'],
            rate(sum(z['五讲出勤人数'] for z in rows_), m['在班人次'])]
    vals[14] = rate(sum(z['任一讲出勤人数'] for z in rows_), m['在班人次'])
    for j, v in enumerate(vals, start=1):
        c = ws.cell(row=r, column=j, value=v)
        c.font = Font(bold=True); c.fill = BLUE_TOTAL; c.border = THIN
        c.alignment = Alignment(horizontal='center')
        if j >= 7:
            c.number_format = '0.00%'
    # 色阶：所有比率列（G-T，含转化率/出勤率等），按列独立做
    color_scale_cols(ws, range(7, 21), 2, last_data)
    widths = {'A': 10, 'B': 16, 'C': 8, 'D': 11, 'E': 12, 'F': 10}
    for j in range(6, 21):
        L = get_column_letter(j)
        if L not in widths:
            widths[L] = 11
    finish_sheet(ws, last_row=r, last_col=20, freeze='F2', pct_from_col=None,
                 widths=widths)

def _cv_fill(v):
    """转化率色阶：>=3% 深绿，2-3% 浅绿，1-2% 浅橙，<1% 浅红（参考截图配色，简单叠加）"""
    if v is None: return None
    if v >= 0.03: return PatternFill('solid', fgColor='63BE7B')
    if v >= 0.02: return PatternFill('solid', fgColor='A9D08E')
    if v >= 0.01: return PatternFill('solid', fgColor='FFD966')
    return PatternFill('solid', fgColor='F4A6A6')

# ---------- 年级交叉辅助 ----------
GRADES = ['一年级', '二年级', '三年级', '四年级', '五年级', '六年级']

def grades_of(rows_):
    """返回该数据子集中实际出现的年级，按 GRADES 顺序，其他值排最后"""
    gs = {d['年级'] for d in rows_ if d['年级']}
    ordered = [g for g in GRADES if g in gs]
    other = sorted(gs - set(GRADES))
    return ordered + other

def write_grade_cross(ws, r, group_rows, dim_name, label_col, has_jw, indent=1):
    """对一个维度分组按年级交叉展开：每个年级一行 + 小计行。
    label_col: 名称列(1-based)。返回下一个可用行号。"""
    ind = '  ' * indent
    for g in grades_of(group_rows):
        gr = [d for d in group_rows if d['年级'] == g]
        if not gr: continue
        m = agg_metrics(gr, has_jw)
        c = ws.cell(row=r, column=label_col, value=ind + g); c.border = THIN
        c.font = Font(color='595959')
        c = ws.cell(row=r, column=label_col + 1, value=m['在班人次']); c.border = THIN
        for j, (name, _) in enumerate(METRICS):
            c = ws.cell(row=r, column=label_col + 2 + j, value=m.get(name))
            c.number_format = '0.00%'; c.border = THIN
        r += 1
    # 小计（该维度组的合计）
    m = agg_metrics(group_rows, has_jw)
    c = ws.cell(row=r, column=label_col, value=ind + dim_name + '小计')
    c.font = Font(bold=True); c.border = THIN
    c = ws.cell(row=r, column=label_col + 1, value=m['在班人次'])
    c.font = Font(bold=True); c.border = THIN
    for j, (name, _) in enumerate(METRICS):
        c = ws.cell(row=r, column=label_col + 2 + j, value=m.get(name))
        c.number_format = '0.00%'; c.font = Font(bold=True); c.fill = BLUE_TOTAL; c.border = THIN
    return r + 1

# ---------- 主流程 ----------
def suppress_excel_error_checks(path):
    """给所有 sheet 注入 <ignoredErrors>，关闭 Excel 的错误检查提示（绿三角/黄色感叹号）"""
    import zipfile, shutil, re
    INJ = ('<ignoredErrors><ignoredError sqref="A1:XFD1048576" '
           'numberStoredAsText="1" evalError="1"/></ignoredErrors>')
    tmp = path + '.tmp'
    with zipfile.ZipFile(path) as zin, \
         zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.match(r'xl/worksheets/sheet\d+\.xml$', item.filename):
                x = data.decode('utf-8')
                if '<ignoredErrors>' not in x:
                    x = x.replace('</worksheet>', INJ + '</worksheet>')
                    data = x.encode('utf-8')
            zout.writestr(item, data)
    shutil.move(tmp, path)

def main():
    src = sys.argv[1] if len(sys.argv) > 1 else '渠道城市线底表-秋8进校1110-处理后-v3-4.xlsx'
    out = sys.argv[2] if len(sys.argv) > 2 else '进校-每日数据更新-自动生成.xlsx'
    if not os.path.isabs(src): src = os.path.abspath(src)
    if not os.path.isabs(out): out = os.path.abspath(out)
    print(f'读取底表: {src}')
    global HAS_JW
    data, HAS_JW = load(src)
    print(f'共 {len(data)} 行')
    periods, latest = latest_periods(data, 3)
    print(f'近三期: {periods}，最新期次: {latest}')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    s1_near3(wb, data, periods, latest)
    s2_team_periods(wb, data, periods)
    s3_channel_latest(wb, data, latest)
    s4_xkz(wb, data, latest)
    s5_channel_share(wb, data, periods, latest)
    s6_city_share(wb, data, periods, latest)
    s7_city_latest(wb, data, latest)
    s8_team_channel(wb, data, latest)
    s9_phone(wb, data, latest)
    s10_fudao(wb, data, latest)
    wb.save(out)

    # openpyxl 不支持序列化 ignoredErrors，用 zipfile 后处理注入 XML
    suppress_excel_error_checks(out)
    print(f'已生成: {out}')
    # 校验
    chk = openpyxl.load_workbook(out)
    print('sheets:', chk.sheetnames)

if __name__ == '__main__':
    main()
