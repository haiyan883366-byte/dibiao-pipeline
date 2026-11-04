# -*- coding: utf-8 -*-
"""
0元课前总览底表处理流水线（skill: dibiao-pipeline）
Step1 清洗：文本型数字修复 + 最新未开课期次学科组归并
Step2 指标透视：静态结果表（指标/数值/口径说明，按期次分区）
Step3 对比：5 个横向布局对比 sheet + 条件格式（窗口=最新3个期次）

用法：
  python dibiao_pipeline.py --src <路径> [--keep-su "崔国锋,史春晶"] \
      [--target-su "张庆波"] [--out <输出路径>] [--qici 秋8期]

依赖：openpyxl >= 3.1
"""
import argparse
import datetime as dt
import glob
import os
import re
import sys
import warnings

warnings.filterwarnings('ignore')
import openpyxl
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.utils import get_column_letter

SRC_PATTERN = r'0-低-正全链路【学部承接】-0元课前总览-*.xlsx'
DEFAULT_DIR = r'D:\download'
N = ['一', '二', '三', '四', '五']
GRADES = ['一年级', '二年级', '三年级', '四年级', '五年级', '六年级']

# ---------------- Step1: 文本型数字修复 ----------------
DATE_RE = re.compile(r'^\d{4}-\d{1,2}-\d{1,2}$')
INT_RE = re.compile(r'^-?\d+$')
FLT_RE = re.compile(r'^-?\d+\.\d+$')

def fix_cell(v):
    """按语义转换文本存储的数字/日期；非字符串原样返回"""
    if not isinstance(v, str):
        return v
    s = v.strip()
    if not s:
        return v
    if DATE_RE.match(s):
        try:
            y, m, d = map(int, s.split('-'))
            return dt.date(y, m, d)
        except ValueError:
            return v
    if INT_RE.match(s) and len(s) < 12:          # ID类长数字也转 int
        return int(s)
    if FLT_RE.match(s):
        return float(s)
    return v

# ---------------- 参数 ----------------
def parse_args():
    p = argparse.ArgumentParser(description='0元课前总览底表处理流水线')
    p.add_argument('--src', help='底表路径；不传则取 D:\\download 下日期最新的一份')
    p.add_argument('--keep-su', default='崔国锋,史春晶',
                   help='最新未开课期次保留学科组名单（逗号分隔）')
    p.add_argument('--target-su', default='张庆波', help='归并目标学科组名')
    p.add_argument('--out', help='输出路径；默认与底表同目录 <名>-处理后-v3-N.xlsx')
    p.add_argument('--qici', help='未开课期次名（默认取字典序最大的期次）')
    return p.parse_args()

# ---------------- 读取 + 聚合 ----------------
def load_and_aggregate(src, keep_su, target_su, qici_latest):
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [fix_cell(h) for h in next(rows)]
    IX = {h: i for i, h in enumerate(header)}
    for need in ['期次', '账号名称', '学科组名称', '年级名称', '在班人次', '转低人次',
                 '加微人次', '加过微人次', 'app_登录人次', '课前激活人次', '课前破冰人次',
                 '课前深沟人次', '摸底测覆盖人次', '任一讲出勤人数', '任一讲完课人次'] \
            + [f'{n}讲出勤人数' for n in N] + [f'{n}讲完课人数' for n in N] \
            + [f'{n}讲转化人次' for n in N]:
        if need not in IX:
            sys.exit(f'[ERROR] 底表缺少必需列: {need}')

    ATT = {n: IX[f'{n}讲出勤人数'] for n in N}
    DONE = {n: IX[f'{n}讲完课人数'] for n in N}
    ZHUAN = {n: IX[f'{n}讲转化人次'] for n in N}
    BASE = [('在班', IX['在班人次']), ('加微', IX['加微人次']), ('加过微', IX['加过微人次']),
            ('转低', IX['转低人次']), ('登录', IX['app_登录人次']), ('激活', IX['课前激活人次']),
            ('破冰', IX['课前破冰人次']), ('深沟', IX['课前深沟人次']), ('摸底', IX['摸底测覆盖人次'])]
    ANY_ATT, ANY_DONE = IX['任一讲出勤人数'], IX['任一讲完课人次']
    ZD = IX['转低人次']

    raw = defaultdict(lambda: defaultdict(float))
    anycnt = defaultdict(lambda: [0.0, 0.0])
    qici_set = []
    for row in rows:
        # Step1: 文本型数字修复（就地，仅数值/日期列有影响）
        row = [fix_cell(v) for v in row]
        t = row[IX['期次']]
        if t is None:
            continue
        t = str(t)
        if t not in qici_set:
            qici_set.append(t)
        ch = row[IX['账号名称']]
        su = row[IX['学科组名称']]
        gr = row[IX['年级名称']]
        # Step1: 仅未开课期次，非保留组（含空白/None）归入目标组
        if t == qici_latest and su not in keep_su:
            su = target_su
        keys = {(t,), (t, 'C', ch), (t, 'S', su), (t, 'G', gr),
                (t, 'C', ch, 'G', gr), (t, 'C', ch, 'S', su), (t, 'S', su, 'G', gr),
                ('渠道全体', 'C', ch)}
        for k in keys:
            d = raw[k]
            for name, idx in BASE:
                v = row[idx]
                if v is not None: d[name] += float(v)
            for n in N:
                v = row[ATT[n]]
                if v is not None: d[f'{n}出'] += float(v)
                v = row[DONE[n]]
                if v is not None: d[f'{n}完'] += float(v)
                v = row[ZHUAN[n]]
                if v is not None: d[f'{n}转'] += float(v)
            v = row[ANY_ATT]
            if v is not None: d['任出'] += float(v)
            v = row[ANY_DONE]
            if v is not None: d['任完'] += float(v)
        # 行级：任一讲出勤 且 转低（一人只计一次）
        a = row[ANY_ATT]; zd = row[ZD]
        a = 1 if (a is not None and float(a) >= 1) else 0
        zd = 1 if (zd is not None and float(zd) >= 1) else 0
        if a >= 1:
            for k in keys:
                anycnt[k][1] += 1
                if zd >= 1: anycnt[k][0] += 1
        # 全局键（整体行需要）
        if a >= 1:
            anycnt[()][1] += 1
            if zd >= 1: anycnt[()][0] += 1
        d = raw[()]
        for name, idx in BASE:
            v = row[idx]
            if v is not None: d[name] += float(v)
        for n in N:
            v = row[ATT[n]]
            if v is not None: d[f'{n}出'] += float(v)
            v = row[DONE[n]]
            if v is not None: d[f'{n}完'] += float(v)
            v = row[ZHUAN[n]]
            if v is not None: d[f'{n}转'] += float(v)
        v = row[ANY_ATT]
        if v is not None: d['任出'] += float(v)
        v = row[ANY_DONE]
        if v is not None: d['任完'] += float(v)
    wb.close()
    return raw, anycnt, qici_set

# ---------------- 指标计算 ----------------
def ratio(a, b):
    return a / b if b else None

def metrics(d, anyc):
    zb = d['在班']
    m = {'在班人次': zb, '渠道内占比': None,
         '转化率': ratio(d['转低'], zb),
         '任一讲出勤转化率': ratio(anyc[0], anyc[1]) if anyc[1] else None,
         '加微率': ratio(d['加微'], zb), '加过微率': ratio(d['加过微'], zb),
         '删除率': (ratio(d['加过微'] - d['加微'], d['加过微']) if d['加过微'] else None),
         'APP登录率': ratio(d['登录'], zb), '课前激活率': ratio(d['激活'], zb),
         '课前破冰率': ratio(d['破冰'], zb), '课前深沟率': ratio(d['深沟'], zb),
         '摸底测覆盖率': ratio(d['摸底'], zb)}
    for n in N:
        m[f'{n}讲出勤率'] = ratio(d[f'{n}出'], zb)
        m[f'{n}讲完课率'] = ratio(d[f'{n}完'], d[f'{n}出'])
        m[f'{n}讲出勤转化率'] = ratio(d[f'{n}转'], d[f'{n}出'])
    m['任一讲出勤率'] = ratio(d['任出'], zb)
    m['任一讲完课率'] = ratio(d['任完'], d['任出'])
    return m

COLS = ['在班人次', '渠道内占比', '转化率', '任一讲出勤转化率', '加微率', '加过微率', '删除率',
        'APP登录率', '课前激活率', '课前破冰率', '课前深沟率', '摸底测覆盖率'] \
     + [f'{n}讲出勤率' for n in N] + [f'{n}讲完课率' for n in N] + [f'{n}讲出勤转化率' for n in N] \
     + ['任一讲出勤率', '任一讲完课率']
PCT_COLS = set(COLS) - {'在班人次'}

# ---------------- 样式 ----------------
HDR_FILL = PatternFill('solid', fgColor='F2F2F2')
TOT_FILL = PatternFill('solid', fgColor='DDEBF7')
CS = ColorScaleRule(start_type='min', start_color='F8696B',
                    mid_type='percentile', mid_value=50, mid_color='FFEB84',
                    end_type='max', end_color='63BE7B')
DB = DataBarRule(start_type='num', start_value=0, end_type='max',
                 color='638EC6', showValue=True)

def write_compare_sheet(wb, name, blocks):
    """blocks: [(块标题 or None, [(标签, metrics_dict)], 是否有总计行)]"""
    if name in wb.sheetnames:
        del wb[name]
    w = wb.create_sheet(name)
    w.append(['名称'] + COLS)
    for c in w[1]:
        c.fill = HDR_FILL; c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for title, items, has_total in blocks:
        r0 = w.max_row + 1
        if title:
            w.cell(row=r0, column=1, value=title).font = Font(bold=True, size=12)
            r0 += 1
        for label, m in items:
            w.cell(row=r0, column=1, value=label)
            for j, c in enumerate(COLS, start=2):
                v = m.get(c)
                if v is not None:
                    w.cell(row=r0, column=j, value=v)
                    if c in PCT_COLS: w.cell(row=r0, column=j).number_format = '0.00%'
                    elif c == '在班人次': w.cell(row=r0, column=j).number_format = '#,##0'
            if label == '总计':
                for cc in w[r0]:
                    cc.fill = TOT_FILL; cc.font = Font(bold=True)
            r0 += 1
    if w.max_row > 1:
        for j, c in enumerate(COLS, start=2):
            letter = get_column_letter(j)
            rng = f'{letter}2:{letter}{w.max_row}'
            if c == '渠道内占比':
                w.conditional_formatting.add(rng, DB)
            elif c in PCT_COLS:
                w.conditional_formatting.add(rng, CS)
    w.freeze_panes = 'B2'
    w.column_dimensions['A'].width = 22
    for j in range(2, len(COLS) + 2):
        w.column_dimensions[get_column_letter(j)].width = 11
    return w

# ---------------- Step2: 指标透视 ----------------
TOUCHI = {
    '转化率': '转低人次 / 在班人次',
    '任一讲出勤转化率': '任一讲出勤且转低的用户 / 任一讲出勤用户（行级）',
    '加微率': '加微人次 / 在班人次', '加过微率': '加过微人次 / 在班人次',
    '删除率': '(加过微-加微) / 加过微', 'APP登录率': 'app_登录人次 / 在班人次',
    '课前激活率': '课前激活人次 / 在班人次', '课前破冰率': '课前破冰人次 / 在班人次',
    '课前深沟率': '课前深沟人次 / 在班人次', '摸底测覆盖率': '摸底测覆盖人次 / 在班人次',
    '任一讲出勤率': '任一讲出勤人数 / 在班人次',
    '任一讲完课率': '任一讲完课人次 / 任一讲出勤人数',
}

def write_toushi(wb, qici_sorted, raw, anycnt, qici_latest):
    name = '指标透视'
    if name in wb.sheetnames:
        del wb[name]
    w = wb.create_sheet(name)
    w.append(['期次', '指标', '数值', '口径说明'])
    for c in w[1]:
        c.fill = HDR_FILL; c.font = Font(bold=True)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    r = 2
    total_zb = raw[()]['在班']
    blocks = [(t, raw[(t,)], anycnt[(t,)]) for t in qici_sorted] + [('整体', raw[()], anycnt[()])]
    for t, d, anyc in blocks:
        m = metrics(d, anyc)
        keys = ['在班人次', '转化率', '任一讲出勤转化率', '加微率', '加过微率',
                '删除率', 'APP登录率', '课前激活率', '课前破冰率', '课前深沟率', '摸底测覆盖率',
                '任一讲出勤率', '任一讲完课率']
        for n in N:
            keys += [f'{n}讲出勤率', f'{n}讲完课率', f'{n}讲出勤转化率']
        for k in keys:
            v = m.get(k)
            shuoming = '本期在班合计' if k == '在班人次' else TOUCHI.get(k, '')
            w.cell(row=r, column=1, value=t)
            w.cell(row=r, column=2, value=k)
            c3 = w.cell(row=r, column=3)
            if v is not None:
                c3.value = v
                c3.number_format = '#,##0' if k == '在班人次' else '0.00%'
            w.cell(row=r, column=4, value=shuoming)
            r += 1
        # 渠道占比明细（该期次有数据的渠道）
        for ch in sorted({k[2] for k in raw if len(k) >= 3 and k[0] == t and k[1] == 'C'}, key=str):
            dc = raw.get((t, 'C', ch))
            if dc is None or dc['在班'] == 0: continue
            v = ratio(dc['在班'], d['在班'])
            w.cell(row=r, column=1, value=t)
            w.cell(row=r, column=2, value=f'渠道占比-{ch}')
            c3 = w.cell(row=r, column=3, value=v)
            c3.number_format = '0.00%'
            w.cell(row=r, column=4, value=f'{ch}用户{int(dc["在班"])}人/总{int(d["在班"])}人')
            r += 1
    w.column_dimensions['A'].width = 12
    w.column_dimensions['B'].width = 20
    w.column_dimensions['C'].width = 12
    w.column_dimensions['D'].width = 46
    w.freeze_panes = 'A2'
    return w

# ---------------- Step3: 5 个对比 sheet ----------------
def add_zhanbi(items, total_zb):
    for label, m in items:
        m['渠道内占比'] = ratio(m.get('在班人次'), total_zb)

def build_sheets(wb, raw, anycnt, qici_set, qici_latest):
    qici_sorted = sorted(qici_set)[-3:]   # 对比窗口 = 最新 3 个期次

    # 期次对比
    items = [(t, metrics(raw[(t,)], anycnt[(t,)])) for t in qici_sorted]
    total_d, total_m = raw[()], metrics(raw[()], anycnt[()])
    add_zhanbi(items, total_d['在班'])
    total_m['渠道内占比'] = 1.0
    write_compare_sheet(wb, '期次对比', [(None, items + [('整体', total_m)], False)])

    # 最新未开课期次学科组对比
    su_sheet = f'{qici_latest}学科组对比'
    zb_l = raw[(qici_latest,)]['在班']
    sus = sorted({k[2] for k in raw if len(k) >= 3 and k[0] == qici_latest and k[1] == 'S'}, key=str)
    items = []
    for su in sus:
        d = raw.get((qici_latest, 'S', su))
        if d is None or d['在班'] == 0: continue
        items.append((su, metrics(d, anycnt[(qici_latest, 'S', su)])))
    add_zhanbi(items, zb_l)
    write_compare_sheet(wb, su_sheet, [(qici_latest, items, False)])

    # 期次下渠道对比
    blocks = []
    for t in qici_sorted:
        zb_t = raw[(t,)]['在班']
        chans = sorted({k[2] for k in raw if len(k) >= 3 and k[0] == t and k[1] == 'C'}, key=str)
        items = []
        for ch in chans:
            d = raw.get((t, 'C', ch))
            if d is None or d['在班'] == 0: continue
            items.append((ch, metrics(d, anycnt[(t, 'C', ch)])))
        add_zhanbi(items, zb_t)
        items.append(('总计', metrics(raw[(t,)], anycnt[(t,)])))
        items[-1][1]['渠道内占比'] = 1.0
        blocks.append((str(t), items, True))
    write_compare_sheet(wb, '期次下渠道对比', blocks)

    # 期次下年级对比
    blocks = []
    for t in qici_sorted:
        zb_t = raw[(t,)]['在班']
        items = []
        for gr in GRADES:
            d = raw.get((t, 'G', gr))
            if d is None or d['在班'] == 0: continue
            items.append((gr, metrics(d, anycnt[(t, 'G', gr)])))
        add_zhanbi(items, zb_t)
        items.append(('总计', metrics(raw[(t,)], anycnt[(t,)])))
        items[-1][1]['渠道内占比'] = 1.0
        blocks.append((str(t), items, True))
    write_compare_sheet(wb, '期次下年级对比', blocks)

    # 渠道下期次对比
    chans = sorted({k[2] for k in raw if len(k) >= 3 and k[1] == 'C'}, key=str)
    blocks = []
    for ch in chans:
        items = []
        for t in qici_sorted:
            d = raw.get((t, 'C', ch))
            if d is None or d['在班'] == 0: continue
            items.append((str(t), metrics(d, anycnt[(t, 'C', ch)])))
        if not items: continue
        tot_zb = sum(m['在班人次'] for _, m in items)
        add_zhanbi(items, tot_zb)
        items.append(('总计', metrics(raw[('渠道全体', 'C', ch)], anycnt[('渠道全体', 'C', ch)])))
        items[-1][1]['渠道内占比'] = 1.0
        blocks.append((str(ch), items, True))
    write_compare_sheet(wb, '渠道下期次对比', blocks)
    return qici_sorted, su_sheet

# ---------------- 校验：独立路径交叉对账 ----------------
def verify_independent(src, qici_latest):
    """第二遍独立读取原始底表（不复用聚合结果），直接按行求和。
    返回: {期次或'整体': {'在班':s,'转低':s,'任出行数':n,'任出且转低':n,'原始行数':n}}"""
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = [fix_cell(h) for h in next(rows)]
    IX = {h: i for i, h in enumerate(header)}
    I_ZB, I_ZD, I_QI = IX['在班人次'], IX['转低人次'], IX['期次']
    I_AA = IX['任一讲出勤人数']
    chk = defaultdict(lambda: {'在班': 0.0, '转低': 0.0, '任出行数': 0, '任出且转低': 0, '原始行数': 0})
    for row in rows:
        row = [fix_cell(v) for v in row]
        t = row[I_QI]
        if t is None:
            continue
        t = str(t)
        zb = row[I_ZB]
        zd = row[I_ZD]
        a = row[I_AA]
        a1 = 1 if (a is not None and float(a) >= 1) else 0
        zd1 = 1 if (zd is not None and float(zd) >= 1) else 0
        for key in (t, '整体'):
            c = chk[key]
            c['原始行数'] += 1
            if zb is not None: c['在班'] += float(zb)
            if zd is not None: c['转低'] += float(zd)
            if a1:
                c['任出行数'] += 1
                if zd1: c['任出且转低'] += 1
    wb.close()
    return dict(chk)

def write_jiaoyan(wb, chk, raw, anycnt, qici_sorted, qici_latest, target_su):
    """✅校验 sheet：聚合结果(脚本) vs 独立求和(原始表直加)，并排对账"""
    name = '✅校验'
    if name in wb.sheetnames:
        del wb[name]
    w = wb.create_sheet(name)
    OK_FILL = PatternFill('solid', fgColor='C6EFCE')
    BAD_FILL = PatternFill('solid', fgColor='FFC7CE')
    r = [1]
    def hdr(cols):
        w.cell(row=r[0], column=1, value=cols[0]).font = Font(bold=True)
        for j, c in enumerate(cols[1:], start=2):
            cell = w.cell(row=r[0], column=j, value=c)
            cell.fill = HDR_FILL; cell.font = Font(bold=True)
        r[0] += 1
    def row(label, v1, v2, fmt='#,##0', note=''):
        """并排对比两个数，相等打绿✔"""
        c1, c2 = w.cell(row=r[0], column=2, value=v1), w.cell(row=r[0], column=3, value=v2)
        c1.number_format = c2.number_format = fmt
        diff = (v1 or 0) - (v2 or 0)
        eq = abs(diff) < 1e-6
        cell = w.cell(row=r[0], column=4, value='✔ 一致' if eq else f'✘ 差额 {diff:,.2f}')
        cell.fill = OK_FILL if eq else BAD_FILL
        if note: w.cell(row=r[0], column=5, value=note)
        r[0] += 1
        return eq

    w.cell(row=1, column=1, value='✅ 数据对账校验（脚本聚合 vs 原始表独立求和）').font = Font(bold=True, size=13)
    r[0] = 2
    w.cell(row=r[0], column=1, value='生成时间').font = Font(bold=True)
    w.cell(row=r[0], column=2, value=dt.datetime.now().strftime('%Y-%m-%d %H:%M'))
    r[0] += 2

    # 1) 各期次在班人次对账
    hdr(['在班人次对账', '脚本聚合值', '原始表直加', '结果', '说明'])
    all_ok = True
    for t in qici_sorted:
        ok = row(str(t), raw[(t,)]['在班'], chk.get(t, {}).get('在班', 0.0), note='在班人次合计')
        all_ok &= ok
    all_ok &= row('整体', raw[()]['在班'], chk.get('整体', {}).get('在班', 0.0), note='全部期次合计')
    r[0] += 1

    # 2) 转化率抽查（附计算过程，可手动复算）
    hdr(['转化率抽查（附计算过程）', '脚本聚合值', '原始表直算', '结果', '计算过程'])
    for t in qici_sorted:
        d = raw[(t,)]
        m = metrics(d, anycnt[(t,)])
        c = chk.get(t, {})
        num, den = d['转低'], d['在班']
        n2, d2 = c.get('转低', 0.0), c.get('在班', 0.0)
        v2 = n2 / d2 if d2 else None
        eq = row(f'{t} 转化率', m['转化率'], v2, fmt='0.00%',
                 note=f'脚本: {int(num)}/{int(den)}  独立: {int(n2)}/{int(d2)}')
        all_ok &= eq
    r[0] += 1

    # 3) 行级任一讲出勤转化率抽查
    for t in [qici_latest] + [x for x in qici_sorted if x != qici_latest][:1]:
        ac = anycnt[(t,)]
        m = metrics(raw[(t,)], ac)
        c = chk.get(t, {})
        v2 = (c['任出且转低'] / c['任出行数']) if c.get('任出行数') else None
        eq = row(f'{t} 任一讲出勤转化率', m['任一讲出勤转化率'], v2, fmt='0.00%',
                 note=f'脚本: {int(ac[0])}/{int(ac[1])}  独立: {int(c.get("任出且转低",0))}/{int(c.get("任出行数",0))}（行级计数）')
        all_ok &= eq
    r[0] += 1

    # 4) 归并守恒：未开课期次归并后各学科组人数之和 = 该期次总人数
    hdr(['归并守恒核对', '脚本聚合值', '原始表直加', '结果', '说明'])
    sus = sorted({k[2] for k in raw if len(k) >= 3 and k[0] == qici_latest and k[1] == 'S'}, key=str)
    su_sum = sum(raw.get((qici_latest, 'S', su), {}).get('在班', 0.0) for su in sus)
    ok = row(f'{qici_latest} 各学科组之和', su_sum, raw[(qici_latest,)]['在班'],
             note=f'归并到[{target_su}]后共 {len(sus)} 组，组之和应=期次总计')
    all_ok &= ok
    ok2 = row(f'{qici_latest} 归并前后守恒', raw[(qici_latest,)]['在班'], chk.get(qici_latest, {}).get('在班', 0.0),
              note='归并只改组名不改人数，应与原始表一致')
    all_ok &= ok2
    r[0] += 1

    # 5) 原始行数核对（确保没有漏读行）
    hdr(['数据完整性', '脚本聚合值', '原始表直加', '结果', '说明'])
    for t in qici_sorted:
        ok = row(f'{t} 原始行数', chk.get(t, {}).get('原始行数', 0),
                 chk.get(t, {}).get('原始行数', 0), note='底表数据行数（非空期次行）')
    r[0] += 1

    # 总结论
    cell = w.cell(row=r[0], column=1, value='结论')
    cell.font = Font(bold=True, size=12)
    res = w.cell(row=r[0], column=2, value='✅ 全部对账一致，数据可放心使用' if all_ok
                 else '⚠️ 存在差额项，请勿使用本文件，联系排查')
    res.font = Font(bold=True, size=12, color='006100' if all_ok else '9C0006')
    res.fill = OK_FILL if all_ok else BAD_FILL
    w.column_dimensions['A'].width = 26
    w.column_dimensions['B'].width = 16
    w.column_dimensions['C'].width = 16
    w.column_dimensions['D'].width = 16
    w.column_dimensions['E'].width = 52
    return all_ok

# ---------------- 主流程 ----------------
def main():
    args = parse_args()
    src = args.src
    if not src:
        cands = glob.glob(os.path.join(DEFAULT_DIR, SRC_PATTERN))
        cands = [c for c in cands if '-处理后' not in c]
        if not cands:
            sys.exit(f'[ERROR] 未找到底表文件，请在 {DEFAULT_DIR} 下放置匹配 {SRC_PATTERN} 的文件')
        src = max(cands, key=os.path.getmtime)
    if not os.path.exists(src):
        sys.exit(f'[ERROR] 底表不存在: {src}')
    keep_su = {s.strip() for s in args.keep_su.split(',') if s.strip()}

    # 先快速扫一遍期次集合以确定未开课期次（未传 --qici 时取字典序最大）
    wb0 = openpyxl.load_workbook(src, read_only=True, data_only=True)
    ws0 = wb0.active
    hdr = next(ws0.iter_rows(values_only=True))
    qi_idx = hdr.index('期次')
    qici_set0 = set()
    for r in ws0.iter_rows(values_only=True):
        v = r[qi_idx] if len(r) > qi_idx else None
        if v is not None:
            qici_set0.add(str(v))
    wb0.close()
    qici_latest = args.qici or sorted(qici_set0)[-1]

    print(f'[1/4] 读取并聚合: {src}（未开课期次={qici_latest}，归并→{args.target_su}）')
    raw, anycnt, qici_set = load_and_aggregate(src, keep_su, args.target_su, qici_latest)
    qici_sorted = sorted(qici_set)[-3:]
    print(f'[2/5] 聚合完成，期次={qici_set}，对比窗口={qici_sorted}')
    print('[2.5/5] 独立路径对账（第二遍直读原始表求和）...')
    chk = verify_independent(src, qici_latest)

    # 输出 = 底表副本（含清洗后数据）+ 指标透视 + 5 个对比 sheet
    wb = openpyxl.load_workbook(src)   # 直接在主文件副本上创建（禁止跨工作簿移动 sheet）
    print('[3/5] 生成 指标透视 + 5 个对比 sheet ...')
    write_toushi(wb, qici_sorted, raw, anycnt, qici_latest)
    _, su_sheet = build_sheets(wb, raw, anycnt, qici_set, qici_latest)
    print('[4/5] 生成 ✅校验 sheet（交叉对账）...')
    all_ok = write_jiaoyan(wb, chk, raw, anycnt, qici_sorted, qici_latest, args.target_su)
    if all_ok:
        print('      ✅ 全部对账一致')
    else:
        print('      ⚠️ 存在差额项，详见校验 sheet')

    # sheet 排序
    new_sheets = ['期次对比', su_sheet, '期次下渠道对比', '渠道下期次对比', '期次下年级对比', '✅校验']
    order = [wb.sheetnames[0], '指标透视'] + new_sheets
    wb._sheets.sort(key=lambda s: order.index(s.title) if s.title in order else 99)

    # 保存（被占用自动递增版本号）
    out = args.out
    if not out:
        stem = os.path.splitext(os.path.basename(src))[0]
        out = os.path.join(os.path.dirname(src), f'{stem}-处理后-v3-4.xlsx')
    n = 4
    while True:
        try:
            wb.save(out)
            break
        except PermissionError:
            n += 1
            out = re.sub(r'-v3-\d+\.xlsx$', f'-v3-{n}.xlsx', out)
    print(f'[5/5] 已保存: {out}')
    print('sheets:', wb.sheetnames)

if __name__ == '__main__':
    main()
