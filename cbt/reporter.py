import csv
import logging
import os
import re
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

_log = logging.getLogger('cbt_report')

_HDR_FILL  = PatternFill('solid', fgColor='1F4E79')
_HDR_FONT  = Font(color='FFFFFF', bold=True)
_OK_FILL   = PatternFill('solid', fgColor='C6EFCE')
_FL_FILL   = PatternFill('solid', fgColor='FFC7CE')
_WARN_FILL = PatternFill('solid', fgColor='FFEB9C')
_WARN_FONT = Font(color='9C6500', bold=True)
_CTR       = Alignment(horizontal='center')


# ── Helpers ───────────────────────────────────────────────────────────────────

def _save(wb, path: str):
    try:
        wb.save(path)
    except PermissionError:
        msg = f'无法保存 {os.path.basename(path)}：文件被其他程序占用，请关闭后重新运行'
        _log.error(msg)
        sys.exit(f'❌  {msg}')


def pod_date(pod_rows) -> str:
    for r in pod_rows:
        m = re.match(r'(\d{4}-\d{2}-\d{2})', str(r.get('time', '')))
        if m:
            return m.group(1)
    return datetime.today().strftime('%Y-%m-%d')


# ── Daily report ──────────────────────────────────────────────────────────────

def write_report(results, path: str, consecutive: list):
    total    = len(results)
    ok_cnt   = sum(1 for r in results if r['status'] == '揽收成功')
    fail_cnt = total - ok_cnt

    failures  = [r for r in results if r['status'] == '揽收失败']
    fail_cell = '\n'.join(f"{r['db_addr']} - {r['reason']}" for r in failures)

    def street_number(addr):
        m = re.match(r'(\d+)', addr.strip())
        return m.group(1) if m else ''

    street_nums = ', '.join(street_number(r['db_addr']) for r in failures if street_number(r['db_addr']))
    consec_cell = '\n'.join(consecutive)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '揽收报告'

    headers = ['地址库总数', '揽收成功', '揽收失败', '', '揽收失败明细', '揽收失败街道号', '⚠️ 连续三天未揽收']
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        if not h:
            continue
        if col == 7 and consecutive:
            c.fill = _WARN_FILL
            c.font = _WARN_FONT
        else:
            c.fill = _HDR_FILL
            c.font = _HDR_FONT
        c.alignment = _CTR

    ws.cell(row=2, column=1, value=total)
    ws.cell(row=2, column=2, value=ok_cnt).fill  = _OK_FILL
    ws.cell(row=2, column=3, value=fail_cnt).fill = _FL_FILL
    ws.cell(row=2, column=4, value='')

    e = ws.cell(row=2, column=5, value=fail_cell)
    e.alignment = Alignment(wrap_text=True, vertical='top')
    if failures:
        e.fill = _FL_FILL

    f = ws.cell(row=2, column=6, value=street_nums)
    f.alignment = Alignment(wrap_text=True, vertical='top')
    if failures:
        f.fill = _FL_FILL

    g = ws.cell(row=2, column=7, value=consec_cell)
    g.alignment = Alignment(wrap_text=True, vertical='top')
    if consecutive:
        g.fill = _WARN_FILL

    for col, w in zip('ABCDEFG', [14, 14, 14, 6, 80, 60, 60]):
        ws.column_dimensions[col].width = w
    ws.row_dimensions[2].height = max(30, max(len(failures), len(consecutive)) * 15)

    _save(wb, path)


# ── History CSV ───────────────────────────────────────────────────────────────

def append_history(results, date_str: str, path: str):
    """Write today's failures to history.csv, replacing any existing rows for the same date."""
    failures = [r for r in results if r['status'] == '揽收失败']

    existing = []
    if os.path.exists(path):
        with open(path, newline='', encoding='utf-8-sig') as f:
            for row in csv.reader(f):
                if row and row[0] != date_str:
                    existing.append(row)
    else:
        existing = [['日期', '地址', '揽收失败原因']]

    with open(path, 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f)
        w.writerows(existing)
        for r in failures:
            w.writerow([date_str, r['db_addr'], r['reason']])


def check_consecutive_failures(path: str, n: int = 3):
    """Returns addresses that failed on every one of the last n recorded dates."""
    if not os.path.exists(path):
        return [], []

    with open(path, newline='', encoding='utf-8-sig') as f:
        rows = [(r[0], r[1]) for r in csv.reader(f) if len(r) >= 2 and r[0] != '日期' and r[0] and r[1]]

    if not rows:
        return [], []

    all_dates = sorted(set(d for d, _ in rows))
    if len(all_dates) < n:
        return [], all_dates

    last_n = set(all_dates[-n:])
    addr_dates = {}
    for date, addr in rows:
        addr_dates.setdefault(addr, set()).add(date)

    flagged = [addr for addr, dates in addr_dates.items() if last_n.issubset(dates)]
    return flagged, sorted(last_n)


def check_pod_stale(date_str: str, history_path: str):
    """Warn if pod.xlsx date already exists in history and today is a different date."""
    if not os.path.exists(history_path):
        return
    today = datetime.today().strftime('%Y-%m-%d')
    if date_str == today:
        return
    with open(history_path, newline='', encoding='utf-8-sig') as f:
        existing = {r[0] for r in csv.reader(f) if r and r[0] != '日期'}
    if date_str in existing:
        _log.warning(
            f'pod.xlsx 的日期 {date_str} 已存在于历史记录中，'
            f'pod.xlsx 可能尚未更新，report.xlsx 将使用旧数据'
        )
