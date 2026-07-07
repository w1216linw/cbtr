#!/usr/bin/env python3
"""
CBT Report — Daily Pickup Status Matcher

Files:
  address_db.xlsx  — address list (read-only)
  pod.xlsx         — daily POD (read-only, updated externally)
  report.xlsx      — daily report, OVERWRITTEN each run
  history.csv      — cumulative failure history, APPENDED each run
  run.log          — run log (warnings, errors, info)
"""

import csv
import logging
import os
import sys
from datetime import datetime

import openpyxl

from cbt.loader   import load_db, load_pod, load_watch_list
from cbt.matcher  import build_pod_index, determine_status
from cbt.reporter import (append_history, check_consecutive_failures,
                           check_pod_stale, pod_date, write_report)

_log = logging.getLogger('cbt_report')


# ── First-run setup ───────────────────────────────────────────────────────────

def first_run_setup(base_dir: str) -> bool:
    """Create template Excel files if missing. Returns True if any were created."""
    from openpyxl.styles import Alignment, Font, PatternFill

    hdr_fill = PatternFill('solid', fgColor='1F4E79')
    hdr_font = Font(color='FFFFFF', bold=True)
    hdr_aln  = Alignment(horizontal='center')

    def _make_header(ws, label: str, width: int = 60):
        ws.column_dimensions['A'].width = width
        c = ws.cell(row=1, column=1, value=label)
        c.fill, c.font, c.alignment = hdr_fill, hdr_font, hdr_aln

    created = []

    addr_path = os.path.join(base_dir, 'address_db.xlsx')
    if not os.path.exists(addr_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '地址库'
        _make_header(ws, '地址')
        ws.cell(row=2, column=1, value='示例：123 Main St, Chicago, IL, 60601（请删除此行，填入实际地址）')
        wb.save(addr_path)
        created.append(('address_db.xlsx', '地址库，每行一个地址'))

    wl_path = os.path.join(base_dir, 'watch_list.xlsx')
    if not os.path.exists(wl_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '监控地址'
        _make_header(ws, '地址')
        wb.save(wl_path)
        created.append(('watch_list.xlsx', '可选，需人工确认的地址'))

    if created:
        print('=' * 56)
        print('  首次运行 — 已自动生成以下模板文件：')
        print('=' * 56)
        for fname, desc in created:
            print(f'  ✅  {fname}  （{desc}）')
        print()
        print('  还需要手动放入同一目录：')
        print('  ❗  pod.xlsx  — 从系统导出的当日揽收数据')
        print()
        print('  文件准备好后，重新运行程序即可。')
        print('=' * 56)
        return True

    return False


# ── Logger ────────────────────────────────────────────────────────────────────

def setup_logger(base_dir: str) -> None:
    log_path = os.path.join(base_dir, 'run.log')
    _log.setLevel(logging.DEBUG)
    if _log.handlers:
        _log.handlers.clear()

    fmt = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

    fh = logging.FileHandler(log_path, encoding='utf-8', mode='a')
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)
    ch.setFormatter(logging.Formatter('⚠️  [%(levelname)s] %(message)s'))

    _log.addHandler(fh)
    _log.addHandler(ch)

    with open(log_path, 'a', encoding='utf-8') as f:
        f.write(f'\n{"=" * 60}\n RUN {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n{"=" * 60}\n')


# ── Interactive ───────────────────────────────────────────────────────────────

def prompt_watch_list_overrides(results: list, watch_list: set) -> bool:
    """Ask for manual confirmation on watch-list addresses that failed. Mutates results in-place."""
    changed = False
    for r in results:
        if r['status'] != '揽收失败':
            continue
        addr = r['db_addr']
        if addr not in watch_list:
            continue
        print()
        print(f'⚠️  监控地址揽收失败：{addr}')
        while True:
            ans = input('    该地址是否实际取货成功？[y/N] ').strip().lower()
            if ans in ('y', 'yes'):
                r['status'] = '揽收成功'
                r['reason'] = ''
                print('    ✅ 已手动确认为揽收成功。')
                changed = True
                break
            elif ans in ('n', 'no', ''):
                break
            else:
                print('    请输入 y 或 n。')
    return changed


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    # sys.executable points to the .exe when frozen by PyInstaller;
    # falls back to the script location when running as plain Python.
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    setup_logger(base_dir)

    db_path      = os.path.join(base_dir, 'address_db.xlsx')
    pod_path     = os.path.join(base_dir, 'pod.xlsx')
    report_path  = os.path.join(base_dir, 'report.xlsx')
    history_path = os.path.join(base_dir, 'history.csv')

    # ── 首次运行：生成模板文件 ───────────────────────────────────────────────
    if first_run_setup(base_dir):
        input('按 Enter 键退出…')
        sys.exit(0)

    # ── 必要文件检查 ─────────────────────────────────────────────────────────
    missing = []
    for path, label in [(db_path, 'address_db.xlsx'), (pod_path, 'pod.xlsx')]:
        if not os.path.exists(path):
            _log.error(f'缺少必要文件：{label}')
            missing.append(label)
    if missing:
        sys.exit(f'❌  缺少文件：{", ".join(missing)}，请检查后重新运行')

    # ── 加载数据 ─────────────────────────────────────────────────────────────
    db_addrs            = load_db(db_path)
    pod_rows, pod_dates = load_pod(pod_path)
    watch_list          = load_watch_list(os.path.join(base_dir, 'watch_list.xlsx'))
    date_str            = pod_date(pod_rows)

    _log.info(f'address_db 加载完成：{len(db_addrs)} 条地址')
    _log.info(f'pod.xlsx 加载完成：{len(pod_rows)} 条记录，日期 {date_str}')

    # ── POD 数据校验 ─────────────────────────────────────────────────────────
    if not db_addrs:
        _log.error('address_db.xlsx 中没有地址数据，请检查文件内容')
        sys.exit('❌  address_db.xlsx 为空')

    if not pod_rows:
        _log.error('pod.xlsx 中没有有效数据行，请确认文件已正确导出')
        sys.exit('❌  pod.xlsx 无有效数据')

    if len(pod_dates) > 1:
        _log.warning(
            f'pod.xlsx 包含多个日期：{", ".join(pod_dates)}；'
            f'已自动使用最新日期 {pod_dates[-1]}，旧日期数据已忽略'
        )

    ratio = len(pod_rows) / len(db_addrs)
    if ratio >= 2:
        _log.warning(
            f'POD 行数（{len(pod_rows)}）是地址库数量（{len(db_addrs)}）的 {ratio:.1f} 倍，'
            f'请确认导出时已按本站点过滤，否则匹配结果可能不准确'
        )

    # ── 一次性迁移：history.xlsx → history.csv ───────────────────────────────
    old_history = os.path.join(base_dir, 'history.xlsx')
    if os.path.exists(old_history) and not os.path.exists(history_path):
        wb = openpyxl.load_workbook(old_history, read_only=True)
        ws = wb.active
        with open(history_path, 'w', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            for row in ws.iter_rows(values_only=True):
                if any(cell is not None for cell in row):
                    w.writerow(['' if cell is None else str(cell) for cell in row])
        wb.close()
        _log.info('已将 history.xlsx 迁移至 history.csv')
        print('✅  已将 history.xlsx 迁移至 history.csv')

    check_pod_stale(date_str, history_path)

    # ── 匹配 & 输出 ──────────────────────────────────────────────────────────
    full_idx, base_idx = build_pod_index(pod_rows)
    results = [determine_status(addr, full_idx, base_idx) for addr in db_addrs]

    if watch_list:
        overridden = prompt_watch_list_overrides(results, watch_list)
        if overridden:
            print()

    append_history(results, date_str, history_path)
    consecutive, last_dates = check_consecutive_failures(history_path)
    write_report(results, report_path, consecutive)

    total    = len(results)
    ok_cnt   = sum(1 for r in results if r['status'] == '揽收成功')
    fail_cnt = total - ok_cnt

    _log.info(f'运行完成 — 总地址: {total}  揽收成功: {ok_cnt}  揽收失败: {fail_cnt}')
    if consecutive:
        _log.warning(
            f'连续 {len(last_dates)} 天揽收失败（{" / ".join(last_dates)}）：'
            f'{"; ".join(consecutive)}'
        )

    print(f'POD日期 : {date_str}')
    print(f'总地址  : {total}  |  揽收成功: {ok_cnt}  |  揽收失败: {fail_cnt}')
    print(f'报告    : {report_path}')
    print(f'历史    : {history_path}')
    print()
    print('── 揽收失败明细 ──')
    for r in results:
        if r['status'] == '揽收失败':
            print(f'  [{r["reason"]}]  {r["db_addr"]}')

    if consecutive:
        print()
        print(f'⚠️  连续 {len(last_dates)} 天揽收失败（{" / ".join(last_dates)}）：')
        for addr in consecutive:
            print(f'  {addr}')


if __name__ == '__main__':
    main()
