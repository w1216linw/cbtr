import logging

from .normalize import norm_full, norm_base, parse_db_addr

_log = logging.getLogger('cbt_report')

# POD 揽收失败原因关键词 → 统一输出文案（顺序优先，先匹配先返回）
_REASON_MAP = [
    ('无包裹可揽收', '系统有订单但无实物'),
    ('无包裹可揽',   '系统有订单但无实物'),
    ('没有包裹可揽', '系统有订单但无实物'),
    ('仓库关门',     '仓库关门'),
    ('节假日关门',   '仓库关门'),
    ('关门',         '仓库关门'),
]


def _map_reason(raw: str) -> str:
    for keyword, label in _REASON_MAP:
        if keyword in raw:
            return label
    if raw:
        _log.debug(f'未匹配的揽收失败原因（原文）："{raw}"')
    return raw


def build_pod_index(pod_rows):
    full_idx, base_idx = {}, {}
    for row in pod_rows:
        raw = str(row['addr']).split(',')[0].strip()
        city = row['city'].lower().strip()
        full_idx.setdefault((norm_full(raw), city), []).append(row)
        base_idx.setdefault((norm_base(raw), city), []).append(row)
    return full_idx, base_idx


def _resolve(matches):
    if not matches:
        return None
    if any(m['status'] == '揽收成功' for m in matches):
        return {'status': '揽收成功', 'reason': ''}
    raw_reasons = list(dict.fromkeys(m['reason'] for m in matches if m['reason']))
    mapped = list(dict.fromkeys(_map_reason(r) for r in raw_reasons))
    return {
        'status': '揽收失败',
        'reason': '；'.join(mapped) if mapped else '揽收失败',
    }


def determine_status(db_addr: str, full_idx: dict, base_idx: dict) -> dict:
    base, full, city = parse_db_addr(db_addr)
    has_unit = (base != full)

    if has_unit:
        matches = full_idx.get((full, city), [])
        if not matches and not city:
            matches = [e for (k, _), v in full_idx.items() if k == full for e in v]
    else:
        matches = base_idx.get((base, city), [])
        if not matches and not city:
            matches = [e for (k, _), v in base_idx.items() if k == base for e in v]

    result = _resolve(matches)
    if result is None:
        result = {'status': '揽收失败', 'reason': '系统无CBT订单'}
    result['db_addr'] = db_addr
    return result
