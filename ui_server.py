"""Shipment Studio: local upload, review, correction and download UI.

Run: python ui_server.py --open
Uses Python's local HTTP server; no cloud upload or extra web dependencies.
"""
from __future__ import annotations

import argparse
import base64
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timezone
from email import policy
from email.parser import BytesParser
import gzip
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import io
import json
import math
import os
from pathlib import Path
import re
import shutil
import threading
from types import SimpleNamespace
from urllib.parse import urlparse, parse_qs, quote
from urllib.request import urlopen
import uuid
import webbrowser
from zipfile import ZipFile, BadZipFile, ZIP_DEFLATED

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter, coordinate_to_tuple

import process

ROOT = Path(__file__).resolve().parent
STORE = Path('/tmp/.ui_jobs') if (os.environ.get('VERCEL') or not os.access(ROOT, os.W_OK)) else ROOT / '.ui_jobs'
STORE.mkdir(parents=True, exist_ok=True)
POOL = ThreadPoolExecutor(max_workers=1)
LOCKS = {}
STATE_LOCK = threading.RLock()
MAX_UPLOAD = 45 * 1024 * 1024


def timestamp():
    return datetime.now(timezone.utc).isoformat()


def serial(value):
    if isinstance(value, (datetime, date)):
        return value.isoformat()[:10]
    return value


def write_json(path, data):
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps(data, ensure_ascii=False, default=serial, allow_nan=False), encoding='utf-8')
    temp.replace(path)


def directory(job_id):
    if not re.fullmatch(r'[a-f0-9]{32}', job_id):
        raise ValueError('Mã phiên không hợp lệ.')
    folder = STORE / job_id
    if not (folder / 'job.json').exists():
        raise ValueError('Không tìm thấy phiên xử lý.')
    return folder


def job_lock(job_id):
    with STATE_LOCK:
        return LOCKS.setdefault(job_id, threading.RLock())


def read_job(job_id):
    return json.loads((directory(job_id) / 'job.json').read_text(encoding='utf-8'))


def update_job(job_id, **kwargs):
    with job_lock(job_id):
        state = read_job(job_id)
        state.update(kwargs)
        write_json(directory(job_id) / 'job.json', state)
    return state


def local_input():
    return next((p for p in ROOT.glob('*.xlsx') if not p.name.startswith('~$')
                 and 'INTERNAL SCHEDULE' in p.name.upper()), None)


def save_upload(upload, path):
    if not isinstance(upload, dict) or not str(upload.get('name', '')).lower().endswith('.xlsx'):
        raise ValueError('Vui lòng chọn file .xlsx.')
    try:
        raw = upload['raw'] if isinstance(upload.get('raw'), bytes) else base64.b64decode(upload['data'], validate=True)
    except (ValueError, KeyError) as exc:
        raise ValueError('Không đọc được file tải lên.') from exc
    if len(raw) > MAX_UPLOAD:
        raise ValueError('Mỗi file tối đa 45 MB.')
    try:
        with ZipFile(io.BytesIO(raw)) as archive:
            if '[Content_Types].xml' not in archive.namelist():
                raise ValueError('File không phải workbook Excel hợp lệ.')
            if sum(i.file_size for i in archive.infolist()) > 300 * 1024 * 1024:
                raise ValueError('Workbook quá lớn sau giải nén (tối đa 300 MB).')
    except BadZipFile as exc:
        raise ValueError('File Excel bị hỏng hoặc chưa đúng định dạng .xlsx.') from exc
    path.write_bytes(raw)
    return Path(upload['name'].replace('\\', '/')).name


def parse_payload(content_type, raw):
    """Accept native browser FormData without expanding Excel files into base64."""
    if not content_type.lower().startswith('multipart/form-data'):
        return json.loads(raw)
    if '\r' in content_type or '\n' in content_type:
        raise ValueError('Content-Type không hợp lệ.')
    message = BytesParser(policy=policy.default).parsebytes(
        ('Content-Type: ' + content_type + '\r\nMIME-Version: 1.0\r\n\r\n').encode('ascii') + raw)
    if not message.is_multipart() or message.defects:
        raise ValueError('Dữ liệu tải lên không hợp lệ.')
    payload = {}
    for part in message.iter_parts():
        name = part.get_param('name', header='content-disposition')
        if name not in {'options', 'input', 'template'} or name in payload or part.defects:
            raise ValueError('Trường tải lên không hợp lệ hoặc bị trùng.')
        data = part.get_payload(decode=True)
        if name == 'options':
            payload[name] = json.loads(data)
        else:
            payload[name] = {'name': part.get_filename(), 'raw': data}
    options = payload.pop('options', {})
    if not isinstance(options, dict) or set(options) - {'week', 'season', 'start', 'end', 'use_local'}:
        raise ValueError('Tùy chọn tải lên không hợp lệ.')
    return {**options, **payload}


def create_job(payload):
    week = int(payload.get('week', 40))
    season = str(payload.get('season', '2728')).strip()
    if not 1 <= week <= 53 or not re.fullmatch(r'[A-Za-z0-9_-]{1,20}', season):
        raise ValueError('Tuần phải từ 1–53. Mùa chỉ gồm chữ, số, dấu _ hoặc -.')
    start = process.week_arg(payload['start']) if payload.get('start') else None
    end = process.week_arg(payload['end']) if payload.get('end') else None
    if start and end and start > end:
        raise ValueError('Tuần bắt đầu phải trước hoặc bằng tuần kết thúc.')
    job_id = uuid.uuid4().hex
    folder = STORE / job_id
    folder.mkdir(parents=True)
    source = folder / 'input.xlsx'
    template = folder / 'template.xlsx'
    if payload.get('use_local'):
        found = local_input()
        if not found:
            raise ValueError('Không tìm thấy file nguồn có sẵn. Hãy tải lên file của bạn.')
        shutil.copy2(found, source)
        source_name = found.name
    else:
        source_name = save_upload(payload.get('input'), source)
    if payload.get('template'):
        template_name = save_upload(payload['template'], template)
    else:
        default = process.find_default_template(ROOT)
        if not default or not default.exists():
            raise ValueError('Hãy tải lên file mẫu BARBIE.')
        shutil.copy2(default, template)
        template_name = default.name
    state = {'id': job_id, 'status': 'queued', 'created': timestamp(), 'source': source_name,
             'template': template_name, 'season': season, 'week': week, 'start': payload.get('start', ''),
             'end': payload.get('end', ''), 'completed': 0, 'total': 0, 'message': 'Đang chờ xử lý…'}
    write_json(folder / 'job.json', state)
    args = SimpleNamespace(input=str(source), template=str(template), output=str(folder/'files'),
                           family=None, season=season, week=week, start_week=start, end_week=end,
                           date=None, strict=False, overwrite=False, source_name=source_name)
    if os.environ.get('VERCEL'):
        # A plain background thread may be suspended after a function responds.
        # Finish within this invocation rather than leave polling stuck forever.
        generate(job_id, args)
        return read_job(job_id)
    POOL.submit(generate, job_id, args)
    return state


def generate(job_id, args):
    try:
        update_job(job_id, status='processing', message='Đang đọc SUM và tạo các file Family…')
        args.on_progress = lambda done, total, name: update_job(
            job_id, completed=done, total=total, message=f'Đã tạo {done}/{total} file · {name}')
        report = process.run(args)
        folder = directory(job_id)
        baseline = folder / 'baseline'
        baseline.mkdir()
        source_values = openpyxl.load_workbook(args.input, data_only=True)
        source_reader = process.SheetReader(source_values['SUM'])
        for index, entry in enumerate(report['families']):
            entry['id'] = index
            entry['revision'] = 0
            entry['edits'] = []
            path = folder/'files'/entry['file']
            # The baseline is an exact byte copy: reuse one parsed workbook for
            # both independent checks rather than parsing the same file 3 times.
            exported = openpyxl.load_workbook(path, data_only=True)
            try:
                entry['source_checks'] = reconcile_source(source_reader, path, entry, workbook=exported)
                shutil.copy2(path, baseline/entry['file'])
                entry.update(analyze(path, baseline/entry['file'], entry,
                                     workbook=exported, original_workbook=exported))
            finally:
                exported.close()
        source_values.close()
        write_json(folder/'review.json', report)
        update_job(job_id, status='ready', total=len(report['families']), completed=len(report['families']),
                   message='Đã tạo xong. Bạn có thể kiểm tra và chỉnh sửa.')
    except Exception as exc:
        update_job(job_id, status='error', message=str(exc))


def reconcile_source(reader, path, entry, workbook=None):
    """Independently compare exported quantities and all rollups to the actual SUM."""
    groups = OrderedDict()
    for col in entry['source_columns']:
        key = tuple(process.clean(reader.value(r,col)) for r in (4,15,2))
        groups.setdefault(key, []).append(col)
    columns = [col for group in groups.values() for col in group]
    week_rows = {}
    for row in range(16, reader.sheet.max_row+1):
        for col in (1,22):
            match = process.WEEK_RE.fullmatch(process.clean(reader.value(row,col)))
            if match:
                week_rows[f'{int(match[1])}/{int(match[2]):02d}'] = row
    wb = workbook if workbook is not None else openpyxl.load_workbook(path, data_only=True)
    b, s, _ = wb.worksheets
    checked = 0

    def check(actual, expected):
        nonlocal checked
        if not equal(actual.value, expected):
            raise ValueError(f'Đối chiếu không khớp nguồn: {entry["name"]}, {actual.parent.title}!{actual.coordinate}.')
        checked += 1

    def total(values):
        return next((v for v in values if isinstance(v,str) and v in process.ERRORS), None) or sum(v for v in values if isinstance(v,(int,float)))

    for out_col, source_col in enumerate(columns,2):
        for out_row, source_row in ((9,11),(10,12)):
            check(b.cell(out_row,out_col), reader.value(source_row,source_col))
    for out_col, group in enumerate(groups.values(),2):
        for out_row, source_row in ((14,11),(15,12)):
            check(s.cell(out_row,out_col), total([reader.value(source_row,col) for col in group]))
    running = 0
    for row in range(14,b.max_row+1):
        source_row = week_rows[b.cell(row,1).value]
        values = [reader.value(source_row,col) for col in columns]
        for out_col, expected in enumerate(values,2):
            check(b.cell(row,out_col), expected)
        for out_col, group in enumerate(groups.values(),2):
            check(s.cell(row+5,out_col), total([reader.value(source_row,col) for col in group]))
        weekly = total(values)
        check(s.cell(row+5,entry['markets']+2), weekly)
        running = total([running,weekly])
        check(s.cell(row+5,entry['markets']+3), running)
    if workbook is None:
        wb.close()
    return checked


def get_report(job_id):
    folder = directory(job_id)
    if read_job(job_id)['status'] != 'ready':
        raise ValueError('Phiên chưa xử lý xong.')
    return json.loads((folder/'review.json').read_text(encoding='utf-8'))


def get_entry(report, file_id):
    if not isinstance(file_id, int) or file_id < 0 or file_id >= len(report['families']):
        raise ValueError('Không tìm thấy file Family.')
    return report['families'][file_id]


def equal(a, b):
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return math.isclose(a, b, rel_tol=1e-12, abs_tol=1e-9)
    return serial(a) == serial(b)


def analyze(path, baseline, entry, workbook=None, original_workbook=None):
    wb = workbook if workbook is not None else openpyxl.load_workbook(path, data_only=True)
    original = original_workbook if original_workbook is not None else openpyxl.load_workbook(baseline, data_only=True)
    problems, changes, compared = [], 0, 0
    for sheet in wb:
        for row in sheet:
            for cell in row:
                if isinstance(cell, MergedCell):
                    continue
                old = original[sheet.title][cell.coordinate].value
                if cell.value is not None or old is not None:
                    compared += 1
                    if not equal(cell.value, old):
                        changes += 1
                if cell.data_type == 'e':
                    problems.append({'sheet': sheet.title, 'cell': cell.coordinate,
                                     'kind': 'error', 'message': str(cell.value)})
    breakdown = wb['Breakdown ']
    for row in range(14, breakdown.max_row + 1):
        value = breakdown.cell(row, 1).value
        try:
            year, week = map(int, str(value).split('/'))
            date.fromisocalendar(year, week, 1)
        except (ValueError, TypeError):
            problems.append({'sheet': 'Breakdown ', 'cell': f'A{row}', 'kind': 'week',
                             'message': f'Tuần {value} không hợp lệ theo ISO.'})
    last = wb.worksheets[1].cell(wb.worksheets[1].max_row, entry['markets'] + 3).value
    if workbook is None:
        wb.close()
    if original_workbook is None:
        original.close()
    return {'problems': problems, 'changed_cells': changes, 'compared_cells': compared,
            'current_total': last if isinstance(last, (int, float)) else None,
            'status_label': 'warning' if problems else ('edited' if changes else 'matched')}


def edit_type(wb, entry, sheet, cell):
    if isinstance(cell, MergedCell) or cell.data_type == 'f':
        return None
    r, c = cell.row, cell.column
    if sheet.title == 'Breakdown ':
        if c == 1 and r >= 14:
            return 'week'
        if 2 <= c <= entry['items'] + 1:
            if r in (9, 10) or r >= 14:
                return 'number'
            if r in (3, 4, 7, 8, 12, 13):
                return 'text'
    elif sheet is wb.worksheets[1]:
        if cell.coordinate in ('E3', 'E4'):
            return 'text'
        if cell.coordinate == 'E6':
            return 'date'
        if r in (17, 18) and 2 <= c <= entry['markets'] + 1:
            return 'text'
        if r >= 19 and c >= entry['markets'] + 6:
            return 'date_optional' if c in (entry['markets']+7, entry['markets']+8) else 'text'
    return None


def color_to_hex(color_obj):
    if not color_obj:
        return None
    try:
        ctype = getattr(color_obj, 'type', None)
        if ctype == 'rgb' and color_obj.rgb:
            s = str(color_obj.rgb)
            if len(s) == 8:
                if s.startswith('00'):
                    return None
                return '#' + s[2:]
            elif len(s) == 6:
                return '#' + s
        elif ctype == 'theme' and getattr(color_obj, 'theme', None) is not None:
            theme_colors = [
                'FFFFFF',  # 0: lt1 (Light 1)
                '000000',  # 1: dk1 (Dark 1)
                'E7E6E6',  # 2: lt2 (Light 2)
                '44546A',  # 3: dk2 (Dark 2)
                '5B9BD5',  # 4: accent1 (Soft Blue - TOTAL header)
                'ED7D31',  # 5: accent2 (Orange)
                'A5A5A5',  # 6: accent3 (Gray)
                'FFC000',  # 7: accent4 (Gold)
                '4472C4',  # 8: accent5 (Royal Blue)
                '70AD47'   # 9: accent6 (Fresh Green - KJ header)
            ]
            theme = int(color_obj.theme)
            tint = float(getattr(color_obj, 'tint', 0.0) or 0.0)
            if 0 <= theme < len(theme_colors):
                hex_code = theme_colors[theme]
                r, g, b = int(hex_code[0:2], 16), int(hex_code[2:4], 16), int(hex_code[4:6], 16)
                if tint > 0:
                    r = int(r + (255 - r) * tint)
                    g = int(g + (255 - g) * tint)
                    b = int(b + (255 - b) * tint)
                elif tint < 0:
                    r = int(r * (1 + tint))
                    g = int(g * (1 + tint))
                    b = int(b * (1 + tint))
                return f'#{r:02X}{g:02X}{b:02X}'
        elif ctype == 'indexed' and getattr(color_obj, 'indexed', None) is not None:
            idx = int(color_obj.indexed)
            if 0 <= idx < len(openpyxl.styles.colors.COLOR_INDEX):
                s = str(openpyxl.styles.colors.COLOR_INDEX[idx])
                if len(s) == 8:
                    if s.startswith('00'):
                        return None
                    return '#' + s[2:]
                elif len(s) == 6:
                    return '#' + s
    except Exception:
        pass
    return None


def preview(job_id, file_id, sheet_index):
    with job_lock(job_id):
        report = get_report(job_id)
        entry = get_entry(report, file_id)
        folder = directory(job_id)
        wb = openpyxl.load_workbook(folder/'files'/entry['file'])
        values = openpyxl.load_workbook(folder/'files'/entry['file'], data_only=True)
        old = openpyxl.load_workbook(folder/'baseline'/entry['file'], data_only=True)
        if sheet_index not in (0, 1, 2):
            raise ValueError('Sheet không hợp lệ.')
        sheet = wb.worksheets[sheet_index]
        cells = []
        for row in sheet:
            record = []
            for cell in row:
                value = values[sheet.title][cell.coordinate].value
                original = old[sheet.title][cell.coordinate].value

                bg_color = None
                if cell.fill and getattr(cell.fill, 'fill_type', None) in ('solid', 'patternFill'):
                    c = getattr(cell.fill, 'start_color', None) or getattr(cell.fill, 'fgColor', None)
                    bg_color = color_to_hex(c)

                fg_color = None
                if cell.font and getattr(cell.font, 'color', None):
                    fg_color = color_to_hex(cell.font.color)

                borders = {}
                if cell.border:
                    for side in ('top', 'bottom', 'left', 'right'):
                        b_side = getattr(cell.border, side, None)
                        if b_side and getattr(b_side, 'style', None):
                            borders[side] = {
                                'style': b_side.style,
                                'color': color_to_hex(getattr(b_side, 'color', None)) or '#d4d4d8'
                            }

                record.append({'address': cell.coordinate, 'value': serial(value), 'original': serial(original),
                               'format': cell.number_format,
                               'formula': cell.value if cell.data_type == 'f' else None,
                               'editable': edit_type(wb, entry, sheet, cell),
                               'changed': not equal(value, original),
                               'error': values[sheet.title][cell.coordinate].data_type == 'e',
                               'comment': (cell.comment.text if cell.comment else
                                           entry.get('review_notes', {}).get(sheet.title, {}).get(cell.coordinate, '')),
                               'bold': bool(cell.font.bold),
                               'italic': bool(getattr(cell.font, 'italic', False)),
                               'font_size': getattr(cell.font, 'size', None),
                               'bg': bg_color,
                               'fg': fg_color,
                               'align': getattr(cell.alignment, 'horizontal', None) if cell.alignment else None,
                               'valign': getattr(cell.alignment, 'vertical', None) if cell.alignment else None,
                               'wrap': bool(getattr(cell.alignment, 'wrap_text', False)) if cell.alignment else False,
                               'borders': borders if borders else None})
            cells.append(record)

        col_widths = {}
        for c_idx in range(1, sheet.max_column + 1):
            col_letter = get_column_letter(c_idx)
            dim = sheet.column_dimensions.get(col_letter)
            if dim and dim.width:
                col_widths[col_letter] = round(dim.width * 7.5, 1)

        row_heights = {}
        for r_idx in range(1, sheet.max_row + 1):
            dim = sheet.row_dimensions.get(r_idx)
            if dim and dim.height:
                row_heights[str(r_idx)] = round(dim.height * 1.33, 1)

        result = {'file': entry, 'revision': entry['revision'], 'sheets': wb.sheetnames,
                  'sheet': sheet.title, 'sheet_index': sheet_index, 'rows': cells,
                  'merges': [str(a) for a in sheet.merged_cells.ranges],
                  'columns': [get_column_letter(c) for c in range(1, sheet.max_column+1)],
                  'col_widths': col_widths, 'row_heights': row_heights}
        for book in (wb, values, old):
            book.close()
        return result


def recalculate(wb):
    """Evaluate only the internal SUM formulas authored by process.py, never eval()."""
    cache, visiting = {}, set()

    def value(sheet, address):
        key = (sheet, address)
        if key in cache:
            return cache[key]
        cell = wb[sheet][address]
        if cell.data_type != 'f':
            return cell.value
        if key in visiting:
            raise ValueError('Phát hiện công thức vòng.')
        visiting.add(key)
        match = re.fullmatch(r'=SUM\((.*)\)', cell.value)
        if not match:
            raise ValueError(f'Công thức chưa hỗ trợ: {sheet}!{address}')
        # A quoted sheet name can contain commas: tokenize references, not split(',').
        refs = re.findall(r"(?:'((?:[^']|'')+)'!)?([A-Z]+[1-9][0-9]*)(?:,|$)", match[1])
        reconstructed = ','.join((f"'{s}'!" if s else '') + a for s, a in refs)
        if reconstructed != match[1]:
            raise ValueError('Tham chiếu công thức không hợp lệ.')
        vals = [value(s.replace("''", "'") if s else sheet, a) for s, a in refs]
        error = next((v for v in vals if isinstance(v, str) and v in process.ERRORS), None)
        result = error or sum(v for v in vals if isinstance(v, (int, float)))
        if isinstance(result, (int, float)) and not math.isfinite(result):
            raise ValueError('Tổng số lượng vượt giới hạn.')
        cache[key] = result
        visiting.remove(key)
        return result

    for sheet in wb:
        for row in sheet:
            for cell in row:
                if cell.data_type == 'f':
                    value(sheet.title, cell.coordinate)
    wb._formula_cache = cache


def sync_summary(wb, entry):
    breakdown, summary, release = wb.worksheets
    # Market header merges define fixed groups; item identity/grouping is read-only.
    groups = []
    col = 2
    while col <= entry['items'] + 1:
        area = next((a for a in breakdown.merged_cells.ranges
                     if a.min_row == 3 and a.min_col == col), None)
        end = area.max_col if area else col
        groups.append((col, end))
        col = end + 1
    if len(groups) != entry['markets']:
        raise ValueError('Cấu trúc thị trường không khớp.')
    for col, (start, end) in enumerate(groups, 2):
        for row, source_row in ((9, 3), (10, 4)):
            process.put(summary, row, col, breakdown.cell(source_row, start).value)
        for row, source_row in ((11, 7), (12, 13)):
            vals = list(dict.fromkeys(process.clean(breakdown.cell(source_row, c).value)
                                     for c in range(start, end+1)))
            process.put(summary, row, col, ' / '.join(v for v in vals if v))
    seen = set()
    for row in range(14, breakdown.max_row+1):
        label = str(breakdown.cell(row, 1).value)
        if label in seen:
            raise ValueError(f'Tuần {label} đã tồn tại trong file.')
        seen.add(label)
        year, week = map(int, label.split('/'))
        try:
            monday = date.fromisocalendar(year, week, 1)
        except ValueError:
            monday = None
        process.put(summary, row+5, 1, label)
        process.put(summary, row+5, entry['markets']+5, week)
        for sheet, r, c in ((summary, row+5, entry['markets']+4), (release, row-11, 1)):
            process.put(sheet, r, c, monday)
            entry.setdefault('review_notes', {}).setdefault(sheet.title, {})[
                f'{get_column_letter(c)}{r}'] = f'Tuần {label}.'


def apply_edit(job_id, file_id, payload):
    with job_lock(job_id):
        report = get_report(job_id)
        entry = get_entry(report, file_id)
        if payload.get('revision') != entry['revision']:
            raise ValueError('File vừa được cập nhật. Hãy tải lại bảng trước khi sửa tiếp.')
        folder = directory(job_id)
        path, baseline = folder/'files'/entry['file'], folder/'baseline'/entry['file']
        wb = openpyxl.load_workbook(path)
        sheet_id = payload.get('sheet')
        address = payload.get('cell', '')
        if sheet_id not in (0, 1, 2) or not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]{0,6}', address):
            raise ValueError('Địa chỉ ô không hợp lệ.')
        sheet = wb.worksheets[sheet_id]
        r, c = coordinate_to_tuple(address)
        if r > sheet.max_row or c > sheet.max_column:
            raise ValueError('Ô nằm ngoài bảng dữ liệu.')
        cell = sheet[address]
        kind = edit_type(wb, entry, sheet, cell)
        if not kind:
            raise ValueError('Ô công thức hoặc cấu trúc chỉ đọc. Hãy sửa dữ liệu ở Breakdown.')
        old_value = serial(cell.value)
        if payload.get('restore'):
            original = openpyxl.load_workbook(baseline, data_only=True)
            new_value = original[sheet.title][address].value
            original.close()
        else:
            new_value = payload.get('value')
            if kind == 'number':
                if new_value in ('', None):
                    new_value = None
                else:
                    try:
                        new_value = float(new_value)
                    except (TypeError, ValueError) as exc:
                        raise ValueError('Nhập số hợp lệ, dùng dấu chấm cho phần thập phân.') from exc
                    if not math.isfinite(new_value) or abs(new_value) > 1e15:
                        raise ValueError('Số lượng phải hữu hạn và không vượt 1.000.000.000.000.000.')
            elif kind == 'week':
                try:
                    year, week = process.week_arg(str(new_value))
                    date.fromisocalendar(year, week, 1)
                except (ValueError, argparse.ArgumentTypeError) as exc:
                    raise ValueError('Nhập tuần ISO hợp lệ dạng YYYY/WW, ví dụ 2028/01.') from exc
                new_value = f'{year}/{week:02d}'
            elif kind.startswith('date'):
                try:
                    new_value = date.fromisoformat(str(new_value)) if new_value else None
                except ValueError as exc:
                    raise ValueError('Ngày cần có dạng YYYY-MM-DD.') from exc
                if new_value is None and kind == 'date':
                    raise ValueError('Ngày báo cáo không được để trống.')
            else:
                new_value = str(new_value or '').strip()
                if len(new_value) > 1000:
                    raise ValueError('Nội dung tối đa 1.000 ký tự.')
        process.put(sheet, r, c, new_value)
        sync_summary(wb, entry)
        recalculate(wb)
        process.save_with_cache(wb, path)
        wb.close()
        entry['revision'] += 1
        entry['edits'].append({'time': timestamp(), 'sheet': sheet.title, 'cell': address,
                               'before': old_value, 'after': serial(new_value),
                               'reason': str(payload.get('reason', '')).strip()[:500],
                               'restored': bool(payload.get('restore'))})
        entry.update(analyze(path, baseline, entry))
        write_json(folder/'review.json', report)
        return entry


class Handler(BaseHTTPRequestHandler):
    server_version = 'ShipmentStudio/1.0'

    def log_message(self, fmt, *args):
        if args and str(args[1] if len(args) > 1 else '') not in ('200', '304'):
            super().log_message(fmt, *args)

    def send(self, data, status=200, content_type='application/json; charset=utf-8', filename=None):
        if not isinstance(data, bytes):
            data = json.dumps(data, ensure_ascii=False, default=serial, allow_nan=False).encode('utf-8')
        compressed = False
        encodings = [part.strip().lower() for part in self.headers.get('Accept-Encoding', '').split(',')]
        if len(data) > 1024 and 'gzip' in encodings and content_type.startswith(('application/json', 'text/')):
            data = gzip.compress(data, compresslevel=1)
            compressed = True
        self.send_response(status)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Vary', 'Accept-Encoding')
        if compressed:
            self.send_header('Content-Encoding', 'gzip')
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        self.send_header('Content-Security-Policy', "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'")
        if filename:
            self.send_header('Content-Disposition', "attachment; filename*=UTF-8''"+quote(filename))
        self.end_headers()
        self.wfile.write(data)

    def guard(self):
        host = self.headers.get('Host', '').split(':')[0].lower()
        if host in {'127.0.0.1', 'localhost'}:
            pass
        elif host.endswith('.vercel.app') or os.environ.get('VERCEL'):
            pass
        else:
            allowed_env = os.environ.get('ALLOWED_HOSTS', '')
            allowed = [h.strip().lower() for h in allowed_env.split(',') if h.strip()] if allowed_env else []
            if not any(host == h or host.endswith('.' + h) for h in allowed):
                raise ValueError(f'Ứng dụng chỉ nhận kết nối localhost hoặc domain được cấp phép.')

        origin = self.headers.get('Origin')
        if origin:
            parsed_origin = urlparse(origin)
            origin_host = (parsed_origin.hostname or '').lower()
            if origin_host not in {'127.0.0.1', 'localhost'} and not origin_host.endswith('.vercel.app') and not os.environ.get('VERCEL'):
                allowed_env = os.environ.get('ALLOWED_HOSTS', '')
                allowed = [h.strip().lower() for h in allowed_env.split(',') if h.strip()] if allowed_env else []
                if not any(origin_host == h or origin_host.endswith('.' + h) for h in allowed):
                    raise ValueError('Nguồn yêu cầu không được phép.')

    def do_GET(self):
        try:
            self.guard()
            parsed = urlparse(self.path)
            path = parsed.path
            query = parse_qs(parsed.query)
            static = {'/': ('index.html', 'text/html; charset=utf-8'),
                      '/app.js': ('app.js', 'text/javascript; charset=utf-8'),
                      '/style.css': ('style.css', 'text/css; charset=utf-8')}
            if path in static:
                name, mime = static[path]
                return self.send((ROOT/'ui'/name).read_bytes(), content_type=mime)
            if path == '/api/config':
                jobs = []
                for p in sorted(STORE.glob('*/job.json'), key=lambda p: p.stat().st_mtime, reverse=True)[:12]:
                    jobs.append(json.loads(p.read_text(encoding='utf-8')))
                default_tpl = process.find_default_template(ROOT)
                tpl_name = default_tpl.name if default_tpl and default_tpl.exists() else 'BARBIE 2728 _Weekly shipment schedule 2728_WK39.xlsx'
                return self.send({'local_input': local_input().name if local_input() else None,
                                  'template': tpl_name,
                                  'jobs': jobs})
            match = re.fullmatch(r'/api/jobs/([a-f0-9]{32})(?:/(report|preview|download|zip))?', path)
            if not match:
                return self.send({'error': 'Không tìm thấy đường dẫn.'}, 404)
            job_id, action = match.groups()
            with job_lock(job_id):
                if not action:
                    return self.send(read_job(job_id))
                report = get_report(job_id)
                if action == 'report':
                    return self.send(report)
                file_id = int(query.get('file', ['0'])[0])
                entry = get_entry(report, file_id)
                if action == 'preview':
                    return self.send(preview(job_id, file_id, int(query.get('sheet', ['0'])[0])))
                folder = directory(job_id)
                if action == 'download':
                    return self.send((folder/'files'/entry['file']).read_bytes(),
                                     content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=entry['file'])
                if action == 'zip':
                    buffer = io.BytesIO()
                    with ZipFile(buffer, 'w', ZIP_DEFLATED) as archive:
                        for entry in report['families']:
                            archive.write(folder/'files'/entry['file'], arcname=entry['file'])
                        archive.writestr('review_report.json', json.dumps(report, ensure_ascii=False, default=serial, indent=2))
                    return self.send(buffer.getvalue(), content_type='application/zip',
                                     filename=f'Shipment_{read_job(job_id)["season"]}_WK{read_job(job_id)["week"]:02d}.zip')
        except (ValueError, KeyError, IndexError, OSError) as exc:
            self.send({'error': str(exc)}, 400)

    def do_POST(self):
        try:
            self.guard()
            length = int(self.headers.get('Content-Length', '0'))
            if not 0 < length <= 125 * 1024 * 1024:
                raise ValueError('Dữ liệu tải lên quá lớn hoặc rỗng.')
            payload = parse_payload(self.headers.get('Content-Type', ''), self.rfile.read(length))
            if not isinstance(payload, dict):
                raise ValueError('Dữ liệu yêu cầu không hợp lệ.')
            if self.path == '/api/jobs':
                return self.send(create_job(payload), 202)
            match = re.fullmatch(r'/api/jobs/([a-f0-9]{32})/edit/([0-9]+)', self.path)
            if match:
                return self.send(apply_edit(match[1], int(match[2]), payload))
            self.send({'error': 'Không tìm thấy đường dẫn.'}, 404)
        except (ValueError, TypeError, KeyError, IndexError, OSError, argparse.ArgumentTypeError) as exc:
            self.send({'error': str(exc)}, 400)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=8765)
    parser.add_argument('--open', action='store_true', help='Mở trình duyệt mặc định')
    args = parser.parse_args()
    STORE.mkdir(exist_ok=True)
    url = f'http://127.0.0.1:{args.port}'
    try:
        server = ThreadingHTTPServer(('127.0.0.1', args.port), Handler)
    except OSError:
        if args.open:
            try:
                with urlopen(url+'/api/config', timeout=2) as response:
                    if 'ShipmentStudio/' in response.headers.get('Server', ''):
                        webbrowser.open(url)
                        return
            except OSError:
                pass
        raise SystemExit(f'Port {args.port} is in use. Open {url} or use --port 8766.')
    # Only the newly bound server can mark interrupted work; a second launcher must not.
    for p in STORE.glob('*/job.json'):
        state = json.loads(p.read_text(encoding='utf-8'))
        if state['status'] in ('queued', 'processing'):
            state.update(status='error', message='Phiên trước bị gián đoạn. Vui lòng tải lên và xử lý lại.')
            write_json(p, state)
    print(f'Shipment Studio: {url}', flush=True)
    if args.open:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        POOL.shutdown(wait=True)


if __name__ == '__main__':
    main()
