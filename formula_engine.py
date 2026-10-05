"""Strict calculator for the schedule's arithmetic, SUM and WEEKNUM formulas.

Unsupported syntax and cycles abort publication. No Python eval or stale caches.
"""
from datetime import date, datetime
import math
import posixpath
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile

import openpyxl
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import range_boundaries, get_column_letter
from openpyxl.utils.datetime import to_excel, from_excel

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
TAG = '{' + NS + '}'
ET.register_namespace('', NS)


class UnsupportedFormula(ValueError):
    pass


class ExcelError(Exception):
    pass


class Reference(list):
    """Keep reference coercion distinct from literal function arguments."""


class Calculator:
    def __init__(self, book):
        self.book = book
        self.cache = {}
        self.active = set()
        self.names = {s.title.casefold(): s.title for s in book}

    def cell(self, sheet, address):
        sheet = self.names.get(sheet.casefold())
        if sheet is None:
            raise ExcelError('#REF!')
        key = sheet, address.replace('$', '').upper()
        if key in self.cache:
            value = self.cache[key]
            if isinstance(value, ExcelError):
                raise value
            return value
        if key in self.active:
            raise UnsupportedFormula(f'Tham chiếu vòng tại {sheet}!{address}.')
        cell = self.book[sheet][key[1]]
        value = cell.value
        if cell.data_type == 'e':
            raise ExcelError(value)
        if cell.data_type == 'f':
            self.active.add(key)
            try:
                if not isinstance(value, str):
                    raise UnsupportedFormula('Chưa hỗ trợ công thức mảng.')
                value = Parser(self, sheet, value).parse()
                value = self.scalar(value)
                if value is None:
                    value = 0
                if isinstance(value, (int, float)) and not math.isfinite(value):
                    raise ExcelError('#NUM!')
                self.cache[key] = value
            except ExcelError as exc:
                self.cache[key] = exc
                raise
            except (UnsupportedFormula, RecursionError) as exc:
                raise UnsupportedFormula(f'{sheet}!{address}: {exc}') from exc
            finally:
                self.active.remove(key)
        if isinstance(value, (date, datetime)):
            value = to_excel(value, self.book.epoch)
        return value

    @staticmethod
    def scalar(value):
        if isinstance(value, Reference):
            if len(value) != 1:
                raise UnsupportedFormula('Chưa hỗ trợ phép tính mảng/implicit intersection.')
            return value[0]
        return value

    def number(self, value):
        value = self.scalar(value)
        if value is None:
            return 0
        try:
            result = float(value)
        except (TypeError, ValueError):
            raise ExcelError('#VALUE!')
        if not math.isfinite(result):
            raise ExcelError('#NUM!')
        return result

    def reference(self, sheet, text):
        if '[' in text or ']' in text:
            raise UnsupportedFormula('Chưa hỗ trợ liên kết workbook ngoài hoặc bảng có tên.')
        if '#REF!' in text:
            raise ExcelError('#REF!')
        if '!' in text:
            sheet, text = text.rsplit('!', 1)
            if sheet.startswith("'") and sheet.endswith("'"):
                sheet = sheet[1:-1].replace("''", "'")
        if not re.fullmatch(r'\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?', text):
            raise UnsupportedFormula('Chưa hỗ trợ tham chiếu: ' + text)
        a, b, c, d = range_boundaries(text)
        a, c = min(a, c), max(a, c)
        b, d = min(b, d), max(b, d)
        if c > 16384 or d > 1048576 or (c-a+1)*(d-b+1) > 200000:
            raise UnsupportedFormula('Vùng tham chiếu quá lớn.')
        return Reference(self.cell(sheet, f'{get_column_letter(col)}{row}')
                         for row in range(b, d+1) for col in range(a, c+1))

    def function(self, name, args):
        if name == 'SUM':
            total = 0
            for arg in args:
                if isinstance(arg, Reference):
                    total += sum(v for v in arg if isinstance(v, (int, float)) and not isinstance(v, bool))
                else:
                    total += self.number(arg)
            return total
        if name == 'WEEKNUM':
            if not 1 <= len(args) <= 2:
                raise UnsupportedFormula('WEEKNUM cần 1 hoặc 2 tham số.')
            serial = self.number(args[0])
            mode = int(self.number(args[1])) if len(args) == 2 else 1
            if mode not in (1, 2, 11, 12, 13, 14, 15, 16, 17, 21) or serial < 1:
                raise ExcelError('#NUM!')
            try:
                day = from_excel(int(serial), self.book.epoch)
                if not isinstance(day, datetime):
                    raise ExcelError('#NUM!')
                if day.year == 1900:
                    raise UnsupportedFormula('Chưa hỗ trợ WEEKNUM năm 1900.')
                if mode == 21:
                    return day.isocalendar().week
                start = 6 if mode in (1, 17) else 0 if mode == 2 else mode-11
                first = datetime(day.year, 1, 1)
                return ((day-first).days + (first.weekday()-start) % 7)//7 + 1
            except (OverflowError, ValueError):
                raise ExcelError('#NUM!')
        raise UnsupportedFormula('Chưa hỗ trợ hàm ' + name)

    def calculate(self):
        for sheet in self.book:
            # Sparse iteration avoids expanding a sheet to a distant blank reference.
            for cell in list(sheet._cells.values()):
                if cell.data_type == 'f':
                    try:
                        self.cell(sheet.title, cell.coordinate)
                    except ExcelError:
                        pass
        return {key: str(v) if isinstance(v, ExcelError) else v for key, v in self.cache.items()}


class Parser:
    PRECEDENCE = {',': 50, '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '+': 10, '-': 10, '*': 20, '/': 20, '^': 30}

    def __init__(self, calc, sheet, formula):
        self.calc, self.sheet = calc, sheet
        try:
            self.tokens = [t for t in Tokenizer(formula).items if t.type != 'WHITE-SPACE']
        except Exception as exc:
            raise UnsupportedFormula('Cú pháp công thức không hợp lệ.') from exc
        self.i = 0
        # Validate the whole expression before evaluating errors like #REF!.
        for t in self.tokens:
            if t.type == 'FUNC' and t.subtype == 'OPEN' and t.value[:-1].upper() not in ('SUM', 'WEEKNUM'):
                raise UnsupportedFormula('Chưa hỗ trợ hàm ' + t.value[:-1])
            if t.type == 'OPERATOR-INFIX' and t.value not in self.PRECEDENCE:
                raise UnsupportedFormula('Chưa hỗ trợ toán tử ' + t.value)
            if t.type in ('ARRAY',):
                raise UnsupportedFormula('Chưa hỗ trợ công thức mảng.')

    def pop(self):
        if self.i >= len(self.tokens):
            raise UnsupportedFormula('Công thức thiếu toán hạng.')
        token = self.tokens[self.i]
        self.i += 1
        return token

    def expression(self, minimum=0):
        token = self.pop()
        if token.type == 'OPERATOR-PREFIX' and token.value in ('+', '-'):
            left = self.expression(40)
            if token.value == '-':
                left = -self.calc.number(left)
        elif token.type == 'PAREN' and token.subtype == 'OPEN':
            left = self.expression()
            end = self.pop()
            if end.type != 'PAREN' or end.subtype != 'CLOSE':
                raise UnsupportedFormula('Thiếu dấu đóng ngoặc.')
        elif token.type == 'FUNC' and token.subtype == 'OPEN':
            args = []
            if self.i < len(self.tokens) and self.tokens[self.i].subtype != 'CLOSE':
                while True:
                    args.append(self.expression())
                    if self.i >= len(self.tokens) or self.tokens[self.i].type != 'SEP':
                        break
                    self.pop()
            end = self.pop()
            if end.type != 'FUNC' or end.subtype != 'CLOSE':
                raise UnsupportedFormula('Thiếu dấu đóng hàm.')
            left = self.calc.function(token.value[:-1].upper(), args)
        elif token.type == 'OPERAND':
            if token.subtype == 'NUMBER':
                left = float(token.value)
            elif token.subtype == 'RANGE':
                left = self.calc.reference(self.sheet, token.value)
            elif token.subtype == 'ERROR':
                raise ExcelError(token.value)
            elif token.subtype == 'TEXT':
                left = token.value[1:-1].replace('""', '"')
            elif token.subtype == 'LOGICAL':
                left = token.value.upper() == 'TRUE'
            else:
                raise UnsupportedFormula('Toán hạng chưa hỗ trợ.')
        else:
            raise UnsupportedFormula('Cú pháp chưa hỗ trợ: ' + token.value)
        while self.i < len(self.tokens):
            op = self.tokens[self.i]
            if op.type == 'OPERATOR-POSTFIX' and op.value == '%':
                self.pop(); left = self.calc.number(left)/100
                continue
            power = self.PRECEDENCE.get(op.value, -1) if op.type == 'OPERATOR-INFIX' else -1
            if power < minimum:
                break
            self.pop()
            right = self.expression(power+1)
            if power == 1:
                a, b = self.calc.scalar(left), self.calc.scalar(right)
                if a is None: a = '' if isinstance(b, str) else False if isinstance(b, bool) else 0
                if b is None: b = '' if isinstance(a, str) else False if isinstance(a, bool) else 0
                def rank(v):
                    return (2, v) if isinstance(v, bool) else (1, v.casefold()) if isinstance(v, str) else (0, v)
                a, b = rank(a), rank(b)
                left = {'=': lambda: a==b, '<>': lambda: a!=b, '<': lambda: a<b,
                        '>': lambda: a>b, '<=': lambda: a<=b, '>=': lambda: a>=b}[op.value]()
                continue
            if op.value == ',':
                if not isinstance(left, Reference) or not isinstance(right, Reference):
                    raise UnsupportedFormula('Phép hợp chỉ áp dụng cho tham chiếu.')
                left = Reference(left + right)
                continue
            a, b = self.calc.number(left), self.calc.number(right)
            try:
                left = {'+': lambda: a+b, '-': lambda: a-b, '*': lambda: a*b,
                        '/': lambda: a/b, '^': lambda: a**b}[op.value]()
            except ZeroDivisionError:
                raise ExcelError('#DIV/0!')
            except (ValueError, OverflowError):
                raise ExcelError('#NUM!')
            if isinstance(left, complex) or not math.isfinite(left):
                raise ExcelError('#NUM!')
        return left

    def parse(self):
        result = self.expression()
        if self.i != len(self.tokens):
            raise UnsupportedFormula('Phần công thức chưa hỗ trợ: ' + self.tokens[self.i].value)
        return result


def edit_and_calculate(path, edits):
    book = openpyxl.load_workbook(path)
    try:
        changed = {}
        formula_changed = False
        for edit in edits:
            cell = book[edit['sheet']][edit['cell']]
            formula_changed = formula_changed or cell.data_type == 'f' or edit['kind'] == 'formula'
            value = edit['value']
            if edit['kind'] == 'date':
                value = datetime.fromisoformat(value)
            cell.value = value
            if edit['kind'] == 'text':
                cell.data_type = 's'
            changed[(cell.parent.title, cell.coordinate)] = (value, cell.data_type)
        calculator = Calculator(book)
        caches = calculator.calculate()
        errors = {key for key, value in calculator.cache.items() if isinstance(value, ExcelError)}
        patch_workbook(path, book, changed, caches, errors, formula_changed)
    finally:
        book.close()


def patch_workbook(path, book, changed, caches, errors, formula_changed=False):
    """Patch only cell payloads; retain styles, dimensions, drawings and other ZIP parts."""
    stage = path.with_suffix('.calculating.xlsx')
    try:
        with ZipFile(path) as source, ZipFile(stage, 'w') as target:
            rels = {r.attrib['Id']: posixpath.normpath(posixpath.join('xl', r.attrib['Target']))
                    if not r.attrib['Target'].startswith('/') else r.attrib['Target'].lstrip('/')
                    for r in ET.fromstring(source.read('xl/_rels/workbook.xml.rels'))}
            sheets = ET.fromstring(source.read('xl/workbook.xml')).find(TAG+'sheets')
            names = {rels[s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']]: s.attrib['name'] for s in sheets}
            for entry in source.infolist():
                if formula_changed and entry.filename == 'xl/calcChain.xml':
                    continue
                data = source.read(entry.filename)
                if formula_changed and entry.filename in ('[Content_Types].xml', 'xl/_rels/workbook.xml.rels'):
                    text = data.decode('utf-8')
                    text = re.sub(r'<(?:Override|Relationship)\b[^>]*(?:calcChain)[^>]*/>', '', text)
                    data = text.encode('utf-8')
                if entry.filename in names:
                    title = names[entry.filename]
                    # Preserve the XML namespace prefixes and all non-cell markup verbatim.
                    text = data.decode('utf-8')
                    if not re.search(r'<sheetData\b', text):
                        raise UnsupportedFormula('Cấu trúc XML sheet chưa hỗ trợ: '+title)
                    present = set()
                    def update(match):
                        raw = match.group(0)
                        node = ET.fromstring(raw.replace('<c ', f'<c xmlns="{NS}" ', 1))
                        address = node.attrib['r']; key = title, address
                        present.add(key)
                        if key not in changed and key not in caches:
                            return raw
                        if key in changed:
                            value, kind = changed[key]
                            for child in list(node):
                                if child.tag in (TAG+'v', TAG+'f', TAG+'is'):
                                    node.remove(child)
                            if kind == 'f':
                                ET.SubElement(node, TAG+'f').text = value[1:]
                            else:
                                set_value(node, value, kind, book.epoch)
                        if key in caches:
                            if formula_changed:
                                # Expand shared formulas so editing the shared master
                                # cannot change formulas in other cells on reopening.
                                for child in list(node):
                                    if child.tag == TAG+'f': node.remove(child)
                                ET.SubElement(node, TAG+'f').text = book[title][address].value[1:]
                            value = caches[key]
                            kind = 'e' if key in errors else 'str' if isinstance(value, str) else 'n'
                            set_value(node, value, kind, book.epoch)
                        # Use local default namespace only for this cell.
                        return ET.tostring(node, encoding='unicode')
                    text = re.sub(r'<c\b[^>]*?(?:/>|>.*?</c>)', update, text, flags=re.S)
                    if any(key[0] == title and key not in present and key not in changed for key in caches):
                        raise UnsupportedFormula('Cấu trúc XML sheet chưa hỗ trợ: '+title)
                    missing = [key for key in changed if key[0] == title and key not in present]
                    for key in missing:
                        value, kind = changed[key]
                        node = ET.Element(TAG+'c', {'r': key[1]})
                        if kind == 'f':
                            ET.SubElement(node, TAG+'f').text = value[1:]
                            value = caches[key]
                            kind = 'e' if key in errors else 'str' if isinstance(value, str) else 'n'
                        set_value(node, value, kind, book.epoch)
                        raw = ET.tostring(node, encoding='unicode')
                        row = str(book[title][key[1]].row)
                        pattern = r'(<row\b[^>]*\br="'+row+r'"[^>]*>)(.*?)(</row>)'
                        def insert(m):
                            body = m[2]
                            col = book[title][key[1]].column
                            following = next((c for c in re.finditer(r'<c\b[^>]*\br="([A-Z]+[0-9]+)"', body)
                                              if book[title][c[1]].column > col), None)
                            offset = following.start() if following else len(body)
                            return m[1]+body[:offset]+raw+body[offset:]+m[3]
                        text, count = re.subn(pattern, insert, text, count=1, flags=re.S)
                        if not count:
                            empty = r'(<row\b[^>]*\br="'+row+r'"[^>]*?)/>'
                            text, count = re.subn(empty, lambda m: m[1]+'>'+raw+'</row>', text, count=1)
                        if not count:
                            new_row = '<row r="'+row+'">'+raw+'</row>'
                            following = next((r for r in re.finditer(r'<row\b[^>]*\br="([0-9]+)"', text) if int(r[1])>int(row)), None)
                            if following:
                                text=text[:following.start()]+new_row+text[following.start():]
                            elif '</sheetData>' in text:
                                text=text.replace('</sheetData>',new_row+'</sheetData>')
                            else:
                                text, count = re.subn(r'<sheetData\b([^>]*?)/>', lambda m: '<sheetData'+m[1]+'>'+new_row+'</sheetData>', text, count=1)
                                if not count:
                                    raise UnsupportedFormula('Không thể ghi ô vào sheet: '+title)
                    data = text.encode('utf-8')
                target.writestr(entry, data)
        stage.replace(path)
    finally:
        stage.unlink(missing_ok=True)


def set_value(node, value, kind, epoch):
    for child in list(node):
        if child.tag in (TAG+'v', TAG+'is'):
            node.remove(child)
    node.attrib.pop('t', None)
    if value is None:
        return
    if isinstance(value, (date, datetime)):
        value = to_excel(value, epoch); kind = 'n'
    if kind == 's':
        node.set('t', 'inlineStr')
        ET.SubElement(ET.SubElement(node, TAG+'is'), TAG+'t', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = str(value)
    else:
        node.set('t', 'b' if isinstance(value, bool) else kind)
        ET.SubElement(node, TAG+'v').text = str(int(value)) if isinstance(value, bool) else str(value)
