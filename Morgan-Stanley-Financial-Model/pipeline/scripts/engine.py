"""Generic row-spec engine: lays out quarterly + annual sheets and resolves formula templates.

Template tokens
  {key}            same sheet, same period column
  {key@-1}         same sheet, period offset (quarters on _Q sheets, years on _A sheets)
  {Sheet!key}      other model sheet (base name: Drivers, Segments, IS, BS, Capital), same column
  {Sheet!key@-4}   other sheet with offset
  {$name}          global setting cell (absolute)
"""
import re, datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.comments import Comment

FIRST_YEAR, LAST_YEAR = 2018, 2030
NQ = (LAST_YEAR - FIRST_YEAR + 1) * 4          # 52 quarters
NH = 35                                          # historical quarters: 1Q18 .. 3Q26
NY = LAST_YEAR - FIRST_YEAR + 1                  # 13 years
COL0 = 3                                         # first period column (C)
QLAB = [(FIRST_YEAR + i // 4, i % 4 + 1) for i in range(NQ)]
QEND = {1: (3, 31), 2: (6, 30), 3: (9, 30), 4: (12, 31)}

FONT = 'Arial'
BLUE = '0000FF'; BLACK = '000000'; GREEN = '008000'
F_IN = PatternFill('solid', fgColor='FFF2CC')        # light yellow: forecast inputs
F_KEY = PatternFill('solid', fgColor='FFFF00')       # yellow: key global assumptions
F_EST = PatternFill('solid', fgColor='FCE4D6')       # light orange: estimates
F_SEC = PatternFill('solid', fgColor='1F3864')       # navy section header
F_SUB = PatternFill('solid', fgColor='D9E1F2')       # light blue sub-header
F_HDR_A = PatternFill('solid', fgColor='EDEDED')
F_HDR_E = PatternFill('solid', fgColor='DDEBF7')
THIN = Side(style='thin', color='808080')
MED = Side(style='medium', color='1F3864')

FMT = {
    'mm': '#,##0_);(#,##0);"-"_)',
    'mm1': '#,##0.0_);(#,##0.0);"-"_)',
    'bn': '#,##0.0_);(#,##0.0);"-"_)',
    'pct': '0.0%_);(0.0%);"-"_)',
    'pct2': '0.00%_);(0.00%);"-"_)',
    'bps': '0.0_);(0.0);"-"_)',
    'eps': '$0.00_);($0.00);"-"_)',
    'px': '$#,##0.00_);($#,##0.00);"-"_)',
    'sh': '#,##0_);(#,##0);"-"_)',
    'sh1': '#,##0.0_);(#,##0.0);"-"_)',
    'x': '0.000"x"',
    'int': '0',
    'ratio': '0.000_);(0.000);"-"_)',
}

TOK = re.compile(r'\{\$(\w+)\}|\{(?:(\w+)!)?(\w+)(?:@(-?\d+))?\}')


class Row:
    def __init__(self, key=None, label='', unit='', fmt='mm', hist=None, fcst=None, f=None, ann=None,
                 kind='line', bold=False, total=False, indent=0, note=None, hist_note=None,
                 fcst_note=None, cell_notes=None, est_cols=None, guard=False, italic=False):
        self.key, self.label, self.unit, self.fmt = key, label, unit, fmt
        if f is not None:                      # same formula for all periods
            hist = hist if hist is not None else ('f', f)
            fcst = fcst if fcst is not None else ('f', f)
        self.hist, self.fcst, self.ann = hist, fcst, ann
        self.kind, self.bold, self.total, self.indent = kind, bold, total, indent
        self.note, self.hist_note, self.fcst_note = note, hist_note, fcst_note
        self.cell_notes = cell_notes or {}     # {period_index: text}
        self.est_cols = est_cols or set()       # period indices flagged as estimates
        self.guard, self.italic = guard, italic
        self.row = None


def Sec(label):
    return Row(label=label, kind='section')


def Sub(label):
    return Row(label=label, kind='sub')


def Blank():
    return Row(kind='blank')


class Model:
    def __init__(self, wb, H):
        self.wb, self.H = wb, H
        self.sheets = {}        # base -> dict(q=ws, a=ws or None, rows=[Row], rowmap={})
        self.globals = {}       # name -> absolute address
        self.order = []

    # ------------------------------------------------------------------ layout
    def add(self, base, title, rows, annual=True, first_row=8):
        wsq = self.wb.create_sheet(f'{base}_Q')
        wsa = self.wb.create_sheet(f'{base}_A') if annual else None
        rowmap = {}
        r = first_row
        for row in rows:
            row.row = r
            if row.key:
                assert row.key not in rowmap, row.key
                rowmap[row.key] = r
            r += 1
        self.sheets[base] = dict(q=wsq, a=wsa, rows=rows, rowmap=rowmap, title=title)
        self.order.append(base)

    def qref(self, base):
        return f'{base}_Q'

    # ------------------------------------------------------------------ resolve
    def resolve(self, tmpl, cur_base, idx, mode):
        """mode 'q' (idx = quarter index) or 'a' (idx = year index). Returns formula str or None."""
        refs = []
        bad = [False]

        def rep(m):
            g, sheet, key, off = m.group(1), m.group(2), m.group(3), m.group(4)
            if g:
                return self.globals[g]
            base = sheet or cur_base
            sh = self.sheets[base]
            if key not in sh['rowmap']:
                raise KeyError(f'{base}!{key} (in {cur_base}: {tmpl})')
            j = idx + (int(off) if off else 0)
            n = NQ if mode == 'q' else NY
            if j < 0 or j >= n:
                bad[0] = True
                return 'X'
            addr = f'{CL(COL0 + j)}{sh["rowmap"][key]}'
            refs.append(addr if base == cur_base else None)
            if base != cur_base:
                sname = f'{base}_{"Q" if mode == "q" else "A"}'
                if mode == 'a' and self.sheets[base]['a'] is None:
                    sname = f'{base}_Q'
                return f"{sname}!{addr}"
            return addr

        out = TOK.sub(rep, tmpl)
        if bad[0]:
            return None
        return '=' + out

    # ------------------------------------------------------------------ styling helpers
    @staticmethod
    def font(color=BLACK, bold=False, italic=False, size=9):
        return Font(name=FONT, color=color, bold=bold, italic=italic, size=size)

    @staticmethod
    def is_link(formula):
        return bool(re.fullmatch(r"=('?[A-Za-z_ ]+'?)!\$?[A-Z]{1,3}\$?\d+", formula or ''))

    def header(self, ws, title, mode):
        n = NQ if mode == 'q' else NY
        ws['A1'] = 'Morgan Stanley (NYSE: MS)'
        ws['A1'].font = self.font(bold=True, size=12)
        ws['A2'] = title
        ws['A2'].font = self.font(bold=True, size=10)
        ws['A3'] = 'Fiscal year (Dec YE)'
        ws['A4'] = 'Period'
        ws['A5'] = 'Period end'
        ws['A6'] = 'A = actual, E = forecast'
        for a in ('A3', 'A4', 'A5', 'A6'):
            ws[a].font = self.font(italic=True, color='595959')
        ws['B4'] = 'Units'
        ws['B4'].font = self.font(bold=True)
        for j in range(n):
            c = COL0 + j
            if mode == 'q':
                y, q = QLAB[j]
                lab, act = f'{q}Q{str(y)[2:]}', j < NH
                ae = 'A' if act else 'E'
                end = datetime.date(y, *QEND[q])
            else:
                y = FIRST_YEAR + j
                lab = f'FY{y}'
                act = y <= 2025
                ae = 'A' if y <= 2025 else ('A/E' if y == 2026 else 'E')
                end = datetime.date(y, 12, 31)
            fill = F_HDR_A if act else F_HDR_E
            for rr, v in ((3, str(y)), (4, lab), (5, end), (6, ae)):
                cell = ws.cell(row=rr, column=c, value=v)
                cell.font = self.font(bold=(rr == 4))
                cell.alignment = Alignment(horizontal='right')
                cell.fill = fill
                if rr == 5:
                    cell.number_format = 'm/d/yyyy'
            ws.cell(row=6, column=c).border = Border(bottom=MED)
        ws.cell(row=6, column=1).border = Border(bottom=MED)
        ws.cell(row=6, column=2).border = Border(bottom=MED)
        ws.column_dimensions['A'].width = 52
        ws.column_dimensions['B'].width = 9
        for j in range(n):
            ws.column_dimensions[CL(COL0 + j)].width = 9.5 if mode == 'q' else 10.5
        ws.freeze_panes = 'C7'
        ws.sheet_view.zoomScale = 85
        ws.sheet_view.showGridLines = False

    # ------------------------------------------------------------------ write
    def write_label(self, ws, row, n):
        r = row.row
        if row.kind == 'section':
            for c in range(1, COL0 + n):
                ws.cell(row=r, column=c).fill = F_SEC
            ws.cell(row=r, column=1, value=row.label).font = self.font('FFFFFF', bold=True)
            return False
        if row.kind == 'sub':
            for c in range(1, COL0 + n):
                ws.cell(row=r, column=c).fill = F_SUB
            ws.cell(row=r, column=1, value=row.label).font = self.font(bold=True)
            return False
        if row.kind == 'blank':
            return False
        cell = ws.cell(row=r, column=1, value=('   ' * row.indent) + row.label)
        cell.font = self.font(bold=row.bold or row.total, italic=row.italic)
        u = ws.cell(row=r, column=2, value=row.unit)
        u.font = self.font(italic=True, color='595959')
        if row.note:
            cell.comment = Comment(row.note, 'Model')
        return True

    def style_value(self, cell, row, color, fill=None, italic=False):
        cell.font = self.font(color, bold=row.bold or row.total, italic=italic or row.italic)
        cell.number_format = FMT[row.fmt]
        if fill:
            cell.fill = fill
        if row.total:
            cell.border = Border(top=THIN)

    def guard_wrap(self, formula, tmpl_refs_formula):
        refs = re.findall(r"(?:[A-Za-z_]+!)?\$?[A-Z]{1,3}\$?\d+", formula[1:])
        refs = [x for x in refs if not re.fullmatch(r'[A-Z]{1,3}', x)]
        uniq = []
        for x in refs:
            if x not in uniq:
                uniq.append(x)
        if not uniq:
            return formula
        return f'=IF(COUNT({",".join(uniq)})<{len(uniq)},"",IFERROR({formula[1:]},""))'

    def write_quarterly(self, base):
        sh = self.sheets[base]
        ws = sh['q']
        self.header(ws, sh['title'] + ' — Quarterly', 'q')
        for row in sh['rows']:
            if not self.write_label(ws, row, NQ):
                continue
            for i in range(NQ):
                spec = row.hist if i < NH else row.fcst
                cell = ws.cell(row=row.row, column=COL0 + i)
                if spec is None:
                    continue
                kind = spec[0]
                if kind == 'data':
                    name = spec[1]
                    scale = spec[2] if len(spec) > 2 else 1
                    v = self.H[name][i]
                    if v is None:
                        continue
                    cell.value = round(v * scale, 6)
                    est = i in row.est_cols
                    self.style_value(cell, row, BLUE, F_EST if est else None, italic=est)
                elif kind == 'f':
                    fml = self.resolve(spec[1], base, i, 'q')
                    if fml is None:
                        continue
                    if row.guard and i < NH:
                        fml = self.guard_wrap(fml, spec[1])
                    cell.value = fml
                    self.style_value(cell, row, GREEN if self.is_link(fml) else BLACK)
                elif kind == 'in':
                    vals = spec[1]
                    k = i - NH
                    v = vals(QLAB[i][0], QLAB[i][1], k) if callable(vals) else (vals[k] if isinstance(vals, (list, tuple)) else vals)
                    if v is None:
                        continue
                    cell.value = v
                    self.style_value(cell, row, BLUE, F_IN)
                elif kind == 'const':
                    cell.value = spec[1]
                    self.style_value(cell, row, BLUE)
                if i == NH:
                    cell.border = Border(left=MED, top=cell.border.top)
            # notes
            if row.hist_note:
                ws.cell(row=row.row, column=COL0).comment = Comment(row.hist_note, 'Model')
            if row.fcst_note:
                ws.cell(row=row.row, column=COL0 + NH).comment = Comment(row.fcst_note, 'Model')
            for i, txt in row.cell_notes.items():
                ws.cell(row=row.row, column=COL0 + i).comment = Comment(txt, 'Model')

    def write_annual(self, base):
        sh = self.sheets[base]
        ws = sh['a']
        if ws is None:
            return
        wq = f'{base}_Q'
        self.header(ws, sh['title'] + ' — Annual', 'a')
        for row in sh['rows']:
            if not self.write_label(ws, row, NY):
                continue
            if row.ann is None:
                continue
            for j in range(NY):
                q0 = 4 * j
                c0, c3 = CL(COL0 + q0), CL(COL0 + q0 + 3)
                r = row.row
                a = row.ann
                if a == 'sum':
                    fml = f'=SUM({wq}!{c0}{r}:{c3}{r})'
                elif a == 'end':
                    fml = f'={wq}!{c3}{r}'
                elif a == 'beg':
                    fml = f'={wq}!{c0}{r}'
                elif a == 'avg':
                    fml = f'=AVERAGE({wq}!{c0}{r}:{c3}{r})'
                elif isinstance(a, tuple) and a[0] == 'f':
                    fml = self.resolve(a[1], base, j, 'a')
                    if fml is None:
                        continue
                else:
                    raise ValueError(a)
                cell = ws.cell(row=r, column=COL0 + j, value=fml)
                color = GREEN if (a in ('sum', 'end', 'beg', 'avg')) else (GREEN if self.is_link(fml) else BLACK)
                self.style_value(cell, row, color)
                if j == 9:     # FY2027 first full forecast year
                    pass
            if row.note:
                ws.cell(row=row.row, column=1).comment = Comment(row.note, 'Model')
