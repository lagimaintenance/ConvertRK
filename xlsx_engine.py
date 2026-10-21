"""Baca COA/contoh XLSX, serta ekspor Buku Bank Excel tanpa menyimpan file di server.

Menggunakan Python standard library untuk membuat OOXML (.xlsx), tanpa macro.
"""
from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal
from io import BytesIO
from pathlib import PurePosixPath
import re
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG = 'http://schemas.openxmlformats.org/package/2006/relationships'


def excel_date(d: date) -> int:
    return (d - date(1899, 12, 30)).days


def as_decimal(value) -> Decimal:
    if value in (None, ''):
        return Decimal(0)
    if isinstance(value, Decimal):
        return value
    return Decimal(str(value).replace(',', ''))


def money_key(value) -> str:
    return str(as_decimal(value).quantize(Decimal('0.01')))


def normalize_desc(text: str) -> str:
    return re.sub(r'\s+', ' ', (text or '').casefold()).strip()


def normalize_code(value) -> str:
    if value is None or value == '':
        return ''
    x = str(value).strip()
    return x[:-2] if x.endswith('.0') else x


def read_xlsx_sheet(payload: bytes, sheet_name: str) -> list[dict[str, object]]:
    """Reader OOXML minimal: inline/shared string, numeric/date-as-serial, formula cache."""
    try:
        zf = zipfile.ZipFile(BytesIO(payload))
    except zipfile.BadZipFile as exc:
        raise ValueError('File Excel .xlsx tidak valid.') from exc
    with zf:
        files = set(zf.namelist())
        if 'xl/workbook.xml' not in files:
            raise ValueError('File Excel tidak memiliki workbook.xml.')
        wb = ET.fromstring(zf.read('xl/workbook.xml'))
        book_rels = ET.fromstring(zf.read('xl/_rels/workbook.xml.rels'))
        rels = {rel.attrib['Id']: rel.attrib.get('Target', '') for rel in book_rels}
        ws_names = []
        selected = None
        for sh in wb.findall('.//{%s}sheet' % S):
            ws_names.append(sh.attrib['name'])
            if sh.attrib['name'].casefold() == sheet_name.casefold():
                selected = sh
        if selected is None:
            raise ValueError(f'Sheet {sheet_name!r} tidak ditemukan. Tersedia: {", ".join(ws_names)}')
        relation = selected.attrib['{%s}id' % R]
        target = rels.get(relation, '')
        if target.startswith('/'): target = target.lstrip('/')
        else: target = str(PurePosixPath('xl') / target)
        # Target dapat memakai '../' pada beberapa penulis XLSX.
        import posixpath
        target = posixpath.normpath(target)
        if target not in files:
            raise ValueError(f'Sheet {sheet_name} tidak dapat dibaca dari arsip Excel.')
        strings = []
        if 'xl/sharedStrings.xml' in files:
            string_tree = ET.fromstring(zf.read('xl/sharedStrings.xml'))
            for si in string_tree.findall('{%s}si' % S):
                strings.append(''.join((n.text or '') for n in si.iter('{%s}t' % S)))
        tree = ET.fromstring(zf.read(target))
        rows = []
        for row in tree.findall('.//{%s}sheetData/{%s}row' % (S, S)):
            result = {}
            for c in row.findall('{%s}c' % S):
                ref = c.attrib['r']
                col = re.match(r'[A-Z]+', ref).group(0)
                typ = c.attrib.get('t')
                v = c.find('{%s}v' % S)
                if typ == 'inlineStr':
                    value = ''.join(n.text or '' for n in c.iter('{%s}t' % S))
                elif typ == 's' and v is not None:
                    value = strings[int(v.text)]
                elif typ in ('str', 'e'):
                    value = v.text if v is not None else None
                elif v is not None and v.text is not None:
                    num = Decimal(v.text)
                    value = int(num) if num == int(num) else num
                else:
                    value = None
                if value is not None:
                    result[col] = value
            if result:
                result['_row'] = int(row.attrib.get('r', len(rows) + 1))
                rows.append(result)
        return rows


def read_coa(payload: bytes) -> dict[str, dict]:
    rows = read_xlsx_sheet(payload, 'Kode Akun')
    codes = {}
    for row in rows:
        code = normalize_code(row.get('A'))
        title = str(row.get('B') or '').strip()
        if code and code.lower() != 'kode akun' and title:
            codes[code] = {'name': title, 'dc': str(row.get('C') or ''), 'type': str(row.get('D') or '')}
    if not codes:
        raise ValueError('COA kosong: cek sheet Kode Akun (kolom A:D).')
    return codes


def read_example(payload: bytes) -> dict:
    """Pemetaan transaksi dari template: cocok tepat menurut tanggal+uraian+nominal."""
    rows = read_xlsx_sheet(payload, 'Bank')
    examples = {}
    patterns = defaultdict(set)
    for row in rows:
        value = row.get('A')
        if not isinstance(value, int) or value < 25000 or value > 90000:
            continue
        try:
            tanggal = (date(1899, 12, 30) + timedelta(days=value)).isoformat()
        except (OverflowError, ValueError):
            continue
        desc = str(row.get('F') or '')
        code = normalize_code(row.get('D'))
        if not desc or not code:
            continue
        debit = money_key(row.get('G'))
        credit = money_key(row.get('H'))
        exact_key = (tanggal, normalize_desc(desc), debit, credit)
        examples[exact_key] = code
        # hanya infer pattern secara konservatif
        patterns[normalize_desc(desc)].add(code)
    return {'exact': examples, 'stable_patterns': {k: next(iter(v)) for k, v in patterns.items() if len(v) == 1}}


def suggest_account(row: dict, coa: dict, example: dict | None = None,
                    custom_rules: list[tuple[str, str]] | None = None) -> tuple[str, bool, str]:
    description = normalize_desc(row['uraian'])
    key = (row['tanggal'], description, money_key(row['debet']), money_key(row['kredit']))
    if example and key in example.get('exact', {}):
        code = example['exact'][key]
        if code in coa:
            return code, True, 'Sesuai contoh (transaksi identik)'
    for keyword, code in custom_rules or []:
        if keyword and keyword.casefold().strip() in description and code in coa:
            return code, True, f'Aturan pengguna: {keyword}'
    if 'biaya adm' in description and '901' in coa:
        return '901', True, 'Biaya administrasi bank'
    if any(x in description for x in ('bca finance', 'angsuran bcaf')) and '203' in coa:
        return '203', True, 'Angsuran BCA Finance'
    # Rekening koran saja tidak cukup untuk menetapkan lawan akun transaksi lain.
    if 'setoran tunai' in description and '121' in coa:
        return '121', False, 'Usulan mengikuti Januari; perlu bukti sumber setoran'
    if description.startswith('tarikan') and '122' in coa:
        return '122', False, 'Usulan mengikuti Januari; perlu verifikasi tujuan'
    return '', False, 'Belum cukup informasi untuk penentuan akun'


# Xlsx OOXML creation using standard library only.
styles = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="2"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/><numFmt numFmtId="165" formatCode="#,##0.00;[Red](#,##0.00);-"/></numFmts>
<fonts count="4">
<font><sz val="10"/><name val="Aptos"/></font>
<font><b/><sz val="16"/><color rgb="FF17365A"/><name val="Aptos Display"/></font>
<font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
<font><b/><sz val="10"/><color rgb="FF17365A"/><name val="Aptos"/></font>
</fonts>
<fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF17365A"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFEAF0F6"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFEDC1"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFDCFCE7"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="hair"><color rgb="FFCBD5E1"/></left><right style="hair"><color rgb="FFCBD5E1"/></right><top style="hair"><color rgb="FFCBD5E1"/></top><bottom style="hair"><color rgb="FFCBD5E1"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center"/></xf>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0"/>
<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="center"/></xf>
<xf numFmtId="165" fontId="3" fillId="3" borderId="1" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0"/>
</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'''


def _excel_col(n: int) -> str:
    s = ''
    while n:
        n, d = divmod(n - 1, 26)
        s = chr(65 + d) + s
    return s


def _cell(row_num: int, col_num: int, value, style: int | None = None, formula: str | None = None,
          cached=None) -> str:
    if value is None and formula is None:
        return ''
    ref = f'{_excel_col(col_num)}{row_num}'
    styling = f' s="{style}"' if style is not None else ''
    if formula:
        val = cached if cached is not None else value
        return f'<c r="{ref}"{styling}><f>{escape(formula)}</f><v>{val}</v></c>'
    if isinstance(value, (date, datetime)):
        serial = excel_date(value.date() if isinstance(value, datetime) else value)
        return f'<c r="{ref}" s="{style if style is not None else 5}"><v>{serial}</v></c>'
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        return f'<c r="{ref}"{styling}><v>{value}</v></c>'
    return f'<c r="{ref}"{styling} t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'


class SheetBuilder:
    def __init__(self, name, widths, merges=None, freeze=5):
        self.name = name
        self.widths = widths
        self.merges = merges or []
        self.freeze = freeze
        self.rows = []
        self.autofilter = None

    def add(self, cells, styles=None, height=None, formulas=None):
        self.rows.append((cells, styles or {}, height, formulas or {}))

    def xml(self) -> str:
        w = ''.join(f'<col min="{i}" max="{i}" width="{width}" customWidth="1"/>'
                    for i, width in enumerate(self.widths, 1))
        parts = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                 f'<worksheet xmlns="{S}"><sheetViews><sheetView workbookViewId="0">'
                 + (f'<pane ySplit="{self.freeze}" topLeftCell="A{self.freeze+1}" activePane="bottomLeft" state="frozen"/>' if self.freeze else '')
                 + '</sheetView></sheetViews>',
                 '<sheetFormatPr defaultRowHeight="17"/>', f'<cols>{w}</cols>', '<sheetData>']
        for idx, (cells, sty, height, formulas) in enumerate(self.rows, 1):
            attrs = f' r="{idx}"' + (f' ht="{height}" customHeight="1"' if height else '')
            content = ''.join(_cell(idx, ci, val, sty.get(ci), formulas.get(ci, (None,None))[0] if ci in formulas else None,
                                    formulas.get(ci, (None,None))[1] if ci in formulas else None)
                              for ci, val in enumerate(cells, 1))
            parts.append(f'<row{attrs}>{content}</row>')
        parts.append('</sheetData>')
        if self.merges:
            parts.append(f'<mergeCells count="{len(self.merges)}">'+''.join(f'<mergeCell ref="{m}"/>' for m in self.merges)+'</mergeCells>')
        if self.autofilter:
            parts.append(f'<autoFilter ref="{self.autofilter}"/>')
        parts.append('<pageMargins left="0.2" right="0.2" top="0.5" bottom="0.5" header="0.2" footer="0.2"/>')
        parts.append('</worksheet>')
        return ''.join(parts)


def _ledger(name: str, tx: list[dict], start_amount: Decimal, company: str, year: int,
            coa: dict, multi_start=None) -> SheetBuilder:
    sheet = SheetBuilder(name, [16, 17, 14, 20, 31, 73, 21, 21, 21, 29],
                         merges=['A1:J1', 'A2:J2'])
    sheet.add([company], {1: 1}, height=27)
    sheet.add(['Buku Bank'], {1: 10}, height=21)
    sheet.add([year], {1: 10})
    sheet.add([], height=8)
    sheet.add(['Tanggal', 'No. Bukti', 'Kode Akun', 'Kode Lawan Akun', 'Nama Akun',
               'Uraian', 'Debet', 'Kredit', 'Saldo', 'Nama Vendor'], {i: 2 for i in range(1, 11)}, height=30)
    sheet.add([], height=7)
    sheet.add([None,None,None,None,None,'SALDO BANK AWAL',None,None,start_amount], {6: 3,9:9},height=22)
    running = start_amount
    for txrow in tx:
        if txrow.get('_new_opening') is not None:
            running = as_decimal(txrow['_new_opening'])
            sheet.add([None,None,None,None,None,'SALDO BANK AWAL - ' + txrow['tanggal'][:7],None,None,running],
                      {6:3,9:9}, height=22)
        running += as_decimal(txrow['debet']) - as_decimal(txrow['kredit'])
        rnum = len(sheet.rows) + 1
        code = normalize_code(txrow.get('kode_lawan'))
        selected = coa.get(code, {})
        debit, credit = as_decimal(txrow['debet']), as_decimal(txrow['kredit'])
        row = [date.fromisoformat(txrow['tanggal']), txrow.get('no_bukti') or None,
               int(txrow.get('kode_bank', 111)), int(code) if code.isdigit() else code or None,
               selected.get('name') or None, txrow['uraian'], debit if debit else None,
               credit if credit else None, running, txrow.get('vendor') or None]
        height = min(110, max(22, 13 * (txrow['uraian'].count('\n') + 1) + 6))
        cell_styles = {1: 5, 6: 8, 7: 4, 8: 4, 9: 4}
        if not txrow.get('verified'):
            cell_styles[4] = 6
        sheet.add(row, cell_styles, height, formulas={9:(f'I{rnum-1}+G{rnum}-H{rnum}', running)})
    sheet.autofilter = f'A5:J{len(sheet.rows)}'
    return sheet


def generate_xlsx(statements: list[dict], tx_rows: list[dict], coa: dict, company: str,
                  bank_code: int = 111) -> bytes:
    if not statements:
        raise ValueError('Tidak ada rekening koran untuk diekspor.')
    statements = sorted(statements, key=lambda s: (s['rekening'],s['year'],s['month']))
    stmt_index = {(s['file'], s['rekening']): s for s in statements}
    buckets = defaultdict(list)
    for tx in tx_rows:
        s = stmt_index[(tx['file'], tx['rekening'])]
        bucket = (s['rekening'],s['year'],s['month'],s['file'])
        tx = dict(tx)
        tx['kode_bank'] = bank_code
        buckets[bucket].append(tx)
    sheets = []
    grouped_accounts = defaultdict(list)
    for s in statements:
        grouped_accounts[s['rekening']].append(s)
    for account_pos,(account,sts) in enumerate(grouped_accounts.items(),1):
        sts.sort(key=lambda s:(s['year'], s['month']))
        combined = []
        previous_balance = None
        for stmt in sts:
            bucket = (stmt['rekening'],stmt['year'],stmt['month'],stmt['file'])
            for ii, tx in enumerate(buckets[bucket]):
                if ii == 0 and previous_balance is not None and previous_balance != stmt['saldo_awal']:
                    tx['_new_opening'] = stmt['saldo_awal']
                combined.append(tx)
            previous_balance = stmt['saldo_akhir']
        label = 'Bank' if account_pos == 1 else f'Bank {account_pos}'
        year_label = sts[0]['year'] if sts[0]['year'] == sts[-1]['year'] else f"{sts[0]['year']}-{sts[-1]['year']}"
        sheets.append(_ledger(label,combined,sts[0]['saldo_awal'],company,year_label,coa))
        for stmt in sts:
            month_name = f"{stmt['year']}-{stmt['month']:02d}"
            if len(grouped_accounts) > 1:
                month_name += f' R{account_pos}'
            bucket = (stmt['rekening'],stmt['year'],stmt['month'],stmt['file'])
            sheets.append(_ledger(month_name,buckets[bucket],stmt['saldo_awal'],company,stmt['year'],coa))
    audits = SheetBuilder('Rekonsiliasi', [25, 44, 20, 21, 21, 21, 21, 21, 22, 22, 22, 24])
    audits.add(['REKONSILIASI REKENING KORAN'],{1:1},height=31)
    audits.add(['Selisih 0 berarti jumlah mutasi dan saldo sesuai dengan PDF.'],height=22)
    audits.add([],height=8)
    audits.add(['Bulan','File','No Rekening','Saldo Awal','Mutasi CR PDF','Debet Olahan',
                'Mutasi DB PDF','Kredit Olahan','Saldo Akhir PDF','Saldo Akhir Hitung','Selisih','Status'],
               {i:2 for i in range(1,13)},height=34)
    for s in statements:
        mismatch = bool(s.get('errors'))
        net = (as_decimal(s['saldo_akhir_pdf']) - as_decimal(s['saldo_akhir'])) if s['saldo_akhir_pdf'] is not None else None
        status = 'PERLU CEK' if mismatch or net is None else 'COCOK'
        values = [f"{s['month']:02d}/{s['year']}",s['file'],s['rekening'],s['saldo_awal'],s['mutasi_cr_pdf'],s['debet'],
                  s['mutasi_db_pdf'],s['kredit'],s['saldo_akhir_pdf'],s['saldo_akhir'],net,status]
        colors = {12: 6 if status == 'PERLU CEK' else 7}
        colors.update({k:4 for k in range(4,12)})
        audits.add(values,colors,26)
    audits.autofilter=f'A4:L{len(audits.rows)}'
    sheets.append(audits)
    reviews = SheetBuilder('Perlu Review',[20,20,17,49,18,18,19,34,49,20],freeze=3)
    reviews.add(['TRANSAKSI BELUM DIVERIFIKASI'],{1:1},height=30)
    reviews.add(['Gunakan halaman aplikasi untuk memperbaiki dan menyetujui klasifikasi akun.'],height=25)
    reviews.add(['File','Tanggal','No Rekening','Uraian','Debet','Kredit','Kode Lawan','Alasan/Klasifikasi','No. Bukti','Nama Vendor'],
                {i:2 for i in range(1,11)},height=33)
    for t in tx_rows:
        if t.get('verified'): continue
        reviews.add([t['file'],date.fromisoformat(t['tanggal']),t['rekening'],t['uraian'],
                     as_decimal(t['debet']) or None,as_decimal(t['kredit']) or None,
                     t.get('kode_lawan'),t.get('catatan','Perlu tinjau'),t.get('no_bukti'),t.get('vendor')],
                    {2:5,4:8,5:4,6:4,7:6},height=48)
    reviews.autofilter=f'A3:J{len(reviews.rows)}'
    sheets.append(reviews)
    coas = SheetBuilder('Kode Akun',[18,47,15,15],freeze=2)
    coas.add(['KODE AKUN'],{1:1},height=27)
    coas.add(['Kode Akun','Nama Akun','D/C','BS/PL'],{i:2 for i in range(1,5)},height=28)
    for code,info in sorted(coa.items(), key=lambda v:(not v[0].isdigit(),int(v[0]) if v[0].isdigit() else v[0])):
        coas.add([int(code) if code.isdigit() else code,info['name'],info.get('dc'),info.get('type')])
    coas.autofilter=f'A2:D{len(coas.rows)}'
    sheets.append(coas)
    # OOXML archive; definitions explicitly closed, no external relationships.
    out=BytesIO()
    with zipfile.ZipFile(out,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=6) as package:
        def write(name,val): package.writestr(name,val.encode('utf-8') if isinstance(val,str) else val)
        overrides = ['<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
        for i,sh in enumerate(sheets,1):
            overrides.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            write(f'xl/worksheets/sheet{i}.xml',sh.xml())
        write('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
              '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
              '<Default Extension="xml" ContentType="application/xml"/>'
              +''.join(overrides)+'</Types>')
        write('_rels/.rels',f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{R}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        names=''.join(f'<sheet name="{escape(sh.name)}" sheetId="{i}" r:id="rId{i}"/>' for i,sh in enumerate(sheets,1))
        write('xl/workbook.xml',f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              f'<workbook xmlns="{S}" xmlns:r="{R}"><sheets>{names}</sheets>'
              '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')
        relationships = ''.join(f'<Relationship Id="rId{i}" Type="{R}/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(sheets)+1))
        relationships += f'<Relationship Id="rId{len(sheets)+1}" Type="{R}/styles" Target="styles.xml"/>'
        write('xl/_rels/workbook.xml.rels',f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              f'<Relationships xmlns="{PKG}">{relationships}</Relationships>')
        write('xl/styles.xml',styles)
    return out.getvalue()
