"""ConvertRK sederhana: PDF rekening koran BCA -> satu sheet GL.

Jalankan: python -m streamlit run app.py
Ketergantungan: streamlit, pdfplumber. Tidak memakai fitz/PyMuPDF.
"""
from __future__ import annotations

from datetime import date
from copy import copy
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
import xml.etree.ElementTree as ET
import hashlib
import re
import zipfile

import pdfplumber
import streamlit as st
from openpyxl import load_workbook


# ====== 1. MEMBACA REKENING KORAN ======
BULAN = {
    'JANUARI': 1, 'FEBRUARI': 2, 'MARET': 3, 'APRIL': 4, 'MEI': 5,
    'JUNI': 6, 'JULI': 7, 'AGUSTUS': 8, 'SEPTEMBER': 9, 'OKTOBER': 10,
    'NOVEMBER': 11, 'DESEMBER': 12,
}
ANGKA = re.compile(r'^-?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}$')
TANGGAL = re.compile(r'^([0-3]\d)/(0[1-9]|1[0-2])$')
MAX_PDF = 100
MAX_BYTES = 80 * 1024 * 1024


class GagalBaca(ValueError):
    pass


def ke_nominal(s: str) -> Decimal:
    teks = str(s).strip()
    if not ANGKA.fullmatch(teks):
        raise GagalBaca(f'Nominal tidak dikenali: {s!r}')
    try:
        return Decimal(teks.replace(',', ''))
    except InvalidOperation as err:
        raise GagalBaca(f'Nominal tidak valid: {s!r}') from err


def baris_pdf(page):
    """Bentuk baris berbasis koordinat PDF, bukan hasil split teks biasa."""
    words = page.extract_words(x_tolerance=2, y_tolerance=2)
    scaling = 595.0 / float(page.width)  # referensi posisi e-statement BCA A4
    rows = []
    for word in sorted(words, key=lambda w: (w['top'], w['x0'])):
        y = float(word['top'])
        entry = {'x': float(word['x0']) * scaling, 'text': word['text']}
        if rows and abs(rows[-1]['y'] - y) < 2.0:
            rows[-1]['words'].append(entry)
        else:
            rows.append({'y': y, 'words': [entry]})
    for row in rows:
        row['words'].sort(key=lambda w: w['x'])
        row['text'] = ' '.join(w['text'] for w in row['words'])
    return rows


def nominal_kolom(row, low, high):
    data = [w['text'] for w in row['words'] if low <= w['x'] < high and ANGKA.fullmatch(w['text'])]
    if len(data) > 1:
        raise GagalBaca('Ada lebih dari satu angka pada kolom mutasi/saldo: ' + str(data))
    return ke_nominal(data[0]) if data else None


def uraian_transaksi(rows):
    descriptions = []
    for row in rows:
        left = [w['text'] for w in row['words'] if 83 <= w['x'] < 190]
        right = [w['text'] for w in row['words'] if 190 <= w['x'] < 302]
        if left:
            descriptions.append(' '.join(left))
        if right:
            descriptions.append(' '.join(right))
    return '\n'.join(descriptions).strip()


def metadata(full_text: str, filename: str):
    match = re.search(
        r'\b(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\s+(20\d{2})\b',
        full_text, re.IGNORECASE,
    )
    if match:
        month, year = BULAN[match.group(1).upper()], int(match.group(2))
    else:
        match = re.search(r'(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)', filename)
        if not match:
            raise GagalBaca('Periode bulan/tahun tidak ditemukan pada PDF atau nama file YYYYMM.')
        year, month = int(match.group(1)), int(match.group(2))
    match_rek = re.search(r'NO\.?\s*REKENING\s*:\s*([\d\s]{8,23})', full_text, re.IGNORECASE)
    account = re.sub(r'\D', '', match_rek.group(1)) if match_rek else ''
    return year, month, account or 'TIDAK TERBACA'


def baca_halaman(page, year: int, month: int):
    rows = baris_pdf(page)
    start_y = 175.0
    end_y = float('inf')
    pdf_totals = {}
    for row in rows:
        upper = row['text'].upper()
        if 'TANGGAL' in upper and 'KETERANGAN' in upper and 'MUTASI' in upper:
            start_y = row['y'] + 3
        if re.search(r'\bMUTASI\s+CR\b', upper):
            amount = nominal_kolom(row, 285, 410)
            if amount is not None:
                pdf_totals['cr'] = amount
            counts = [int(w['text']) for w in row['words'] if 405 <= w['x'] < 460 and w['text'].isdigit()]
            if len(counts) == 1:
                pdf_totals['cr_count'] = counts[0]
            end_y = min(end_y, row['y'] - 1)
        elif re.search(r'\bMUTASI\s+DB\b', upper):
            amount = nominal_kolom(row, 285, 410)
            if amount is not None:
                pdf_totals['db'] = amount
            counts = [int(w['text']) for w in row['words'] if 405 <= w['x'] < 460 and w['text'].isdigit()]
            if len(counts) == 1:
                pdf_totals['db_count'] = counts[0]
            end_y = min(end_y, row['y'] - 1)
        elif re.search(r'\bSALDO\s+AKHIR\s*:', upper):
            amount = nominal_kolom(row, 285, 410)
            if amount is not None:
                pdf_totals['akhir'] = amount
            end_y = min(end_y, row['y'] - 1)
        elif re.search(r'\bSALDO\s+AWAL\s*:', upper):
            amount = nominal_kolom(row, 285, 410)
            if amount is not None:
                pdf_totals['awal'] = amount
            end_y = min(end_y, row['y'] - 1)

    candidates = []
    for idx, row in enumerate(rows):
        if not (start_y < row['y'] < end_y):
            continue
        possible = [w['text'] for w in row['words'] if 15 <= w['x'] < 82 and TANGGAL.fullmatch(w['text'])]
        if possible:
            candidates.append((idx, possible[0]))

    opening = None
    tx = []
    for pos, (first, ddmm) in enumerate(candidates):
        following = candidates[pos + 1][0] if pos + 1 < len(candidates) else len(rows)
        body = [r for r in rows[first:following] if r['y'] < end_y]
        if not body:
            continue
        first_row = body[0]
        desc = uraian_transaksi(body)
        if desc.upper().startswith('SALDO AWAL'):
            opening = nominal_kolom(first_row, 480, 595)
            continue
        try:
            d, m = [int(v) for v in ddmm.split('/')]
            dated = date(year, m, d)
        except ValueError as err:
            raise GagalBaca(f'Tanggal tidak valid: {ddmm}/{year}') from err
        if m != month:
            raise GagalBaca(f'Tanggal {ddmm} tidak sesuai periode {month:02d}/{year}.')
        mutation = nominal_kolom(first_row, 350, 460)
        if mutation is None:
            raise GagalBaca(f'Mutasi {ddmm} tidak berhasil dibaca.')
        debit_flag = any(w['text'].upper() == 'DB' for w in first_row['words'] if 430 <= w['x'] < 481)
        pdf_balance = nominal_kolom(first_row, 480, 595)
        if not desc:
            raise GagalBaca(f'Uraian transaksi {ddmm} tidak terbaca.')
        tx.append({
            'tanggal': dated,
            'uraian': desc,
            # Debet GL Bank = mutasi CR pada rekening koran
            'debet': Decimal('0') if debit_flag else mutation,
            'kredit': mutation if debit_flag else Decimal('0'),
            'saldo_pdf': pdf_balance,
            'halaman': page.page_number,
        })
    return tx, opening, pdf_totals


def baca_pdf(content: bytes, name: str):
    if not content.lstrip().startswith(b'%PDF'):
        raise GagalBaca('File bukan PDF valid.')
    try:
        with pdfplumber.open(BytesIO(content)) as pdf:
            text = '\n'.join((p.extract_text() or '') for p in pdf.pages)
            if len(text.strip()) < 60:
                raise GagalBaca('PDF berupa scan/gambar atau tidak punya teks; OCR belum didukung.')
            if 'SALDO AWAL' not in text.upper() or 'MUTASI' not in text.upper():
                raise GagalBaca('Format rekening koran belum dikenali (saat ini BCA e-statement).')
            year, month, account = metadata(text, name)
            transactions = []
            opening = None
            totals = {}
            for page in pdf.pages:
                events, start, footer = baca_halaman(page, year, month)
                transactions.extend(events)
                if opening is None and start is not None:
                    opening = start
                totals.update(footer)
    except GagalBaca:
        raise
    except Exception as err:
        raise GagalBaca(f'PDF tidak bisa diproses: {type(err).__name__}: {err}') from err

    if opening is None:
        opening = totals.get('awal')
    if opening is None:
        raise GagalBaca('Saldo awal tidak ditemukan; data belum bisa dihitung dengan benar.')
    if not transactions:
        raise GagalBaca('Tidak ada transaksi yang berhasil dibaca.')

    running = opening
    warnings = []
    for index, tx in enumerate(transactions, 1):
        running += tx['debet'] - tx['kredit']
        tx['saldo'] = running
        tx['urutan'] = index
        if tx['saldo_pdf'] is not None and tx['saldo_pdf'] != running:
            warnings.append(f'Transaksi {index} ({tx["tanggal"]:%d/%m/%Y}): saldo mutasi berbeda.')
    total_debet = sum((row['debet'] for row in transactions), Decimal(0))
    total_kredit = sum((row['kredit'] for row in transactions), Decimal(0))
    if 'cr' in totals and totals['cr'] != total_debet:
        warnings.append(f'Total CR PDF {totals["cr"]:,.2f} berbeda dari hasil {total_debet:,.2f}.')
    if 'db' in totals and totals['db'] != total_kredit:
        warnings.append(f'Total DB PDF {totals["db"]:,.2f} berbeda dari hasil {total_kredit:,.2f}.')
    if 'cr_count' in totals and totals['cr_count'] != sum(bool(t['debet']) for t in transactions):
        warnings.append('Jumlah transaksi CR tidak sesuai dengan ringkasan PDF.')
    if 'db_count' in totals and totals['db_count'] != sum(bool(t['kredit']) for t in transactions):
        warnings.append('Jumlah transaksi DB tidak sesuai dengan ringkasan PDF.')
    if 'akhir' in totals and totals['akhir'] != running:
        warnings.append(f'Saldo akhir PDF {totals["akhir"]:,.2f} berbeda dari hasil {running:,.2f}.')
    for label, description in [('cr','Mutasi CR'),('db','Mutasi DB'),('akhir','Saldo akhir')]:
        if label not in totals:
            warnings.append(f'Ringkasan {description} PDF tidak terbaca.')
    if account == 'TIDAK TERBACA':
        warnings.append('Nomor rekening tidak ditemukan.')
    return {
        'file': Path(name).name, 'rekening': account,
        'year': year, 'month': month, 'saldo_awal': opening,
        'saldo_akhir': running, 'debet': total_debet, 'kredit': total_kredit,
        'transactions': transactions, 'warnings': warnings,
    }


def terima_upload(files):
    """Terima PDF atau ZIP (dibaca di RAM). Beri error eksplisit bila ada masalah."""
    candidates = []
    for uploaded in files:
        data = uploaded.getvalue()
        if uploaded.name.lower().endswith('.pdf'):
            candidates.append((Path(uploaded.name).name, data))
        elif uploaded.name.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(BytesIO(data)) as archive:
                    infos = [i for i in archive.infolist()
                             if not i.is_dir() and i.filename.lower().endswith('.pdf')]
                    if len(infos) + len(candidates) > MAX_PDF:
                        raise GagalBaca(f'ZIP berisi lebih dari {MAX_PDF} PDF.')
                    if any(i.file_size > 20 * 1024 * 1024 for i in infos):
                        raise GagalBaca('Ada PDF di ZIP yang ukurannya lebih dari 20 MB.')
                    if sum(i.file_size for i in infos) + sum(len(p) for _, p in candidates) > MAX_BYTES:
                        raise GagalBaca('Total isi PDF terlalu besar (maksimum 80 MB).')
                    for info in infos:
                        candidates.append((Path(info.filename).name, archive.read(info)))
            except zipfile.BadZipFile as err:
                raise GagalBaca(f'{uploaded.name}: ZIP rusak/tidak valid.') from err
    if len(candidates) > MAX_PDF:
        raise GagalBaca(f'Terlalu banyak PDF (batas {MAX_PDF}).')
    if sum(len(p) for _, p in candidates) > MAX_BYTES:
        raise GagalBaca('Total ukuran PDF melebihi 80 MB.')
    return candidates


# ====== 2. MENGISI TEMPLATE EXCEL ASLI (FORMAT BANK) ======
# Bank_template.xlsx adalah versi KOSONG dan anonim dari file Bank(1).xlsx.
# Font, warna, border, tinggi header, dan lebar kolom tetap dari file contoh.
# Semua transaksi masuk ke SATU sheet GL, hanya 10 kolom A:J.

TEMPLATE = Path(__file__).with_name('Bank_template.xlsx')
KOLOM = [
    'Tanggal', 'No. Bukti', 'Kode Akun', 'Kode Lawan Akun', 'Nama Akun',
    'Uraian', 'Debet', 'Kredit', 'Saldo', 'Nama Vendor',
]


def teks_aman_excel(value) -> str:
    """Singkirkan karakter XML ilegal, agar hasil XLSX dapat dibuka Excel."""
    text = str(value)
    text = ''.join(
        ch for ch in text
        if (ord(ch) in (9, 10, 13) or 32 <= ord(ch) <= 0xD7FF
            or 0xE000 <= ord(ch) <= 0xFFFD or 0x10000 <= ord(ch) <= 0x10FFFF)
        and (ord(ch) & 0xFFFF) not in (0xFFFE, 0xFFFF)
    )[:32767]
    # Penting: jangan izinkan keterangan dari PDF menjadi rumus Excel.
    if text.startswith(('=', '+', '-', '@')):
        text = "'" + text[:32766]
    return text


def buat_excel(statements, company: str, code: str) -> bytes:
    """Konversi PDF ke 1 sheet GL menggunakan style ASLI Bank(1).xlsx."""
    if not TEMPLATE.is_file():
        raise GagalBaca(
            'Bank_template.xlsx tidak ditemukan. Pastikan file ini berada '
            'di folder yang sama dengan app.py (termasuk saat deploy).'
        )
    try:
        book = load_workbook(TEMPLATE, keep_links=False)
        if len(book.worksheets) != 1:
            raise GagalBaca('Template harus berisi hanya satu sheet.')
        ws = book.active
        ws.title = 'GL'
        header = [ws.cell(5, col).value for col in range(1, 11)]
        if header != KOLOM:
            raise GagalBaca('Header Bank_template.xlsx tidak sesuai contoh Buku Bank.')

        # Salin style asli sebelum membuat ulang data. _style berisi seluruh
        # atribut font, color, border, alignment, protection, dan num format.
        saldo_style = [copy(ws.cell(7, col)._style) for col in range(1, 11)]
        mutasi_style = [copy(ws.cell(8, col)._style) for col in range(1, 11)]
        original_row7_height = ws.row_dimensions[7].height
        original_row8_height = ws.row_dimensions[8].height

        # Isi lama template memang kosong, namun bersihkan kembali agar
        # jika template diganti pengguna tidak ada transaksi yang tersisa.
        for row in ws.iter_rows(min_row=7, max_row=ws.max_row, max_col=10):
            for cell in row:
                cell.value = None
        ws['A1'] = teks_aman_excel(company.strip() or 'BUKU BANK')
        ws['A2'] = 'Buku Bank'
        year_set = sorted({s['year'] for s in statements})
        ws['A3'] = year_set[0] if len(year_set) == 1 else ', '.join(map(str, year_set))

        bank_code = (int(code) if code.isdigit() and len(code) <= 14
                     and not code.startswith('0') else teks_aman_excel(code))
        row_num = 7
        statements = sorted(statements, key=lambda s: (
            s['rekening'], s['year'], s['month'], s['file']))
        for index, stmt in enumerate(statements):
            # Satu baris SALDO BANK AWAL untuk tiap periode/rekening.
            # Bukan mutasi, sehingga tidak dihitung sebagai debit/kredit.
            for col in range(1, 11):
                ws.cell(row_num, col)._style = copy(saldo_style[col - 1])
            ws.cell(row_num, 6).value = (
                'SALDO BANK AWAL' if len(statements) == 1 else
                f'SALDO BANK AWAL - {stmt["rekening"]} - {stmt["month"]:02d}/{stmt["year"]}'
            )
            ws.cell(row_num, 9).value = float(stmt['saldo_awal'])
            ws.row_dimensions[row_num].height = original_row7_height
            row_num += 1

            for item in stmt['transactions']:
                values = [
                    item['tanggal'], None, bank_code, None, None,
                    teks_aman_excel(item['uraian']),
                    float(item['debet']) if item['debet'] else None,
                    float(item['kredit']) if item['kredit'] else None,
                    float(item['saldo']), None,
                ]
                for col, value in enumerate(values, 1):
                    cell = ws.cell(row_num, col)
                    cell._style = copy(mutasi_style[col - 1])
                    cell.value = value
                # Multi-baris keterangan tetap lengkap. Styling tidak berubah.
                lines = item['uraian'].count('\n') + 1
                ws.row_dimensions[row_num].height = (
                    min(110, max(17, 14 * lines + 4)) if lines > 1
                    else original_row8_height
                )
                row_num += 1

        # Hapus area sisa kosong dari template agar satu-satunya tabel
        # berakhir di transaksi terakhir; tidak ada data bank contoh.
        if ws.max_row >= row_num:
            ws.delete_rows(row_num, ws.max_row - row_num + 1)
        # Header, font, border, format tanggal/angka, dan lebar kolom
        # sengaja TIDAK diubah agar sesuai persis dengan contoh.

        buffer = BytesIO()
        book.save(buffer)
        result = buffer.getvalue()
        book.close()

        # Pemeriksaan internal struktur agar Excel tidak meminta repair.
        with zipfile.ZipFile(BytesIO(result)) as archive:
            broken = archive.testzip()
            if broken:
                raise GagalBaca(f'File XLSX korup: {broken}')
            for name in archive.namelist():
                if name.endswith(('.xml', '.rels')):
                    ET.fromstring(archive.read(name))
        checked = load_workbook(BytesIO(result), read_only=True, keep_links=False)
        try:
            if checked.sheetnames != ['GL']:
                raise GagalBaca('Output bukan satu sheet GL.')
            if [checked['GL'].cell(5, c).value for c in range(1, 11)] != KOLOM:
                raise GagalBaca('Header output berubah.')
            if checked['GL'].max_row != row_num - 1:
                raise GagalBaca('Jumlah baris output tidak sesuai transaksi.')
        finally:
            checked.close()
        return result
    except GagalBaca:
        raise
    except Exception as exc:
        raise GagalBaca(f'Gagal membuat Excel: {type(exc).__name__}: {exc}') from exc


# ====== 3. STREAMLIT: CUKUP UPLOAD -> PROSES -> DOWNLOAD ======
def rupiah(n):
    return 'Rp ' + f'{n:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')


def main():
    st.set_page_config(page_title='ConvertRK – PDF ke GL', page_icon='📑', layout='wide')
    st.title('ConvertRK — PDF Rekening Koran ke GL')
    st.caption('Upload PDF BCA atau ZIP → konversi semua mutasi → Excel 1 sheet GL, dengan font, warna, dan format Buku Bank contoh. Tanpa mapping akun.')

    left, right = st.columns([3, 1])
    with left:
        company = st.text_input('Nama perusahaan (opsional)', value='PT Cendrawasih Dinamika Mandiri')
    with right:
        code = st.text_input('Kode akun bank', value='111', help='Kolom Kode Akun. Lawan akun dibiarkan kosong.')

    uploads = st.file_uploader('Pilih PDF rekening koran atau ZIP',
                               type=['pdf', 'zip'], accept_multiple_files=True)
    if not uploads:
        st.info('Format saat ini: BCA e-statement PDF berbasis teks seperti contoh Januari 2026. PDF scan belum didukung.')
        return

    signature = (tuple((f.name, f.size, hashlib.sha256(f.getvalue()).hexdigest()) for f in uploads),
                 company, code)
    if st.button('Konversi ke GL', type='primary'):
        st.session_state.pop('convertRK_result', None)
        try:
            files = terima_upload(uploads)
            if not files:
                raise GagalBaca('Tidak ada PDF ditemukan.')
            statements, failures, skipped = [], [], []
            found_periods = set()
            for filename, content in files:
                try:
                    stmt = baca_pdf(content, filename)
                    key = (stmt['rekening'], stmt['year'], stmt['month'])
                    if key in found_periods:
                        skipped.append(f'{filename}: periode/rekening sama dengan file sebelumnya (tidak digandakan).')
                        continue
                    found_periods.add(key)
                    statements.append(stmt)
                except GagalBaca as exc:
                    failures.append(f'{filename}: {exc}')
            if not statements:
                raise GagalBaca('Tidak ada PDF yang berhasil diproses.\n' + '\n'.join(failures))
            book = buat_excel(statements, company, code.strip() or '111')
            st.session_state['convertRK_result'] = (signature, statements, book, failures, skipped)
        except GagalBaca as exc:
            st.error(str(exc))

    result = st.session_state.get('convertRK_result')
    if not result:
        return
    saved_signature, statements, book, failures, skipped = result
    if signature != saved_signature:
        st.warning('Unggahan/pengaturan berubah. Klik **Konversi ke GL** lagi untuk membuat Excel terbaru.')
        return

    n = sum(len(s['transactions']) for s in statements)
    incoming = sum((s['debet'] for s in statements), Decimal(0))
    outgoing = sum((s['kredit'] for s in statements), Decimal(0))
    m1, m2, m3 = st.columns(3)
    m1.metric('Mutasi terbaca', f'{n:,}')
    m2.metric('Total mutasi masuk / Debet', rupiah(incoming))
    m3.metric('Total mutasi keluar / Kredit', rupiah(outgoing))

    if failures:
        st.error(f'{len(failures)} PDF GAGAL diproses — transaksi dari file ini BELUM masuk Excel:')
        for msg in failures:
            st.write('• ' + msg)
    if skipped:
        st.warning('File periode ganda tidak diimpor: ' + '; '.join(skipped))

    warnings = [(s['file'], w) for s in statements for w in s['warnings']]
    if warnings:
        with st.expander(f'Perhatian: {len(warnings)} temuan perbandingan saldo PDF', expanded=True):
            st.warning('Excel tetap dibuat, tetapi periksa PDF terkait karena ada potensi mutasi terlewat atau salah baca.')
            for filename, message in warnings:
                st.write(f'• {filename}: {message}')
    else:
        st.success('Mutasi dan saldo akhir seluruh PDF yang berhasil dibaca cocok dengan ringkasan rekening koran.')

    with st.expander('Ringkasan per rekening koran', expanded=True):
        st.dataframe([
            {'File': s['file'], 'Rekening': s['rekening'], 'Periode': f'{s["month"]:02d}/{s["year"]}',
             'Jumlah Mutasi': len(s['transactions']), 'Saldo Awal': float(s['saldo_awal']),
             'Debet': float(s['debet']), 'Kredit': float(s['kredit']),
             'Saldo Akhir': float(s['saldo_akhir']),
             'Status': 'SESUAI' if not s['warnings'] else 'PERLU CEK'}
            for s in statements
        ], hide_index=True, use_container_width=True)

    preview = []
    for s in statements:
        for tx in s['transactions']:
            preview.append({'Tanggal': tx['tanggal'].strftime('%d/%m/%Y'),
                            'Uraian': tx['uraian'].replace('\n', ' | '),
                            'Debet': float(tx['debet']), 'Kredit': float(tx['kredit']),
                            'Saldo': float(tx['saldo']), 'No. Rekening': s['rekening']})
    st.caption('Pratinjau 30 transaksi pertama. Kode lawan akun dan nama akun pada Excel sengaja kosong.')
    st.dataframe(preview[:30], hide_index=True, use_container_width=True)

    st.download_button('⬇️ Download GL Excel (1 Sheet)', data=book,
                       file_name='GL_Rekening_Koran.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                       type='primary')
    st.caption('Untuk data rahasia, disarankan jalankan secara lokal. Jangan unggah rekening koran ke repository GitHub.')


if __name__ == '__main__':
    main()
