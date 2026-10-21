"""Parsing rekening koran BCA PDF berbasis posisi teks (tidak menggunakan OCR)."""
from __future__ import annotations

from dataclasses import dataclass, asdict
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
import zipfile
from io import BytesIO

import fitz  # PyMuPDF

MONTHS = {
    'JANUARI': 1, 'FEBRUARI': 2, 'MARET': 3, 'APRIL': 4, 'MEI': 5,
    'JUNI': 6, 'JULI': 7, 'AGUSTUS': 8, 'SEPTEMBER': 9, 'OKTOBER': 10,
    'NOVEMBER': 11, 'DESEMBER': 12,
}
MONEY_RX = re.compile(r'^-?\d{1,3}(?:,\d{3})*\.\d{2}$|^-?\d+\.\d{2}$')
DATE_RX = re.compile(r'^\d{2}/\d{2}$')


class ParsingError(ValueError):
    pass


def rupiah(text: str) -> Decimal:
    clean = (text or '').replace(',', '').strip()
    if not MONEY_RX.fullmatch(clean):
        raise ParsingError(f'Nominal tidak valid: {text!r}')
    try:
        return Decimal(clean)
    except InvalidOperation as exc:
        raise ParsingError(f'Nominal tidak valid: {text!r}') from exc


def _text_line(words: list[tuple]) -> str:
    return ' '.join(w[4] for w in sorted(words, key=lambda t: t[0])).strip()


def _lines(words: list[tuple]) -> list[dict]:
    """Pecah kata menjadi baris berdasarkan posisi y, bukan urutan PDF stream."""
    grouped = {}
    for word in words:
        y = round(word[1] / 2) * 2
        grouped.setdefault(y, []).append(word)
    return [{'y': y, 'words': sorted(ws, key=lambda v: v[0])}
            for y, ws in sorted(grouped.items())]


def _details(lines: list[dict], *, min_x=84, max_x=302) -> str:
    chunks = []
    for line in lines:
        words = [w for w in line['words'] if min_x <= w[0] < max_x]
        # BCA menyimpan nama transaksi di x=88 dan referensi di x=194.
        # Walaupun ada pada baris yang sama, pada contoh Excel keduanya dipisah newline.
        groups = [[w for w in words if w[0] < 190], [w for w in words if w[0] >= 190]]
        for part in groups:
            if part:
                chunk = _text_line(part).strip()
                if chunk:
                    chunks.append(chunk)
    return '\n'.join(chunks)


def _one_money(line: dict, lo: float, hi: float) -> Decimal | None:
    money = [w[4] for w in line['words'] if lo <= w[0] < hi and MONEY_RX.fullmatch(w[4])]
    if len(money) > 1:
        raise ParsingError(f'Terdapat beberapa nominal pada satu kolom: {money}')
    return rupiah(money[0]) if money else None


def _statement_meta(fulltext: str, filename: str) -> tuple[int, int, str, str]:
    period = re.search(r'\b(JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER)\s+(20\d{2})\b', fulltext, re.I)
    if period:
        month, year = MONTHS[period.group(1).upper()], int(period.group(2))
    else:
        fallback = re.search(r'(?<!\d)(20\d{2})(0[1-9]|1[0-2])(?!\d)', filename)
        if fallback is None:
            raise ParsingError('Bulan/tahun rekening koran tidak ditemukan pada PDF maupun nama file YYYYMM.')
        year, month = int(fallback.group(1)), int(fallback.group(2))
    # Hanya bagian header, bukan nomor transfer di bagian mutasi.
    header = fulltext.split('CATATAN:')[0]
    acct_match = re.search(r'NO\.\s*REKENING\s*:\s*([\d\s]{8,22})', header, re.I)
    account = re.sub(r'\D', '', acct_match.group(1)) if acct_match else ''
    if not account:
        acct_match = re.search(r'NO\.\s*REKENING\s*:?\s*\n\s*:?\s*\n\s*(\d{8,16})', header, re.I)
        account = acct_match.group(1) if acct_match else ''
    currency = 'IDR'
    currency_match = re.search(r'MATA\s*UANG\s*:?\s*\n\s*:?\s*\n\s*([A-Z]{3})', header, re.I)
    if currency_match:
        currency = currency_match.group(1).upper()
    return year, month, account, currency


def _page_events(page: fitz.Page, year: int, month: int) -> tuple[list[dict], Decimal | None, dict]:
    all_lines = _lines(page.get_text('words', sort=False))
    # Cari baris total BCA: SALDO AWAL :, MUTASI CR :, MUTASI DB :, SALDO AKHIR :.
    summary_y = None
    summary = {}
    for ln in all_lines:
        content = _text_line(ln['words']).upper()
        if 'SALDO AWAL' in content and any(195 <= w[0] <= 285 for w in ln['words']):
            summary_y = min(summary_y, ln['y'] - 1) if summary_y is not None else ln['y'] - 1
        if 'MUTASI CR' in content or 'MUTASI DB' in content or 'SALDO AKHIR' in content:
            label = 'cr' if 'MUTASI CR' in content else 'db' if 'MUTASI DB' in content else 'akhir'
            vals = [w[4] for w in ln['words'] if 290 <= w[0] < 395 and MONEY_RX.fullmatch(w[4])]
            if vals:
                summary[label] = rupiah(vals[0])
                if label == 'cr':
                    summary_y = min(summary_y, ln['y'] - 10) if summary_y is not None else ln['y'] - 10
    candidates = []
    for idx, ln in enumerate(all_lines):
        if ln['y'] < 195 or (summary_y is not None and ln['y'] > summary_y):
            continue
        ds = [w for w in ln['words'] if w[0] < 80 and DATE_RX.fullmatch(w[4])]
        if ds:
            candidates.append((idx, ds[0][4]))
    transactions = []
    opening = None
    for pos, (first_idx, day_month) in enumerate(candidates):
        following = candidates[pos + 1][0] if pos + 1 < len(candidates) else len(all_lines)
        event_lines = all_lines[first_idx:following]
        if summary_y is not None:
            event_lines = [ln for ln in event_lines if ln['y'] <= summary_y]
        if not event_lines:
            continue
        first = event_lines[0]
        header_desc = _details([first]).upper()
        if 'SALDO AWAL' in header_desc:
            bal = _one_money(first, 485, 590)
            if bal is not None:
                opening = bal
            continue
        try:
            dd, mm = (int(v) for v in day_month.split('/'))
            the_date = date(year, mm, dd)
        except ValueError as exc:
            raise ParsingError(f'Tanggal tidak valid: {day_month}, {month}/{year}') from exc
        if mm != month:
            raise ParsingError(f'Tanggal {day_month} berbeda dari periode laporan {month:02d}/{year}.')
        mutation = _one_money(first, 354, 460)
        if mutation is None:
            raise ParsingError(f'Nominal mutasi tidak ditemukan pada {day_month}.')
        is_db = any(w[4].upper() == 'DB' for w in first['words'] if 430 <= w[0] <= 480)
        balance = _one_money(first, 485, 590)
        description = _details(event_lines)
        if not description:
            raise ParsingError(f'Keterangan kosong pada {day_month}.')
        transactions.append({
            'tanggal': the_date.isoformat(),
            'uraian': description,
            'debet': Decimal('0') if is_db else mutation,
            'kredit': mutation if is_db else Decimal('0'),
            'saldo_tertera': balance,
            'page': page.number + 1,
        })
    return transactions, opening, summary


def parse_bca_pdf(payload: bytes, filename: str) -> dict:
    if not payload.startswith(b'%PDF'):
        raise ParsingError(f'{filename}: bukan berkas PDF.')
    try:
        document = fitz.open(stream=payload, filetype='pdf')
    except Exception as exc:
        raise ParsingError(f'{filename}: PDF rusak, tidak valid, atau terkunci.') from exc
    try:
        fulltext = '\n'.join(page.get_text(sort=True) for page in document)
        if len(fulltext.strip()) < 80:
            raise ParsingError(f'{filename}: PDF tampaknya berupa scan/gambar. OCR belum didukung.')
        if 'REKENING GIRO' not in fulltext.upper() and 'SALDO AWAL' not in fulltext.upper():
            raise ParsingError(f'{filename}: format belum dikenali. Saat ini mendukung rekening koran BCA seperti contoh.')
        year, month, account, currency = _statement_meta(fulltext, filename)
        if not account:
            raise ParsingError(f'{filename}: nomor rekening tidak terbaca.')
        transactions = []
        opening = None
        summary = {}
        for page in document:
            tx, initial, footer = _page_events(page, year, month)
            transactions.extend(tx)
            if opening is None and initial is not None:
                opening = initial
            summary.update(footer)
        if opening is None:
            raise ParsingError(f'{filename}: saldo awal tidak terbaca dari PDF.')
        if not transactions:
            raise ParsingError(f'{filename}: tidak ditemukan transaksi mutasi.')
        calc = opening
        anomalies = []
        for number, tx in enumerate(transactions, 1):
            calc += tx['debet'] - tx['kredit']
            tx['saldo_hitung'] = calc
            tx['no_urut'] = number
            if tx['saldo_tertera'] is not None and tx['saldo_tertera'] != calc:
                anomalies.append(f"Transaksi #{number} ({tx['tanggal']}): saldo PDF {tx['saldo_tertera']}, hasil hitung {calc}.")
        debet = sum((t['debet'] for t in transactions), Decimal('0'))
        kredit = sum((t['kredit'] for t in transactions), Decimal('0'))
        if 'cr' in summary and summary['cr'] != debet:
            anomalies.append(f"Total mutasi CR: PDF {summary['cr']}, hasil parse {debet}.")
        if 'db' in summary and summary['db'] != kredit:
            anomalies.append(f"Total mutasi DB: PDF {summary['db']}, hasil parse {kredit}.")
        if 'akhir' in summary and summary['akhir'] != calc:
            anomalies.append(f"Saldo akhir: PDF {summary['akhir']}, hasil hitung {calc}.")
        for field, label in [('cr','Mutasi CR'), ('db','Mutasi DB'), ('akhir','Saldo akhir')]:
            if field not in summary:
                anomalies.append(f'{label} pada ringkasan PDF tidak dapat dibaca; perlu diperiksa manual.')
        # Jangan ubah saldo PDF secara diam-diam jika tidak konsisten.
        return {
            'file': Path(filename).name,
            'bank': 'BCA', 'rekening': account, 'currency': currency,
            'year': year, 'month': month,
            'saldo_awal': opening,
            'debet': debet, 'kredit': kredit,
            'saldo_akhir': calc,
            'saldo_akhir_pdf': summary.get('akhir'),
            'mutasi_cr_pdf': summary.get('cr'),
            'mutasi_db_pdf': summary.get('db'),
            'errors': anomalies,
            'transactions': transactions,
        }
    finally:
        document.close()


def expand_uploads(uploaded: list) -> list[tuple[str, bytes]]:
    """Baca upload PDF dan ZIP di RAM, tanpa mengekstrak ZIP ke disk."""
    found = []
    used_names = set()
    total_bytes = 0
    for item in uploaded:
        data = item.getvalue()
        if item.name.lower().endswith('.pdf'):
            candidates = [(Path(item.name).name, data)]
        elif item.name.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(BytesIO(data)) as archive:
                    infos = [i for i in archive.infolist() if not i.is_dir() and i.filename.lower().endswith('.pdf')
                             and not Path(i.filename).name.startswith('.')]
                    if len(infos) > 50 or sum(i.file_size for i in infos) > 60 * 1024 * 1024:
                        raise ParsingError('ZIP terlalu besar: maksimum 50 PDF dan 60 MB data PDF.')
                    candidates = [(Path(i.filename).name, archive.read(i)) for i in infos]
            except zipfile.BadZipFile as exc:
                raise ParsingError(f'{item.name}: file ZIP tidak valid.') from exc
        else:
            continue
        for name, payload in candidates:
            if not name or len(payload) > 20 * 1024 * 1024:
                raise ParsingError(f'{name}: PDF terlalu besar (maksimum 20 MB/file).')
            total_bytes += len(payload)
            if total_bytes > 60 * 1024 * 1024:
                raise ParsingError('Total file PDF lebih dari 60 MB.')
            if name in used_names:
                raise ParsingError(f'Nama file PDF duplikat: {name}. Ubah namanya terlebih dahulu.')
            found.append((name, payload))
            used_names.add(name)
    if len(found) > 50:
        raise ParsingError('Maksimum 50 file PDF dalam sekali pemrosesan.')
    return found
