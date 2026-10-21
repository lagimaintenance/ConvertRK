"""ConvertRK - aplikasi Streamlit rekening koran BCA PDF -> Buku Bank Excel.

Menjalankan secara lokal: streamlit run app.py
Deploy gratis: Streamlit Community Cloud, Main file: app.py
"""
from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path
import hashlib

import pandas as pd  # Hanya untuk tabel editor interaktif, bukan menulis/membaca Excel.
import streamlit as st

from bank_parser import ParsingError, expand_uploads, parse_bca_pdf
from coa_default import DEFAULT_COA
from xlsx_engine import (
    as_decimal, generate_xlsx, normalize_code, read_coa, read_example, suggest_account,
)

BASE = Path(__file__).resolve().parent
st.set_page_config(page_title='ConvertRK | PDF ke Buku Bank', page_icon='🏦', layout='wide',
                   initial_sidebar_state='expanded')
st.markdown('''<style>
.block-container {padding-top:1.8rem; max-width:1480px;}
h1 {color:#17365A; font-weight:750; letter-spacing:-0.025em;}
div[data-testid="stMetric"] {border:1px solid #dbe4ed;border-radius:12px;padding:14px 17px;background:#f8fafc;}
.stButton button[kind="primary"] {background-color:#164e63; border-color:#164e63;}
[data-testid="stFileUploader"] {border-radius:12px;}
.small-caption {color:#64748b;font-size:0.88rem;}
</style>''', unsafe_allow_html=True)


def local_bytes(filename: str) -> bytes | None:
    candidate = BASE / filename
    try:
        return candidate.read_bytes() if candidate.is_file() else None
    except OSError:
        return None


def account_config(coa_upload, example_upload):
    coa = DEFAULT_COA.copy()
    coa_from = 'COA bawaan dari Kode Akun.xlsx yang Anda berikan'
    if coa_upload is not None:
        coa = read_coa(coa_upload.getvalue())
        coa_from = f'COA dari upload: {coa_upload.name}'
    elif local_bytes('Kode Akun.xlsx'):
        coa = read_coa(local_bytes('Kode Akun.xlsx'))
        coa_from = 'COA lokal: Kode Akun.xlsx'
    example = None
    if example_upload is not None:
        example = read_example(example_upload.getvalue())
    elif local_bytes('Bank(1).xlsx'):
        example = read_example(local_bytes('Bank(1).xlsx'))
    return coa, example, coa_from


def display_rp(value) -> str:
    val = as_decimal(value)
    return 'Rp ' + f'{val:,.0f}'.replace(',', '.')


def to_editor_rows(records: list[dict]):
    return pd.DataFrame([{  # Hanya satu tampilan tabel untuk review dan koreksi, tidak mengubah mutasi.
        'File': r['file'], 'Rekening':r['rekening'], 'Tanggal':date.fromisoformat(r['tanggal']),
        'Uraian':r['uraian'], 'Debet':float(r['debet']), 'Kredit':float(r['kredit']),
        'Kode Lawan Akun':r['kode_lawan'] or None, 'Verifikasi':bool(r['verified']),
        'No. Bukti':r.get('no_bukti', ''), 'Nama Vendor':r.get('vendor', ''),
        'Keterangan Sistem':r['catatan'],
    } for r in records])


def derive_export_rows(df, raw, coa):
    edits = df.to_dict('records')
    if len(edits) != len(raw):
        raise ValueError('Jumlah transaksi berubah. Jalankan Proses ulang.')
    result=[]
    for rec, original in zip(edits, raw):
        code = normalize_code(rec.get('Kode Lawan Akun'))
        if code not in coa and code != '':
            raise ValueError(f"Kode akun {code} tidak terdapat di COA pada {original['tanggal']}.")
        approved=bool(rec.get('Verifikasi'))
        if not code:
            approved=False
        if rec.get('Kode Lawan Akun') != (original.get('kode_lawan') or None):
            why='Akun ditetapkan / diubah pengguna'
        else:
            why=original['catatan']
        record={**original, 'kode_lawan':code, 'verified':approved, 'catatan':why,
                'no_bukti':str(rec.get('No. Bukti') or '').strip(),
                'vendor':str(rec.get('Nama Vendor') or '').strip()}
        result.append(record)
    return result


st.title('ConvertRK')
st.caption('Konversi rekening koran BCA PDF menjadi Buku Bank Excel • Kode akun • Pemeriksaan saldo • Koreksi klasifikasi')

with st.sidebar:
    st.subheader('1. Pengaturan Buku Bank')
    company=st.text_input('Nama perusahaan', value='PT Cendrawasih Dinamika Mandiri')
    st.divider()
    st.subheader('2. File Referensi (opsional)')
    coa_upload=st.file_uploader('Kode Akun.xlsx', type=['xlsx'], key='upload_coa',
            help='Tidak perlu jika memakai 46 kode akun bawaan. Anda boleh mengganti COA dengan unggahan terbaru.')
    example_upload=st.file_uploader('Bank(1).xlsx (contoh klasifikasi)', type=['xlsx'], key='upload_example',
            help='Jika diunggah, aplikasi mencocokkan jurnal yang persis sama dengan contoh, termasuk Januari 2026.')
    try:
        coa, example, coa_from=account_config(coa_upload,example_upload)
    except (ValueError, KeyError, OSError) as e:
        st.error(str(e)); st.stop()
    st.caption(f'{len(coa)} akun • {coa_from}')
    options=[code for code in sorted(coa, key=lambda x: int(x) if x.isdigit() else 99999)]
    bank_default=options.index('111') if '111' in options else 0
    main_bank=st.selectbox('Kode akun bank',options,index=bank_default,format_func=lambda k:f'{k} — {coa[k]["name"]}')
    st.divider()
    st.info('Data rekening koran diproses dalam sesi aplikasi. Jangan unggah rekening koran atau file contoh transaksi ke GitHub publik.')

st.subheader('Unggah Rekening Koran')
uploads=st.file_uploader('PDF rekening koran BCA atau ZIP kumpulan PDF',
                         type=['pdf','zip'],accept_multiple_files=True,
                         help='Maksimum 50 PDF / 60 MB. PDF perlu mempunyai teks yang dapat diseleksi (bukan hasil scan gambar).')

with st.expander('Aturan pemetaan tambahan (opsional)'):
    st.write('Kata kunci dapat menetapkan akun secara otomatis. Aturan ini diprioritaskan setelah pencocokan transaksi identik dengan contoh.')
    rules_df=st.data_editor(pd.DataFrame([{'Kata kunci':'','Kode Akun':''}]),
          num_rows='dynamic',hide_index=True,key='rule_table',use_container_width=True,
          column_config={'Kode Akun':st.column_config.SelectboxColumn('Kode Akun',options=['']+options)})
    rules=[]
    for entry in rules_df.to_dict('records'):
        phrase=str(entry.get('Kata kunci') or '').strip()
        code=normalize_code(entry.get('Kode Akun'))
        if phrase and code in coa:
            rules.append((phrase,code))

current_signature=tuple((f.name, f.size, hashlib.sha256(f.getvalue()).hexdigest()[:12]) for f in (uploads or []))
process=st.button('Proses rekening koran', type='primary', disabled=not uploads)
if process:
    # Jangan pertahankan hasil lama bila ekstraksi PDF baru gagal.
    st.session_state.pop('bank_statements',None)
    st.session_state.pop('bank_records',None)
    try:
        raw_files=expand_uploads(uploads)
        if not raw_files: raise ParsingError('Tidak ada PDF untuk diproses.')
        statements=[]
        records=[]
        used_periods=set()
        for name,data in raw_files:
            s=parse_bca_pdf(data,name)
            key=(s['rekening'],s['year'],s['month'])
            if key in used_periods:
                raise ParsingError(f'Ditemukan lebih dari satu file periode {s["month"]:02d}/{s["year"]} untuk rekening {s["rekening"]}. Gabungkan sebagai satu PDF terlebih dahulu atau hapus duplikat.')
            used_periods.add(key)
            statements.append(s)
            for item in s['transactions']:
                code,verified,reason=suggest_account(item,coa,example,rules)
                records.append({**item,'file':s['file'],'rekening':s['rekening'],
                                'kode_lawan':code,'verified':verified,'catatan':reason,
                                'no_bukti':'','vendor':''})
        # Urutan rekening/periode/file/no urut; akumulasi mengikuti sumber PDF.
        records.sort(key=lambda r:(r['rekening'],r['tanggal'],r['file'],r['no_urut']))
        statements.sort(key=lambda s:(s['rekening'],s['year'],s['month']))
        st.session_state['bank_statements']=statements
        st.session_state['bank_records']=records
        st.session_state['bank_files_signature']=current_signature
        st.session_state['bank_editor_version']=st.session_state.get('bank_editor_version',0)+1
        st.success(f'Berhasil membaca {len(statements)} PDF dan {len(records)} transaksi.')
    except (ParsingError, ValueError, OSError) as exc:
        st.error(f'Gagal memproses: {exc}')

if 'bank_statements' in st.session_state and 'bank_records' in st.session_state:
    statements=st.session_state['bank_statements']
    records=st.session_state['bank_records']
    st.divider()
    st.subheader('Hasil ekstraksi dan pemeriksaan')
    stale = current_signature != st.session_state.get('bank_files_signature')
    if stale:
        st.warning('Daftar PDF berubah atau dihapus. Hasil di bawah berasal dari pemrosesan sebelumnya; klik Proses rekening koran sebelum mengunduh.')
    debit=sum((s['debet'] for s in statements),Decimal(0))
    credit=sum((s['kredit'] for s in statements),Decimal(0))
    errors=[(s['file'],warning) for s in statements for warning in s['errors']]
    a,b,c,d=st.columns(4)
    a.metric('Transaksi',f'{len(records):,}')
    b.metric('Mutasi Masuk (Debet)',display_rp(debit))
    c.metric('Mutasi Keluar (Kredit)',display_rp(credit))
    d.metric('Rekonsiliasi PDF', 'Sesuai' if not errors else f'{len(errors)} temuan')
    summary=[]
    for s in statements:
        summary.append({'File':s['file'],'Rekening':s['rekening'],
                        'Bulan':f'{s["month"]:02d}/{s["year"]}',
                        'Saldo Awal':float(s['saldo_awal']),
                        'Debet':float(s['debet']),'Kredit':float(s['kredit']),
                        'Saldo Akhir':float(s['saldo_akhir']),
                        'Status':'COCOK' if not s['errors'] else 'PERLU CEK'})
    with st.expander('Rincian rekonsiliasi per PDF', expanded=bool(errors)):
        st.dataframe(pd.DataFrame(summary),hide_index=True,use_container_width=True)
        for filename,message in errors:
            st.error(f'{filename}: {message}')

    st.subheader('Koreksi kode lawan akun')
    st.write('Edit **Kode Lawan Akun**, **No. Bukti** dan **Nama Vendor**. Centang **Verifikasi** sesudah memeriksa bukti transaksi. Kolom debit/kredit hanya-baca agar angka rekening koran tidak berubah.')
    editor_key=f"transactions_{st.session_state.get('bank_editor_version',0)}"
    edited=st.data_editor(to_editor_rows(records),key=editor_key,hide_index=True,
                          num_rows='fixed',use_container_width=True,height=540,
                          disabled=['File','Rekening','Tanggal','Uraian','Debet','Kredit','Keterangan Sistem'],
                          column_config={
                              'File':st.column_config.TextColumn(width='medium'),
                              'Uraian':st.column_config.TextColumn(width='large'),
                              'Debet':st.column_config.NumberColumn(format='%.2f'),
                              'Kredit':st.column_config.NumberColumn(format='%.2f'),
                              'Tanggal':st.column_config.DateColumn(format='DD/MM/YYYY'),
                              'Kode Lawan Akun':st.column_config.SelectboxColumn(
                                  'Kode Lawan Akun',options=options,required=False,
                                  help='Hanya memilih kode dari COA; nama akun otomatis terisi saat ekspor.'),
                              'Verifikasi':st.column_config.CheckboxColumn(
                                  'Verifikasi',help='Centang bila klasifikasi akun sudah ditinjau dan benar.'),
                          })
    try:
        final_rows=derive_export_rows(edited,records,coa)
        review=[r for r in final_rows if not r['verified']]
        if review:
            st.warning(f'{len(review)} transaksi belum diverifikasi. Tetap dapat diekspor dan akan dicantumkan di sheet Perlu Review.')
        else:
            st.success('Semua transaksi sudah memiliki kode akun yang diverifikasi.')
        if errors:
            allow=st.checkbox('Saya memahami bahwa ada selisih/kesalahan ekstraksi dan ingin mengunduh untuk pemeriksaan.')
        else:
            allow=True
        if allow and not stale:
            book=generate_xlsx(statements,final_rows,coa,company or 'Perusahaan',int(main_bank))
            start=min((s['year'],s['month']) for s in statements)
            finish=max((s['year'],s['month']) for s in statements)
            name=f'Buku_Bank_{start[0]}{start[1]:02d}-{finish[0]}{finish[1]:02d}.xlsx'
            st.download_button('Unduh Buku Bank Excel',data=book,file_name=name,
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                               type='primary',use_container_width=True)
        st.caption('File ekspor berisi: Bank (gabungan), sheet per bulan, Rekonsiliasi, Perlu Review, dan Kode Akun.')
    except Exception as exc:
        st.error(f'Tidak dapat membuat Excel: {exc}')
else:
    st.info('Pilih rekening koran PDF/ZIP di atas, lalu klik **Proses rekening koran**.')

with st.expander('Batasan & cara pemeriksaan'):
    st.markdown('''- **Format input**: rekening koran BCA PDF teks seperti contoh Januari 2026. Format bank lain belum otomatis dikenali.
- **PDF scan**: belum didukung OCR; gunakan e-statement yang teksnya bisa diseleksi.
- **Saldo**: sistem membandingkan saldo berjalan dengan saldo pada PDF dan mutasi total di bagian bawah.
- **Klasifikasi**: pencocokan persis dengan contoh adalah prioritas. Transaksi ambigu ditandai untuk verifikasi.
- **Privasi**: tidak ada database atau penyimpanan permanen di aplikasi; berkas diproses di memori sesi server.
- **Lawan akun**: klasifikasi final tetap memerlukan pemeriksaan dokumen dan substansi transaksi.''')
