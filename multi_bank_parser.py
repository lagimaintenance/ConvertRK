"""Parser rekening koran PDF teks (BCA, BSI, CIMB, Mandiri, Muamalat, Permata).

Prinsip:
- Debit/credit di lembar bank adalah pengeluaran/pemasukan, sementara
  Debet/Kredit di GL bank = pemasukan/pengeluaran.
- Pertahankan semua transaksi; jangan deduplikasi berdasarkan nominal/referensi.
- Kegagalan atau perbedaan saldo harus eksplisit, tidak dianggap sukses.
- PDF gambar tanpa lapisan teks tidak didukung (tidak melakukan OCR).
"""
from __future__ import annotations

from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from io import BytesIO
from pathlib import Path
from collections import defaultdict
import re
import zipfile
import shutil
import subprocess

import pdfplumber

D0 = Decimal('0.00')
NUM = r'-?(?:\d{1,3}(?:,\d{3})+|\d+)\.\d{2}-?'
NUM_RE = re.compile(r'^' + NUM + r'$')
MON_ID = {'JANUARI':1,'FEBRUARI':2,'MARET':3,'APRIL':4,'MEI':5,'JUNI':6,
          'JULI':7,'AGUSTUS':8,'SEPTEMBER':9,'OKTOBER':10,'NOVEMBER':11,'DESEMBER':12}

class ParsingError(ValueError):
    pass


def amount(value, *, nullable=False):
    if value is None or not str(value).strip() or str(value).strip() in ('-', '—'):
        if nullable: return None
        return D0
    s = str(value).strip().replace(' ', '').replace('\u00a0','')
    negative = s.startswith('(') and s.endswith(')') or s.startswith('-') or s.endswith('-')
    s = s.strip('()-')
    # Format US: 1,234,567.89; format Indonesia: 1.234.567,89;
    # serta nominal bulat 1.234.567 atau 1234567 pada bank lain.
    if NUM_RE.fullmatch(s):
        normalized=s.replace(',','')
    elif re.fullmatch(r'\d{1,3}(?:\.\d{3})+,\d{2}',s):
        normalized=s.replace('.','').replace(',','.')
    elif re.fullmatch(r'\d+,\d{2}',s):
        normalized=s.replace(',','.')
    elif re.fullmatch(r'\d{1,3}(?:\.\d{3})+',s):
        normalized=s.replace('.','')
    elif re.fullmatch(r'\d{1,3}(?:,\d{3})+',s):
        normalized=s.replace(',','')
    elif re.fullmatch(r'\d+',s):
        normalized=s
    else:
        raise ParsingError(f'Nilai uang tidak dapat dikenali: {value!r}')
    try: num = Decimal(normalized)
    except InvalidOperation as e:raise ParsingError(f'Angka tidak valid: {value!r}') from e
    return -num if negative else num


def clean(text):
    return re.sub(r'\s+', ' ', str(text or '')).strip()


def date_fmt(s, *patterns):
    s=clean(s)
    for pattern in patterns:
        try:return datetime.strptime(s,pattern).date()
        except ValueError: pass
    raise ParsingError(f'Tanggal tidak dapat dibaca: {s!r}')


def tx(tanggal, uraian, debet=0, kredit=0, saldo_pdf=None, halaman=1, reference=''):
    return {'tanggal': tanggal, 'uraian':clean(uraian), 'debet':Decimal(debet),
            'kredit':Decimal(kredit), 'saldo_pdf':saldo_pdf, 'halaman':halaman,
            'referensi': clean(reference)}


def parse_bsi(pdf):
    records=[]
    for pg in pdf.pages:
        for table in pg.extract_tables():
            for row in table:
                if len(row)<8 or not str(row[0] or '').strip().isdigit() or not re.fullmatch(r'20\d{6}', str(row[2] or '').strip()):continue
                dk=clean(row[4]).upper()
                if dk not in ('D','K'):continue
                money=abs(amount(row[5])); descr=clean(row[7]); refer=clean(row[1]);
                records.append(tx(date_fmt(row[2],'%Y%m%d'),f'{descr} | Ref: {refer}',
                                  money if dk=='K' else D0,
                                  money if dk=='D' else D0,
                                  amount(row[6]), pg.page_number, refer))
    header=pdf.pages[0].extract_text() or ''
    found=re.search(r'Rekening\s*:\s*(?:IDR\s+)?(\d+)',header,re.I)
    op=re.search(r'Saldo Awal\s*:\s*('+NUM+r')',header,re.I)
    end=re.search(r'Saldo Akhir\s*:\s*('+NUM+r')',header,re.I)
    td=re.search(r'Total Debet\s*:\s*('+NUM+r')',header,re.I)
    tc=re.search(r'Total Kredit\s*:\s*('+NUM+r')',header,re.I)
    return records, {'rekening':found.group(1) if found else '',
                     'awal':amount(op.group(1)) if op else None,
                     'akhir':amount(end.group(1)) if end else None,
                     'kredit':amount(td.group(1)) if td else None,
                     'debet':amount(tc.group(1)) if tc else None,
                     'expected_count':None}


def parse_cimb(pdf):
    records=[]; numbers=[]
    for pg in pdf.pages:
        for tb in pg.extract_tables():
            for row in tb:
                if len(row)<15 or not clean(row[0]).isdigit() or not re.match(r'^\d\d/\d\d/\d\d',clean(row[1])):continue
                datestr=clean(row[1]).split()[0]
                ref=clean(row[14])
                debit=amount(row[9]);credit=amount(row[11]);bal=amount(row[12]);
                desc=clean(row[7])
                if ref:desc=f'{desc} | Ref: {ref}'
                records.append(tx(date_fmt(datestr,'%m/%d/%y'),desc,credit,debit,bal,pg.page_number,ref))
                numbers.append(int(clean(row[0])))
    records.reverse()  # PDF urut DESC; saldo pembukuan harus ascending
    header=pdf.pages[0].extract_text() or ''
    def get(k):
        m=re.search(k+r'\s*('+NUM+r')',header,re.I)
        return amount(m.group(1)) if m else None
    acct=re.search(r'Account Number\s+(\d+)',header,re.I)
    end=records[-1]['saldo_pdf'] if records else None
    db=get('Total Debit');cr=get('Total Credit')
    op=end - cr + db if end is not None and cr is not None and db is not None else None
    return records,{'rekening':acct.group(1) if acct else '', 'awal':op, 'akhir':end,
                    'kredit':db, 'debet':cr, 'expected_count':max(numbers) if numbers else None,
                    'sequential_ids':numbers}


def parse_mandiri(pdf):
    records=[]
    for pg in pdf.pages:
        for tb in pg.extract_tables():
            for row in tb:
                if len(row)<6 or not re.match(r'^\d{2}/\d{2}/20\d\d', clean(row[0])):continue
                dateval=clean(row[0])[:10];ref=clean(row[2])
                desc=clean(row[1]); db=amount(row[3]);cr=amount(row[4]);bal=amount(row[5]);
                if ref and ref!='-':desc+= ' | Ref: '+ref
                records.append(tx(date_fmt(dateval,'%d/%m/%Y'),desc,cr,db,bal,pg.page_number,ref))
    first=(pdf.pages[0].extract_text() or '')
    last=(pdf.pages[-1].extract_text() or '')
    no=re.search(r'Account No\s+(\d+)',first,re.I)
    op=re.search(r'Opening Balance\s*\n?\s*('+NUM+r')',first,re.I)
    db=re.search(r'Total Amount Debited\s*('+NUM+r')',last,re.I)
    cr=re.search(r'Total Amount Credited\s*('+NUM+r')',last,re.I)
    end=re.search(r'Closing Balance\s*('+NUM+r')',last,re.I)
    count_db=re.search(r'No of Debit\s*(\d+)',last,re.I)
    count_cr=re.search(r'No of Credit\s*(\d+)',last,re.I)
    return records,{'rekening':no.group(1) if no else '',
                    'awal':amount(op.group(1)) if op else None,
                    'akhir':amount(end.group(1)) if end else None,
                    'kredit':amount(db.group(1)) if db else None,
                    'debet':amount(cr.group(1)) if cr else None,
                    'expected_count':int(count_db.group(1))+int(count_cr.group(1)) if count_db and count_cr else None}


def parse_muamalat(pdf):
    records=[]
    # Keterangan mulai di kolom x >= 610; angka di x=335-550;
    # saldo sering bersambung dengan teks sehingga gunakan prefix regex.
    for pg in pdf.pages:
        words=pg.extract_words(x_tolerance=2,y_tolerance=2)
        lines=[]
        for w in sorted(words,key=lambda z:(round(z['top']),z['x0'])):
            y=w['top']
            if lines and abs(lines[-1]['y']-y)<2:
                lines[-1]['words'].append(w)
            else:lines.append({'y':y,'words':[w]})
        active=None
        for line in lines:
            ww=sorted(line['words'],key=lambda z:z['x0'])
            row=' '.join(w['text'] for w in ww)
            m=re.search(r'\b\d{2}-[A-Za-z]{3}-20\d\d\b',row)
            first_ref=[w['text'] for w in ww if 28<=w['x0']<130 and re.match(r'^[A-Za-z0-9]{7,}$',w['text'])]
            if m and first_ref and 105 <= line['y'] <= 560:
                datestr=m.group(0)
                dbtokens=[w['text'] for w in ww if 295<=w['x0']<420 and NUM_RE.fullmatch(w['text'])]
                crtokens=[w['text'] for w in ww if 420<=w['x0']<535 and NUM_RE.fullmatch(w['text'])]
                balw=[w['text'] for w in ww if 535<=w['x0']<610]
                if not balw:raise ParsingError(f'Muamalat halaman {pg.page_number}: saldo transaksi tidak ditemukan: {row[:100]}')
                bmatch=re.match(r'^('+NUM+r')',balw[0])
                if not bmatch:raise ParsingError(f'Muamalat halaman {pg.page_number}: saldo tidak valid {balw}')
                balance=amount(bmatch.group(1))
                d=amount(dbtokens[0]) if dbtokens else D0
                c=amount(crtokens[0]) if crtokens else D0
                if bool(d)==bool(c):raise ParsingError(f'Muamalat halaman {pg.page_number}: debit/kredit ambigu: {row[:130]}')
                descwords=[w['text'] for w in ww if w['x0']>=610]
                tail=balw[0][len(bmatch.group(1)):]
                if tail:descwords.insert(0,tail)
                ref=first_ref[0]
                active=tx(date_fmt(datestr,'%d-%b-%Y'), ' '.join(descwords), c, d, balance,pg.page_number,ref)
                active['uraian']=clean((active['uraian']+' | Ref: '+ref))
                records.append(active)
            elif active is not None and 105<=line['y']<=565:
                extra=' '.join(w['text'] for w in ww if w['x0']>=610)
                if extra and not re.search(r'copyright|reserved|all rights',extra,re.I):active['uraian']=clean(active['uraian'].replace(' | Ref: '+active['referensi'],'')+' '+extra+' | Ref: '+active['referensi'])
            elif active is None and records and 105<=line['y']<=565:
                # Lanjutan uraian baris halaman sebelumnya.
                extra=' '.join(w['text'] for w in ww if w['x0']>=610)
                if extra and not re.search(r'\b(?:reference|number|description|credit|debit|balance|copyright|reserved)\b',extra,re.I):
                    item=records[-1]
                    item['uraian']=clean(item['uraian'].replace(' | Ref: '+item['referensi'],'')+' '+extra+' | Ref: '+item['referensi'])
    header=pdf.pages[0].extract_text() or ''
    acct=re.search(r'Account No\.\s*:\s*(\d+)',header,re.I)
    # PDF Muamalat contoh tidak mencantumkan opening/closing pada header.
    # Opening boleh dihitung dari saldo transaksi PERTAMA (bukan asumsi 0).
    op=(records[0]['saldo_pdf']-records[0]['debet']+records[0]['kredit']) if records else None
    end=records[-1]['saldo_pdf'] if records else None
    return records,{'rekening':acct.group(1) if acct else '', 'awal':op,'akhir':end,
                    'opening_inferred':True,'closing_inferred':True}


PERMATA_NEW=re.compile(r'^(\d{1,4})(\d{2}-[A-Za-z]{3}-20\d\d)\s+(\d{2}-[A-Za-z]{3}-20\d\d)\s+(.*)$')
PERMATA_AMOUNTS=re.compile(r'\s+('+NUM+r')\s+('+NUM+r')\s*$')

def parse_permata(pdf):
    records=[];seq=[];opening=None;closing=None
    # Pastikan halaman berat diproses dalam batch kecil, supaya memori
    # Streamlit Community Cloud tidak membesar terus untuk PDF 60+ halaman.
    original_bytes = pdf.stream.getvalue()
    total_pages = len(pdf.pages)
    text_pages=None
    # Poppler (pdftotext) opsional dan sangat cepat untuk PDF Permata yang
    # menumpuk glyph dua kali. Tersedia lewat packages.txt di Streamlit Cloud.
    if shutil.which('pdftotext'):
        try:
            proc=subprocess.run(['pdftotext','-layout','-','-'],input=original_bytes,
                                stdout=subprocess.PIPE,stderr=subprocess.PIPE,
                                timeout=120,check=True)
            if proc.stdout.strip():
                text_pages=proc.stdout.decode('utf-8',errors='replace').split('\f')
        except (OSError,subprocess.SubprocessError):
            text_pages=None
    if text_pages is None:
        # Cadangan pada Mac tanpa Poppler: pdfplumber, membaca tiap 8 halaman
        # agar pemakaian RAM stabil. Halaman dengan glyph bertumpuk dibersihkan.
        text_pages=[]
        for start in range(0,total_pages,8):
            with pdfplumber.open(BytesIO(original_bytes)) as portion:
                for pg in portion.pages[start:start+8]:
                    text_pages.append(pg.dedupe_chars(tolerance=1).extract_text(x_tolerance=2,y_tolerance=3) or '')
                    pg.close()
    for ix,text in enumerate(text_pages):
        if not text.strip():continue
        active=None;before_new=[];passed_header=False
        for line in text.splitlines():
            lm=PERMATA_NEW.match(line.strip())
            if lm:
                if before_new and records:
                    records[-1]['uraian']=clean(records[-1]['uraian']+' '+clean(' '.join(before_new)))
                before_new=[]
                n=int(lm.group(1));dt=date_fmt(lm.group(2),'%d-%b-%Y')
                rest=lm.group(4)
                nums=PERMATA_AMOUNTS.search(rest)
                if not nums:raise ParsingError(f'Permata halaman {ix+1}, transaksi {n}: debit/kredit tidak ditemukan.')
                debit=amount(nums.group(1));credit=amount(nums.group(2))
                description=clean(rest[:nums.start()])
                description=re.sub(r'^\d{3,6}\s+\d{8,}\s*','',description)
                active=tx(dt,description,credit,debit,None,ix+1,str(n))
                records.append(active);seq.append(n)
                continue
            o=re.search(r'Opening Ledger Balance per\s+\d{2}-[A-Za-z]{3}-20\d\d\s+('+NUM+r')',line,re.I)
            if o:opening=amount(o.group(1))
            c=re.search(r'Closing Ledger Balance per\s+\d{2}-[A-Za-z]{3}-20\d\d\s+('+NUM+r')',line,re.I)
            if c:closing=amount(c.group(1))
            st=line.strip()
            if st.lower().startswith('code ref no'):passed_header=True;continue
            if st and not re.match(r'^(?:Inquiry Time|Account\s*:|Group\s*:|Period\s*:|Total(?:\s|$)|Transactio|Transaction|Opening |Closing |Ineffective |Hold Amount|Loan Facility|Effective Balance|Description:|No\.\s+Post|Record not found|Ref No|n Code|#\s*=|@\s*=|\d{1,4}\s+Post Date)',st,re.I):
                if active is not None:active['uraian']=clean(active['uraian']+' '+st)
                elif passed_header and not re.search(r'(?:^\d+\s+Post Date|^\w+ Code|^Ref No)',st,re.I):before_new.append(st)
    head=pdf.pages[0].extract_text() or ''
    m=re.search(r'Account\s*:\s*(\d+)',head,re.I)
    return records,{'rekening':m.group(1) if m else '', 'awal':opening,'akhir':closing,
                    'expected_count': max(seq) if seq else None,'sequential_ids':seq}


def parse_bca(pdf, name):
    """BCA format asli Januari: date DD/MM, DB flag, saldo halaman berikut."""
    def rows(pg):
        lines=[]; norm=595.0/float(pg.width)
        for w in sorted(pg.extract_words(x_tolerance=2,y_tolerance=2), key=lambda w:(w['top'],w['x0'])):
            y=float(w['top']);entry={'x':float(w['x0'])*norm,'text':w['text']}
            if lines and abs(lines[-1]['y']-y)<2:lines[-1]['words'].append(entry)
            else:lines.append({'y':y,'words':[entry]})
        for l in lines:
            l['words'].sort(key=lambda w:w['x']);l['text']=' '.join(w['text'] for w in l['words'])
        return lines
    records=[];opening=None;footer={}
    full=' '.join(pg.extract_text() or '' for pg in pdf.pages)
    match=re.search(r'\b('+'|'.join(MON_ID)+r')\s+(20\d\d)\b',full,re.I)
    if match:month,year=MON_ID[match.group(1).upper()],int(match.group(2))
    else:
        m=re.search(r'(20\d{2})(0[1-9]|1[0-2])',name)
        if not m:raise ParsingError('BCA: bulan dan tahun tidak dapat dikenali.')
        year,month=int(m.group(1)),int(m.group(2))
    acc=re.search(r'NO\.?\s*REKENING\s*:\s*([\d\s]{8,25})',full,re.I)
    for pg in pdf.pages:
        ll=rows(pg); end=float('inf');start=0
        def money(l,low,high):
            possible=[w['text'] for w in l['words'] if low<=w['x']<high and NUM_RE.fullmatch(w['text'])]
            return amount(possible[0]) if possible else None
        for l in ll:
            u=l['text'].upper()
            if 'TANGGAL' in u and 'KETERANGAN' in u and 'MUTASI' in u:start=l['y']+2
            for tag,key in [('MUTASI CR','debet'),('MUTASI DB','kredit'),('SALDO AKHIR','akhir'),('SALDO AWAL :','awal')]:
                if tag in u:
                    val=money(l,280,415)
                    if val is not None:footer[key]=val
                    end=min(end,l['y']-1)
        matches=[]
        for ix,l in enumerate(ll):
            if start<l['y']<end:
                a=[w['text'] for w in l['words'] if 15<=w['x']<82 and re.fullmatch(r'\d\d/\d\d',w['text'])]
                if a:matches.append((ix,a[0]))
        for j,(ix,dm) in enumerate(matches):
            endix=matches[j+1][0] if j+1<len(matches) else len(ll)
            body=[l for l in ll[ix:endix] if l['y']<end]
            if not body:continue
            first=body[0];parts=[]
            for l in body:
                a=' '.join(w['text'] for w in l['words'] if 83<=w['x']<190)
                b=' '.join(w['text'] for w in l['words'] if 190<=w['x']<302)
                if a:parts.append(a)
                if b:parts.append(b)
            desc=clean(' '.join(parts))
            bal=money(first,480,595)
            if desc.upper().startswith('SALDO AWAL'):
                if bal is not None:opening=bal
                continue
            moneyval=money(first,350,460)
            if moneyval is None:raise ParsingError(f'BCA halaman {pg.page_number} tanggal {dm}: mutasi tidak ditemukan')
            is_db=any(w['text'].upper()=='DB' for w in first['words'] if 430<=w['x']<480)
            day,mon=map(int,dm.split('/'))
            records.append(tx(date(year,mon,day),desc,D0 if is_db else moneyval,moneyval if is_db else D0,bal,pg.page_number))
    if opening is None:opening=footer.get('awal')
    return records,{'rekening':re.sub(r'\D','',acc.group(1)) if acc else '', 'awal':opening,
                    'akhir':footer.get('akhir'), 'debet':footer.get('debet'),'kredit':footer.get('kredit')}


def parse_generic(pdf):
    """Fallback untuk bank lain dengan tabel dan kolom tanggal/debet/kredit/saldo.
    Tidak menebak posisi bila judul kolom tidak terbaca.
    """
    records=[]; acc=''; detected_header=False
    headers={'DATE':['tanggal','transaction date','posting date','post date','tgl','date'],
             'DEBIT':['debit','debet','db'], 'CREDIT':['credit','kredit','cr'],
             'BALANCE':['balance','saldo','running balance'],
             'DESCRIPTION':['keterangan','description','remark','narration','uraian','details']}
    def heading(cell):return clean(cell).lower()
    for pg in pdf.pages:
        for table in pg.extract_tables():
            map_cols=None
            for r in table:
                if map_cols is None:
                    mapped={}
                    for i,col in enumerate(r):
                        st=heading(col)
                        for key,names in headers.items():
                            if st in names and key not in mapped:mapped[key]=i
                    if all(k in mapped for k in ['DATE','DEBIT','CREDIT','BALANCE']):
                        map_cols=mapped;detected_header=True
                    continue
                if not map_cols:continue
                try:
                    dcell=clean(r[map_cols['DATE']]);dstr=re.search(r'\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{2}-[A-Za-z]{3}-\d{4}|20\d{6}',dcell)
                    if not dstr:continue
                    d=date_fmt(dstr.group(),'%d/%m/%Y','%d-%m-%Y','%d-%b-%Y','%Y%m%d','%d/%m/%y')
                    db=amount(r[map_cols['DEBIT']]);cr=amount(r[map_cols['CREDIT']]);bal=amount(r[map_cols['BALANCE']])
                    if not db and not cr:continue
                    desc=clean(r[map_cols['DESCRIPTION']]) if 'DESCRIPTION' in map_cols else 'Transaksi rekening koran'
                    records.append(tx(d,desc,cr,db,bal,pg.page_number))
                except (ParsingError,IndexError,TypeError):continue
    if not records:raise ParsingError('Format bank belum dikenali. PDF harus berbasis teks atau memiliki tabel tanggal, debit, kredit, dan saldo.')
    # generic: infer opening from first row. Validate every available running balance.
    op=records[0]['saldo_pdf']-records[0]['debet']+records[0]['kredit']
    return records,{'rekening':acc,'awal':op,'akhir':records[-1]['saldo_pdf'], 'opening_inferred':True,'closing_inferred':True}


def identify(text):
    """Kenali BANK PENERBIT dari header, bukan nama bank lawan transaksi."""
    s=text.upper()[:5000]
    if 'TRANSACTION INQUIRY' in s and ('PERMATA' in s or 'OPENING LEDGER BALANCE' in s):return 'Permata'
    if 'LAPORAN REKENING KORAN' in s and 'ACCOUNT STATEMENT REPORT' in s:return 'Mandiri'
    if 'BSI TABUNGAN' in s and ('TRXID' in s or 'D/K' in s):return 'BSI'
    if 'CURRENT ACCOUNT' in s and 'ACCOUNT NUMBER' in s and ('BIZCHANNEL' in s or 'CIMB' in s):return 'CIMB Niaga'
    if 'GIRO IB HIJRAH' in s and 'ACCOUNT NO.' in s:return 'Muamalat'
    if 'REKENING GIRO' in s and 'MUTASI' in s:return 'BCA'
    return 'Umum (tabel)'


def parse_bank_pdf(content,name):
    if not content.lstrip().startswith(b'%PDF'):raise ParsingError('Bukan file PDF valid.')
    try:
        with pdfplumber.open(BytesIO(content)) as pdf:
            if not pdf.pages:raise ParsingError('PDF kosong.')
            head=pdf.pages[0].extract_text() or ''
            if len(head.strip())<25:raise ParsingError('PDF scan/gambar tanpa teks tidak didukung; gunakan PDF digital atau OCR terlebih dahulu.')
            bank=identify(head)
            methods={'BSI':parse_bsi,'CIMB Niaga':parse_cimb,'Mandiri':parse_mandiri,'Muamalat':parse_muamalat,'Permata':parse_permata}
            if bank=='BCA':records,meta=parse_bca(pdf,name)
            else:records,meta=methods.get(bank,parse_generic)(pdf)
    except ParsingError:raise
    except Exception as err:raise ParsingError(f'Gagal membaca PDF: {type(err).__name__}: {err}') from err
    if not records:raise ParsingError(f'Tidak ada mutasi yang terdeteksi (format {bank}).')
    if meta.get('awal') is None:raise ParsingError('Saldo awal tidak ditemukan dan tidak dapat dihitung.')
    opening=meta['awal']; running=opening; problems=[]
    for i,t in enumerate(records,1):
        running=running+t['debet']-t['kredit'];t['saldo']=running;t['urutan']=i
        if t['saldo_pdf'] is not None and running!=t['saldo_pdf']:
            if sum('saldo per transaksi' in x for x in problems)<4:
                problems.append(f'saldo per transaksi #{i} ({t["tanggal"]:%d/%m/%Y}): PDF {t["saldo_pdf"]:,.2f}, hitung {running:,.2f}')
    total_in=sum((t['debet'] for t in records),D0)
    total_out=sum((t['kredit'] for t in records),D0)
    for k,label,val in [('debet','masuk',total_in),('kredit','keluar',total_out)]:
        if meta.get(k) is not None and meta[k]!=val:
            problems.append(f'total {label} PDF {meta[k]:,.2f} tidak cocok dengan hasil {val:,.2f}')
    if meta.get('akhir') is not None and meta['akhir']!=running:
        problems.append(f'saldo akhir PDF {meta["akhir"]:,.2f} tidak cocok dengan hasil {running:,.2f}')
    if meta.get('expected_count') is not None and len(records)!=meta['expected_count']:
        problems.append(f'jumlah transaksi {len(records)} tidak sesuai dengan jumlah PDF {meta["expected_count"]}')
    seq=meta.get('sequential_ids')
    if seq and sorted(seq)!=list(range(1,max(seq)+1)):
        problems.append('urutan nomor transaksi pada PDF tidak lengkap atau ganda')
    # Untuk bank yang tidak mencantumkan saldo awal/akhir, keberhasilan
    # rekonsiliasi internal bukan verifikasi terhadap ringkasan PDF.
    if meta.get('opening_inferred'):problems.append('saldo awal dihitung dari transaksi pertama (tidak tercantum eksplisit pada PDF)')
    if meta.get('closing_inferred'):problems.append('saldo akhir memakai saldo transaksi terakhir (tidak tercantum ringkasan terpisah)')
    if not meta.get('rekening'):problems.append('nomor rekening tidak terdeteksi')
    first=records[0]['tanggal'];months={(t['tanggal'].year,t['tanggal'].month) for t in records}
    return {'bank':bank,'file':Path(name).name,'rekening':meta.get('rekening') or 'TIDAK TERBACA',
            'year':first.year,'month':first.month,'saldo_awal':opening,'saldo_akhir':running,
            'debet':total_in,'kredit':total_out,'transactions':records,'warnings':problems,
            'months':months,'verified':not problems}


def expand_uploads(files,max_pdf=100,max_bytes=90*1024*1024):
    contents=[]
    for file in files:
        data=file.getvalue()
        if file.name.lower().endswith('.pdf'):contents.append((Path(file.name).name,data))
        elif file.name.lower().endswith('.zip'):
            try:
                with zipfile.ZipFile(BytesIO(data)) as z:
                    for i in z.infolist():
                        if i.is_dir() or not i.filename.lower().endswith('.pdf'):continue
                        if i.file_size>25*1024*1024:raise ParsingError(f'PDF {i.filename} melebihi 25 MB.')
                        contents.append((Path(i.filename).name,z.read(i)))
            except zipfile.BadZipFile as e:raise ParsingError(f'ZIP rusak: {file.name}') from e
        if len(contents)>max_pdf:raise ParsingError('PDF terlalu banyak (maksimum 100).')
        if sum(len(x[1]) for x in contents)>max_bytes:raise ParsingError('Total file terlalu besar (maksimum 90 MB).')
    return contents
