import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from copy import copy
import xlrd, datetime, io, zipfile, base64, os
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.colors import white, black

st.set_page_config(page_title="Hotel Guest Processor", page_icon="🛎️", layout="centered")

# ── Load embedded templates ──────────────────────────────────────────────
@st.cache_resource
def load_template(name):
    path = os.path.join(os.path.dirname(__file__), f'tmpl_{name}.b64')
    with open(path, 'r') as f:
        return base64.b64decode(f.read())

# ── Lookup tables ─────────────────────────────────────────────────────────
NAT_KBTT = {
    'Liên bang Nga':'RUS - Russia','Ka-dắc-xtan':'KAZ - Kazakhstan',
    'CH Hàn Quốc':'KOR - Korea (South)','Mỹ':'USA - United States of America',
    'U-dơ-bê-ki-xtan':'UZB - Uzbekistan','Kiếc-ghi-di-a':'KGZ - Kyrgyzstan',
    'Ta-gi-ki-xtan':'TJK - Tajikistan','Ca-na-da':'CAN - Canada',
    'Anh':'GBR - United Kingdom','Đức':'DEU - Germany',
    'CH Liên bang Đức':'DEU - Germany',
    'Trung Quốc':'CHN - China','Ô-xtrây-li-a':'AUS - Australia',
    'Bê-la-rút':'BLR - Belarus','U-crai-na':'UKR - Ukraine',
    'Môn-đô-va':'MDA - Moldova','Phần Lan':'FIN - Finland','Pháp':'FRA - France',
    'Đan Mạch':'DNK - Denmark','Mô-ri-xơ':'MUS - Mauritius',
}
NAT_DK14 = {
    'RUS':'Russia  (Liên bang Nga)','UZB':'Uzbekistan  ( U-dơ-bê-ki-xtan )',
    'KAZ':'Kazakhstan  ( Ka-dắc-xtan )','KOR':'Korea (South)  ( CH Hàn Quốc )',
    'KGZ':'Kyrgyzstan  ( Kiếc-ghi-di-a )','TJK':'Tajikistan  ( Ta-gi-ki-xtan )',
    'UKR':'Ukraine  ( U-crai-na )','USA':'United States  ( Mỹ )',
    'VNM':'Vietnam  ( Việt Nam )','CAN':'Canada  ( Ca-na-da )',
    'GBR':'United Kingdom  ( Anh )','AUS':'Australia  ( Ô-xtrây-li-a )',
    'BLR':'Belarus  ( Bê-la-rút )','CHN':'China  ( Trung Quốc )',
    'DEU':'Germany  ( Đức )','MDA':'Moldova  ( Môn-đô-va )',
    'FIN':'Finland  ( Phần Lan )','FRA':'France  ( Pháp )',
    'DNK':'Denmark  ( Đan Mạch )','MUS':'Mauritius  ( Mô-ri-xơ )',
}
LOAI_GIAY = {
    'Căn cước công dân':'8 - Thẻ Căn Cước','Hộ chiếu':'4 - Hộ chiếu',
    'Chứng minh nhân dân':'2 - Thẻ CMND','Căn cước':'1 - Thẻ CCCD',
    'Giấy khai sinh':'5 - Giấy khai sinh',
}
TINH = {
    'DAK LAK':'605 - Đắk Lắk','DAK NONG':'607 - Đắk Nông','PHU YEN':'509 - Phú Yên',
    'HO CHI MINH':'701 - TP. Hồ Chí Minh','THP HO CHI MINH':'701 - TP. Hồ Chí Minh',
    'HCM':'701 - TP. Hồ Chí Minh','TP HCM':'701 - TP. Hồ Chí Minh',
    'GIA LAI':'603 - Gia Lai','QUANG NGAI':'505 - Quảng Ngãi','LAM DONG':'703 - Lâm Đồng',
    'LAM DONG.':'703 - Lâm Đồng','KHANH HOA':'511 - Khánh Hòa','BEN TRE':'817 - Bến Tre',
    'VINH LONG':'809 - Vĩnh Long','NINH THUAN':'513 - Ninh Thuận','DONG NAI':'713 - Đồng Nai',
    'TIEN GIANG':'807 - Tiền Giang','LONG AN':'801 - Long An','BINH DUONG':'707 - Bình Dương',
    'BINH THUAN':'705 - Bình Thuận','CAN THO':'815 - TP. Cần Thơ','DA NANG':'501 - TP. Đà Nẵng',
    'HA NOI':'101 - TP. Hà Nội','BA RIA VUNG TAU':'711 - Bà Rịa - Vũng Tàu',
    'DONG THAP':'803 - Đồng Tháp','AN GIANG':'805 - An Giang','KIEN GIANG':'819 - Kiên Giang',
    'HA TINH':'407 - Hà Tĩnh','BINH DINH':'507 - Bình Định','VIET NAM':'',
}

# ── Helpers ───────────────────────────────────────────────────────────────
def fmt(v):
    if v is None or str(v) in ('NaT','nan',''): return ''
    if hasattr(v,'strftime'): return v.strftime('%d/%m/%Y')
    return str(v).strip()[:10]

def make_code(prefix, ns):
    if not ns: return prefix
    p = ns.replace('-','/').split('/')
    return f"{prefix}{p[0].zfill(2)}{p[1].zfill(2)}{p[2][-2:]}" if len(p)==3 else prefix

def cp(src, dst):
    for a in ('font','fill','border','alignment'):
        v = getattr(src,a)
        if v: setattr(dst,a,copy(v))
    dst.number_format = src.number_format

def serial2date(s):
    if not s: return None
    try: return datetime.datetime(1899,12,30)+datetime.timedelta(days=int(s))
    except: return None

def wb_to_bytes(wb):
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()

# ── Processing ────────────────────────────────────────────────────────────
def process_xlsx(xlsx_bytes, rate):
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active
    don_gia_col = next((c.column for c in ws[1] if c.value=='ĐƠN GIÁ'), None)
    conv = 0
    if don_gia_col:
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            cell = row[don_gia_col-1]
            if cell.value and isinstance(cell.value,(int,float)) and 0 < cell.value < 1000:
                cell.value = round(cell.value * rate)
                cell.fill = PatternFill("solid", start_color="FFFF00")
                conv += 1
    return wb, conv

def split_wb(wb, loai):
    wb2 = load_workbook(io.BytesIO(wb_to_bytes(wb)))
    ws2 = wb2.active
    lc = next(c.column for c in ws2[1] if c.value=='LOẠI KHÁCH')
    dels = [row[0].row for row in ws2.iter_rows(min_row=2,max_row=ws2.max_row)
            if row[lc-1].value != loai]
    for r in reversed(dels): ws2.delete_rows(r)
    for i, row in enumerate(ws2.iter_rows(min_row=2,max_row=ws2.max_row),1): row[0].value=i
    return wb2

def build_kbtt(df_intl):
    wb = load_workbook(io.BytesIO(load_template('kbtt')))
    ws = wb['KBTT']
    ref = [ws.cell(4,c) for c in range(1,12)]
    for r in range(ws.max_row,3,-1): ws.delete_rows(r)
    for i,(_,row) in enumerate(df_intl.iterrows(),1):
        er=i+3
        ht=str(row.get('HỌ TÊN ',row.get('HỌ TÊN',''))).strip()
        ns=fmt(row['NGÀY SINH']); nd=fmt(row['NGÀY ĐẾN']); ni=fmt(row.get('NGÀY ÐI',row.get('NGÀY ĐI','')))
        gt='M - Nam' if str(row.get('GIỚI TÍNH','')).strip()=='Nam' else 'F - Nữ'
        qt=NAT_KBTT.get(str(row.get('QUỐC TỊCH','')).strip(), str(row.get('QUỐC TỊCH','')))
        sh=str(row.get('SỐ GIẤY TỜ','')).strip(); sp=str(row.get('SỐ PHÒNG','')).strip()
        vals=[i,ht,ns,'D - Ngày',gt,qt,sh,sp,nd,ni,ni]
        for ci,val in enumerate(vals,1):
            cell=ws.cell(er,ci); cell.value=val if isinstance(val,int) else str(val)
            cp(ref[ci-1],cell)
    return wb

def build_vnm(df_vn):
    wb = load_workbook(io.BytesIO(load_template('vnm')))
    wsn = next((s for s in wb.sheetnames if 'KHACH' in s or 'DS' in s), wb.sheetnames[0])
    ws = wb[wsn]
    ref = [ws.cell(5,c) for c in range(1,ws.max_column+1)]
    for r in range(ws.max_row,4,-1): ws.delete_rows(r)
    gks_cnt=0; gbl_cnt=0
    for i,(_,row) in enumerate(df_vn.iterrows(),1):
        er=i+4
        ht=str(row.get('HỌ TÊN ',row.get('HỌ TÊN',''))).strip()
        ns=fmt(row['NGÀY SINH']); nd=fmt(row['NGÀY ĐẾN']); ni=fmt(row.get('NGÀY ÐI',row.get('NGÀY ĐI','')))
        gt='F - Nữ' if str(row.get('GIỚI TÍNH','')).strip()=='Nữ' else 'M - Nam'
        sg_raw=str(row.get('SỐ GIẤY TỜ','')).strip()
        lg_raw=str(row.get('LOẠI GIẤY TỜ','')).strip()
        is_gks='GKS' in sg_raw.upper(); is_gbl='GBL' in sg_raw.upper()
        ten_giay=''
        if is_gks:
            sg=make_code('GKS',ns); lg='5 - Giấy khai sinh'; gks_cnt+=1
        elif is_gbl:
            sg=make_code('GBL',ns); lg='Giấy tờ khác'; ten_giay='Giấy bảo lãnh'; gbl_cnt+=1
        else:
            sg=sg_raw; lg=LOAI_GIAY.get(lg_raw,lg_raw)
        tinh=TINH.get(str(row.get('TP/TỈNH','')).strip().upper(),'')
        dc=str(row.get('ÐỊA CHỈ',row.get('ĐỊA CHỈ',''))).strip()
        sp=str(row.get('SỐ PHÒNG','')).strip()
        vals=[i,ht,ns,gt,'VNM - Viet Nam',lg,ten_giay,sg,'','1 - Thường trú',tinh,'',dc,nd,ni,sp,'1 - Du lịch','','']
        for ci,val in enumerate(vals,1):
            cell=ws.cell(er,ci); cell.value=val if isinstance(val,int) else str(val)
            if ci<=len(ref): cp(ref[ci-1],cell)
    return wb, gks_cnt, gbl_cnt

def build_dk14(xls_bytes):
    wb2=xlrd.open_workbook(file_contents=xls_bytes)
    ws2=wb2.sheet_by_index(0)
    data=[[ws2.cell_value(r,c) for c in range(ws2.ncols)]
          for r in range(1,ws2.nrows) if any(ws2.cell_value(r,c) for c in range(ws2.ncols))]
    wb_t=load_workbook(io.BytesIO(load_template('dk14')))
    ws_t=wb_t.active
    cw={col:ws_t.column_dimensions[col].width for col in ws_t.column_dimensions}
    rh={r:ws_t.row_dimensions[r].height for r in ws_t.row_dimensions if r<=17}
    wb_o=Workbook(); ws_o=wb_o.active
    for col,w in cw.items(): ws_o.column_dimensions[col].width=w
    for r,h in rh.items():
        if h: ws_o.row_dimensions[r].height=h
    def cc(s,d):
        d.value=s.value
        for a in ('font','fill','border','alignment'):
            v=getattr(s,a)
            if v: setattr(d,a,copy(v))
        d.number_format=s.number_format
    for r in range(1,18):
        for c in range(1,14): cc(ws_t.cell(r,c),ws_o.cell(r,c))
    for mc in ws_t.merged_cells.ranges:
        if mc.min_row<=17:
            ws_o.merge_cells(start_row=mc.min_row,start_column=mc.min_col,
                             end_row=mc.max_row,end_column=mc.max_col)
    wb_t.close()
    thin=Side(style='thin'); bdr=Border(left=thin,right=thin,top=thin,bottom=thin)
    fn=Font(name='Times New Roman',size=12)
    ac=Alignment(horizontal='center',vertical='center',wrap_text=True)
    al=Alignment(horizontal='left',vertical='center',wrap_text=True)
    for i,row in enumerate(data,1):
        er=i+17
        name=str(row[1]).strip() if row[1] else ''
        gender=str(row[3]).strip().upper() if row[3] else ''
        country=NAT_DK14.get(str(row[4]).strip().upper() if row[4] else '',str(row[4] or ''))
        passport=str(row[5]).strip() if row[5] else ''
        if passport.endswith('.0'): passport=passport[:-2]
        address=str(row[6]).strip() if row[6] else '   '
        dob=serial2date(row[2]); arr=serial2date(row[7]); dep=serial2date(row[8])
        room=str(row[9]).strip() if row[9] else ''
        notify=str(row[10]).strip() if row[10] else ''
        cols=[(1,i,ac),(2,name,al),(3,dob if gender=='M' else None,ac),
              (4,dob if gender=='F' else None,ac),(5,country,ac),(6,passport,ac),
              (7,address,al),(8,arr,ac),(9,dep,ac),(10,room,ac),(11,notify,al),(12,'',ac),(13,'',ac)]
        for ci,val,aln in cols:
            cell=ws_o.cell(er,ci); cell.value=val; cell.font=fn; cell.border=bdr; cell.alignment=aln
            if ci in (3,4) and val: cell.number_format='DD/MM/YYYY'
            elif ci in (8,9) and val: cell.number_format='DD/MM/YYYY'
    return wb_o, len(data)


# ── Regcard PDF builder ───────────────────────────────────────────────────
def load_regcard_template():
    path = os.path.join(os.path.dirname(__file__), 'tmpl_regcard.b64')
    with open(path, 'r') as f:
        return base64.b64decode(f.read())

def _rc_clean_name(n):
    if pd.isna(n): return ''
    return str(n).strip().rstrip(',').strip()

def _rc_conf(c):
    if pd.isna(c): return ''
    return str(int(c)) if isinstance(c,(int,float)) else str(c)

def _rc_date(d):
    if pd.isna(d): return ''
    if hasattr(d,'strftime'):
        return f"{d.month:02d}/{d.day:02d}/{d.year}"
    s=str(d).strip(); p=s.split('/')
    if len(p)==3:
        dd,mm,yy=p
        if len(yy)==2: yy='20'+yy
        return f"{dd}/{mm}/{yy}"
    return s

def _rc_nights(arr,dep):
    try:
        if hasattr(arr,'strftime'):
            a=pd.Timestamp(year=arr.year,month=arr.day,day=arr.month)
        else:
            p=str(arr).split('/'); a=pd.Timestamp(f"20{p[2]}-{p[1]}-{p[0]}")
        p=str(dep).split('/')
        d=pd.Timestamp(f"20{p[2]}-{p[1]}-{p[0]}") if len(p[2])==2 else pd.to_datetime(dep)
        return str((d-a).days)
    except: return ''

def build_regcards(xlsx_bytes, only_main=True):
    """Tạo PDF regcard hàng loạt. only_main=True: chỉ khách chính (có Conf#)."""
    df = pd.read_excel(io.BytesIO(xlsx_bytes))
    H = 841.0
    FONT = "Times-Roman"; SIZE = 9.8
    POS = {
        'name':(125.5,109.6),'conf':(526.8,108.5),'arrival':(119.9,144.2),
        'departure':(329.9,144.2),'nights':(500.0,144.2),'type':(113.6,179.9),
        'rm':(360.6,179.9),'company':(221.4,216.3),
    }
    BLANK = [
        (124,99,245,111),(525,97,565,110),(118,133,167,146),(328,133,377,146),
        (498,133,510,146),(112,169,134,181),(359,169,383,181),(219,205,266,218),
        (135,277,540,290),(142,444,285,457),
    ]
    tmpl_bytes = load_regcard_template()
    writer = PdfWriter()
    count = 0
    for _, row in df.iterrows():
        # Skip non-main rows if only_main
        if only_main and pd.isna(row.get('Conf#')):
            continue
        name = _rc_clean_name(row.get('Name'))
        if not name:
            continue
        data = {
            'name': name,
            'conf': _rc_conf(row.get('Conf#')),
            'arrival': _rc_date(row.get('Arrival')),
            'departure': _rc_date(row.get('Departure')),
            'nights': _rc_nights(row.get('Arrival'), row.get('Departure')),
            'type': str(row.get('Type')) if pd.notna(row.get('Type')) else '',
            'rm': str(row.get('Rm')) if pd.notna(row.get('Rm')) else '',
            'company': str(row.get('Company')) if pd.notna(row.get('Company')) else '',
        }
        buf = io.BytesIO()
        c = rl_canvas.Canvas(buf, pagesize=(595,841))
        c.setFillColor(white)
        for x0,top,x1,bot in BLANK:
            c.rect(x0-1, H-bot-1, (x1-x0)+2, (bot-top)+2, fill=1, stroke=0)
        c.setFillColor(black)
        c.setFont(FONT, SIZE)
        for key,(x,bottom) in POS.items():
            if data[key]:
                c.drawString(x, H-bottom, data[key])
        c.save(); buf.seek(0)
        base = PdfReader(io.BytesIO(tmpl_bytes))
        overlay = PdfReader(buf)
        page = base.pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
        count += 1
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue(), count


# ── UI ────────────────────────────────────────────────────────────────────

# Custom CSS
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .block-container {padding-top: 2.5rem; padding-bottom: 2rem; max-width: 820px;}
    .app-header {
        position: relative; overflow: hidden;
        background: linear-gradient(135deg, #0f3460, #16213e, #1a1a2e, #0f3460);
        background-size: 300% 300%;
        animation: gradientShift 12s ease infinite;
        border-radius: 18px; padding: 1.75rem 2rem; margin-bottom: 1.75rem;
        box-shadow: 0 8px 28px rgba(15,52,96,0.35);
    }
    @keyframes gradientShift {
        0% {background-position: 0% 50%;}
        50% {background-position: 100% 50%;}
        100% {background-position: 0% 50%;}
    }
    .header-glow {
        position: absolute; top: -60%; right: -20%;
        width: 300px; height: 300px; border-radius: 50%;
        background: radial-gradient(circle, rgba(96,165,250,0.25) 0%, transparent 70%);
        animation: floatGlow 8s ease-in-out infinite;
    }
    @keyframes floatGlow {
        0%,100% {transform: translate(0,0) scale(1);}
        50% {transform: translate(-30px,20px) scale(1.2);}
    }
    .header-content {position: relative; z-index: 1;}
    .welcome-line {
        display: flex; align-items: center; gap: 8px; margin-bottom: 8px;
    }
    .wave {
        display: inline-block; font-size: 1.3rem;
        animation: wave 2.2s ease-in-out infinite;
        transform-origin: 70% 70%;
    }
    @keyframes wave {
        0%,60%,100% {transform: rotate(0deg);}
        10% {transform: rotate(14deg);}
        20% {transform: rotate(-8deg);}
        30% {transform: rotate(14deg);}
        40% {transform: rotate(-4deg);}
        50% {transform: rotate(10deg);}
    }
    .typing {
        color: #7dd3fc; font-size: 1.05rem; font-weight: 600;
        overflow: hidden; white-space: nowrap;
        border-right: 2px solid #7dd3fc;
        width: 0; animation: typing 2s steps(24, end) forwards, blink 0.8s step-end infinite;
    }
    @keyframes typing {from {width: 0;} to {width: 100%;}}
    @keyframes blink {50% {border-color: transparent;}}
    .app-header h1 {
        color: #ffffff; font-size: 1.7rem; font-weight: 700;
        margin: 0; display: flex; align-items: center; gap: 12px; letter-spacing: -0.5px;
        animation: slideUp 0.6s ease 0.3s both;
    }
    @keyframes slideUp {from {opacity: 0; transform: translateY(12px);} to {opacity: 1; transform: translateY(0);}}
    .app-header p {
        color: rgba(255,255,255,0.7); margin: 6px 0 0 0; font-size: 0.9rem;
        animation: slideUp 0.6s ease 0.5s both;
    }
    .app-logo {animation: bellRing 3s ease-in-out infinite;}
    @keyframes bellRing {
        0%,50%,100% {transform: rotate(0deg);}
        5% {transform: rotate(12deg);} 10% {transform: rotate(-10deg);}
        15% {transform: rotate(8deg);} 20% {transform: rotate(-6deg);}
        25% {transform: rotate(0deg);}
    }
    .clock-badge {
        display: inline-block; margin-top: 12px;
        background: rgba(255,255,255,0.12); backdrop-filter: blur(6px);
        color: rgba(255,255,255,0.9); font-size: 0.82rem; font-weight: 500;
        padding: 6px 14px; border-radius: 20px;
        border: 1px solid rgba(255,255,255,0.15);
        animation: slideUp 0.6s ease 0.7s both;
    }
    #live-clock {font-variant-numeric: tabular-nums; font-weight: 700; color: #7dd3fc;}
    .menu-card {
        border-radius: 16px; padding: 1.5rem 1.4rem; margin-bottom: 0.75rem;
        border: 1.5px solid #e2e8f0; background: #ffffff;
        transition: all 0.25s ease; cursor: default; min-height: 168px;
        position: relative; overflow: hidden;
    }
    .menu-card:hover {
        transform: translateY(-4px); box-shadow: 0 12px 28px rgba(15,52,96,0.15);
        border-color: #0f3460;
    }
    .menu-card::before {
        content: ""; position: absolute; top: 0; left: 0; right: 0; height: 4px;
    }
    .menu-blue::before {background: linear-gradient(90deg, #0f3460, #3b82f6);}
    .menu-green::before {background: linear-gradient(90deg, #16a34a, #4ade80);}
    .menu-icon {
        font-size: 2.4rem; margin-bottom: 0.6rem; display: inline-block;
        animation: menuFloat 3s ease-in-out infinite;
    }
    @keyframes menuFloat {0%,100% {transform: translateY(0);} 50% {transform: translateY(-6px);}}
    .menu-title {font-size: 1.1rem; font-weight: 700; color: #0f3460; margin-bottom: 0.4rem;}
    .menu-desc {font-size: 0.82rem; color: #64748b; line-height: 1.5;}
    .app-logo {
        width: 46px; height: 46px; background: rgba(255,255,255,0.12);
        border-radius: 12px; display: inline-flex; align-items: center;
        justify-content: center; font-size: 1.5rem;
    }
    .section-label {
        font-size: 0.72rem; font-weight: 700; color: #64748b;
        text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 0.5rem;
    }
    div[data-testid="stFileUploader"] {
        border: 1.5px dashed #cbd5e1; border-radius: 12px;
        padding: 0.5rem; background: #f8fafc;
    }
    .stButton button {border-radius: 10px; font-weight: 600; padding: 0.7rem;}
    .stDownloadButton button {border-radius: 10px; font-weight: 700; padding: 0.85rem;}
    div[data-testid="stMetric"] {
        background: #f8fafc; border: 1px solid #e2e8f0;
        border-radius: 12px; padding: 0.9rem 0.5rem; text-align: center;
    }
    div[data-testid="stMetricValue"] {color: #0f3460; font-weight: 700;}
</style>
""", unsafe_allow_html=True)

# Animated welcome header
_now = datetime.datetime.now()
_hour = _now.hour
if _hour < 11:
    _greet, _emoji = "Chào buổi sáng", "🌅"
elif _hour < 14:
    _greet, _emoji = "Chào buổi trưa", "☀️"
elif _hour < 18:
    _greet, _emoji = "Chào buổi chiều", "🌤️"
else:
    _greet, _emoji = "Chào buổi tối", "🌙"

_weekdays = ["Thứ Hai","Thứ Ba","Thứ Tư","Thứ Năm","Thứ Sáu","Thứ Bảy","Chủ Nhật"]
_wd = _weekdays[_now.weekday()]
_datestr = f"{_wd}, {_now.day:02d}/{_now.month:02d}/{_now.year}"

st.markdown(f"""
<div class="app-header">
    <div class="header-glow"></div>
    <div class="header-content">
        <div class="welcome-line">
            <span class="wave">{_emoji}</span>
            <span class="typing">{_greet}, Ta Châu!</span>
        </div>
        <h1><span class="app-logo">🛎️</span> Hotel Guest Processor</h1>
        <p>Xử lý dữ liệu khách sạn tự động — KBTT · Thông báo lưu trú · ĐK14 · Regcard</p>
        <div class="clock-badge">📅 {_datestr}&nbsp;&nbsp;·&nbsp;&nbsp;<span id="live-clock">{_now.strftime('%H:%M:%S')}</span></div>
    </div>
</div>
<script>
(function(){{
    function tick(){{
        var el = window.parent.document.getElementById('live-clock');
        if(el){{
            var d = new Date();
            el.textContent = String(d.getHours()).padStart(2,'0')+':'+String(d.getMinutes()).padStart(2,'0')+':'+String(d.getSeconds()).padStart(2,'0');
        }}
    }}
    setInterval(tick, 1000);
}})();
</script>
""", unsafe_allow_html=True)

# Menu selection (session state)
if "menu" not in st.session_state:
    st.session_state.menu = None

def go_menu(name):
    st.session_state.menu = name

# ── Menu landing screen ───────────────────────────────────────────────────
if st.session_state.menu is None:
    st.markdown('<div class="section-label">✨ Chọn chức năng</div>', unsafe_allow_html=True)
    st.write("")
    mcol1, mcol2 = st.columns(2)
    with mcol1:
        st.markdown("""
        <div class="menu-card menu-blue">
            <div class="menu-icon">📋</div>
            <div class="menu-title">Xử lý hồ sơ hàng ngày</div>
            <div class="menu-desc">Quy đổi tỷ giá · Tách Quốc tế/Việt Nam · KBTT · Thông báo lưu trú · ĐK14</div>
        </div>
        """, unsafe_allow_html=True)
        st.button("Mở chức năng này  →", key="btn_daily", use_container_width=True,
                  on_click=go_menu, args=("daily",))
    with mcol2:
        st.markdown("""
        <div class="menu-card menu-green">
            <div class="menu-icon">🖨️</div>
            <div class="menu-title">Tạo Regcard PDF</div>
            <div class="menu-desc">Điền dữ liệu booking lên mẫu Registration Card · Xuất PDF hàng loạt</div>
        </div>
        """, unsafe_allow_html=True)
        st.button("Mở chức năng này  →", key="btn_regcard", use_container_width=True,
                  on_click=go_menu, args=("regcard",))

    st.divider()
    st.caption("🔒 File mẫu KBTT · Thông báo lưu trú VNM · ĐK14 · Regcard đã tích hợp sẵn — xử lý an toàn")

# ── Daily processing screen ───────────────────────────────────────────────
if st.session_state.menu == "daily":
    st.button("←  Quay lại menu", key="back_daily", on_click=go_menu, args=(None,))
    st.write("")
    st.markdown('<div class="section-label">⚙️ Cài đặt</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        rate = st.number_input("💱 Tỷ giá USD/EUR → VNĐ", value=29535.15, step=0.01, format="%.2f")
    with col2:
        today = datetime.date.today()
        date_str = st.text_input("📅 Ngày (dùng cho tên file)", value=f"{today.day}_{today.month:02d}")

    st.write("")
    st.markdown('<div class="section-label">📂 Tải file lên</div>', unsafe_allow_html=True)
    col_x, col_s = st.columns(2)
    with col_x:
        xlsx_file = st.file_uploader("File XLSX — Dữ liệu khách (bắt buộc)", type=['xlsx'], key="daily_xlsx")
    with col_s:
        xls_file = st.file_uploader("File XLS — Nguồn ĐK14 (tùy chọn)", type=['xls'], key="daily_xls")

    st.write("")

    if st.button("⚡ Bắt đầu xử lý", type="primary", disabled=xlsx_file is None, use_container_width=True):
        with st.spinner("Đang xử lý..."):
            try:
                xlsx_bytes = xlsx_file.read()
                progress = st.progress(0, text="Quy đổi tỷ giá...")

                wb, conv = process_xlsx(xlsx_bytes, rate)
                df = pd.read_excel(io.BytesIO(xlsx_bytes))
                df_intl = df[df['LOẠI KHÁCH']=='Quốc tế'].reset_index(drop=True)
                df_vn   = df[df['LOẠI KHÁCH']=='Việt Nam'].reset_index(drop=True)

                progress.progress(15, text="Tách file Quốc tế / Việt Nam...")
                wb_intl = split_wb(wb, 'Quốc tế')
                wb_vn   = split_wb(wb, 'Việt Nam')

                progress.progress(35, text="Điền mẫu KBTT...")
                wb_kbtt = build_kbtt(df_intl)

                progress.progress(55, text="Điền mẫu Thông báo lưu trú VNM...")
                wb_vnm, gks_cnt, gbl_cnt = build_vnm(df_vn)

                progress.progress(75, text="Đóng gói ZIP...")

                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    zf.writestr(f'converted_{date_str}.xlsx',      wb_to_bytes(wb))
                    zf.writestr(f'KhachQuocTe_{date_str}.xlsx',    wb_to_bytes(wb_intl))
                    zf.writestr(f'KhachVietNam_{date_str}.xlsx',   wb_to_bytes(wb_vn))
                    zf.writestr(f'ho_so_KBTT_NNN_{date_str}.xlsx', wb_to_bytes(wb_kbtt))
                    zf.writestr(f'thong_bao_luu_tru_VNM_{date_str}.xlsx', wb_to_bytes(wb_vnm))
                    has_dk14 = False
                    if xls_file:
                        progress.progress(85, text="Điền mẫu ĐK14...")
                        xls_bytes = xls_file.read()
                        wb_dk14, dk_count = build_dk14(xls_bytes)
                        zf.writestr(f'dk14_{date_str}.xlsx', wb_to_bytes(wb_dk14))
                        has_dk14 = True

                progress.progress(100, text="Hoàn tất!")
                progress.empty()

                st.success("✅ Xử lý hoàn tất!")

                c1,c2,c3,c4 = st.columns(4)
                c1.metric("Tổng khách", len(df))
                c2.metric("Quốc tế", len(df_intl))
                c3.metric("Việt Nam", len(df_vn))
                c4.metric("GKS + GBL", f"{gks_cnt} + {gbl_cnt}")

                st.info(f"💱 Đã quy đổi tỷ giá cho **{conv}** ô (đã tô vàng)")

                files_made = ["📄 converted (file chung)", "🌍 KhachQuocTe", "🇻🇳 KhachVietNam",
                              "📝 KBTT NNN", "📑 Thông báo lưu trú VNM"]
                if has_dk14:
                    files_made.append("🚔 ĐK14")
                st.markdown("**File đã tạo:** " + " · ".join(files_made))

                st.download_button(
                    label="⬇️ Tải về tất cả file (ZIP)",
                    data=zip_buf.getvalue(),
                    file_name=f"hotel_{date_str}.zip",
                    mime="application/zip",
                    use_container_width=True,
                    type="primary"
                )
            except Exception as e:
                st.error(f"❌ Lỗi: {e}")
                st.exception(e)

# ── Regcard screen ────────────────────────────────────────────────────────
if st.session_state.menu == "regcard":
    st.button("←  Quay lại menu", key="back_regcard", on_click=go_menu, args=(None,))
    st.write("")
    st.markdown('<div class="section-label">🖨️ Tạo Registration Card hàng loạt</div>', unsafe_allow_html=True)
    st.caption("Điền dữ liệu từ file Excel (Booking list) lên mẫu Regcard PDF gốc — giữ nguyên 100% form.")

    rc_file = st.file_uploader("File Excel dữ liệu booking (.xlsx)", type=['xlsx'], key="rc_xlsx")

    only_main = st.checkbox("Chỉ tạo cho khách chính (có mã Conf#)", value=True,
                            help="Bỏ chọn để tạo regcard cho tất cả khách, kể cả khách đi cùng phòng")

    st.write("")

    if st.button("🖨️ Tạo Regcard PDF", type="primary", disabled=rc_file is None, use_container_width=True):
        with st.spinner("Đang tạo PDF..."):
            try:
                rc_bytes = rc_file.read()
                pdf_data, count = build_regcards(rc_bytes, only_main=only_main)

                if count == 0:
                    st.warning("⚠️ Không tìm thấy khách nào để tạo regcard. Kiểm tra lại file.")
                else:
                    st.success(f"✅ Đã tạo {count} regcard!")
                    st.metric("Số regcard", count)
                    st.download_button(
                        label=f"⬇️ Tải về {count} Regcard (PDF)",
                        data=pdf_data,
                        file_name=f"regcards_{datetime.date.today().strftime('%d_%m')}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                        type="primary"
                    )
            except Exception as e:
                st.error(f"❌ Lỗi: {e}")
                st.exception(e)


