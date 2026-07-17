from __future__ import annotations

import io, math, re, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
DEFAULT_FILE = Path(__file__).with_name("مقارنة مشتريات جدة.xlsx")

st.set_page_config(page_title="تحليل أسعار المشتريات V3", page_icon="📊", layout="wide")
st.markdown("""
<style>
html,body,[class*=css]{direction:rtl;text-align:right}.main .block-container{padding-top:1rem;max-width:1500px}
.hero{background:linear-gradient(135deg,#17324d,#0e7490);color:#fff;border-radius:20px;padding:22px 28px;margin-bottom:14px}.hero h1{margin:0;font-size:31px}.hero p{margin:8px 0 0;opacity:.92}
div[data-testid=stMetric]{background:#fff;border:1px solid #e5e7eb;border-radius:15px;padding:14px;box-shadow:0 2px 12px rgba(0,0,0,.05)}
.note{background:#f8fafc;border-right:5px solid #0e7490;padding:12px 14px;border-radius:8px}.decision{padding:14px;border-radius:12px;background:#f8fafc;border:1px solid #e2e8f0;margin:8px 0}
</style>""", unsafe_allow_html=True)


def col_index(ref:str)->int:
    letters=re.match(r"[A-Z]+",ref).group(0); n=0
    for ch in letters:n=n*26+ord(ch)-64
    return n


def read_xlsx_cached(source:bytes)->Tuple[pd.DataFrame,List[str]]:
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        shared=[]
        if "xl/sharedStrings.xml" in z.namelist():
            root=ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si",NS):shared.append("".join((t.text or "") for t in si.findall(".//a:t",NS)))
        sheet=ET.fromstring(z.read("xl/worksheets/sheet1.xml")); rows={}
        for row in sheet.findall(".//a:sheetData/a:row",NS):
            r=int(row.attrib["r"]); rows[r]={}
            for c in row.findall("a:c",NS):
                idx=col_index(c.attrib["r"]); typ=c.attrib.get("t"); v=c.find("a:v",NS); val=None
                if v is not None and v.text is not None:
                    raw=v.text
                    if typ=="s": val=shared[int(raw)]
                    elif typ=="b": val=raw=="1"
                    else:
                        try: val=float(raw)
                        except: val=raw
                rows[r][idx]=val
    header=rows.get(6,{}); base={2:"كود الصنف",3:"اسم الصنف",4:"التصنيف",5:"الوحدة",6:"المورد",7:"السعر الأساسي"}
    month_cols=[]
    for c in range(9,21):
        h=header.get(c)
        if h is not None and str(h).strip(): month_cols.append((c,str(int(h)) if isinstance(h,float) and h.is_integer() else str(h)))
    month_cols=month_cols[:5]
    recs=[]
    for r in sorted(k for k in rows if k>6):
        vals=rows[r]
        if vals.get(2) in (None,"") or vals.get(3) in (None,""):continue
        rec={label:vals.get(idx) for idx,label in base.items()}
        for idx,m in month_cols:
            try:x=float(vals.get(idx))
            except:x=math.nan
            rec[f"شهر {m}"]=x if pd.notna(x) and x>0 else math.nan
        recs.append(rec)
    df=pd.DataFrame(recs); months=[f"شهر {m}" for _,m in month_cols]
    return df,months


def enrich(df:pd.DataFrame,months:List[str])->pd.DataFrame:
    o=df.copy()
    for m in months:o[m]=pd.to_numeric(o[m],errors="coerce")
    o["أول سعر"]=o[months].bfill(axis=1).iloc[:,0]; o["آخر سعر"]=o[months].ffill(axis=1).iloc[:,-1]
    o["التغير الكلي %"]=(o["آخر سعر"]/o["أول سعر"]-1)*100
    o.loc[o["أول سعر"].isna()|o["آخر سعر"].isna(),"التغير الكلي %"]=math.nan
    o["متوسط السعر"]=o[months].mean(axis=1); o["أعلى سعر"]=o[months].max(axis=1); o["أقل سعر"]=o[months].min(axis=1)
    o["التذبذب %"]=o[months].std(axis=1)/o["متوسط السعر"]*100; o["عدد الشهور المتاحة"]=o[months].count(axis=1)
    o["الاتجاه"]=o["التغير الكلي %"].apply(lambda x:"مرتفع" if pd.notna(x) and x>.5 else("منخفض" if pd.notna(x) and x<-.5 else "مستقر"))
    o["درجة المخاطر"]=(o["التغير الكلي %"].abs().fillna(0)*2+o["التذبذب %"].fillna(0)*3).clip(0,100)
    o["مستوى المخاطر"]=o["درجة المخاطر"].apply(lambda x:"مرتفع" if x>=60 else("متوسط" if x>=30 else "منخفض"))
    for i in range(1,len(months)):
        p,c=months[i-1],months[i]; o[f"تغير {p} ← {c} %"]=(o[c]/o[p]-1)*100
    o["التوصية"]=o.apply(decision,axis=1)
    return o


def decision(r):
    ch=r.get("التغير الكلي %",0); risk=r.get("درجة المخاطر",0)
    if pd.isna(ch):return "بيانات غير كافية"
    if ch>=10 or risk>=60:return "تفاوض مع المورد / ابحث عن بديل"
    if ch>=3:return "راقب السعر وحاول تثبيت سعر تعاقدي"
    if ch<=-5:return "فرصة شراء أو تثبيت السعر الحالي"
    return "استمرار الشراء مع المتابعة"


def grouped(df,key,months):
    rows=[]
    for name,g in df.groupby(key,dropna=False):
        rows.append({key:name if pd.notna(name) else "غير محدد","عدد الأصناف":len(g),"متوسط التغير %":g["التغير الكلي %"].mean(),"مرتفعة":int((g.الاتجاه=="مرتفع").sum()),"منخفضة":int((g.الاتجاه=="منخفض").sum()),"متوسط المخاطر":g["درجة المخاطر"].mean(),**{m:g[m].mean() for m in months}})
    return pd.DataFrame(rows).sort_values("متوسط التغير %",ascending=False)


def pct(x):return "—" if pd.isna(x) else f"{x:,.2f}%"
def money(x):return "—" if pd.isna(x) else f"{x:,.2f} ر.س"


def make_excel(items,cats,sups,months)->bytes:
    wb=Workbook(); ws=wb.active; ws.title="Dashboard"; ws.sheet_view.rightToLeft=True
    navy="17324D"; teal="0E7490"; red="FECACA"; green="DCFCE7"; yellow="FEF3C7"; white="FFFFFF"
    ws.merge_cells("A1:H2"); ws["A1"]="تقرير تحليل تغير أسعار المشتريات"; ws["A1"].font=Font(size=20,bold=True,color=white); ws["A1"].fill=PatternFill("solid",fgColor=teal); ws["A1"].alignment=Alignment(horizontal="center",vertical="center")
    kpis=[("إجمالي الأصناف",len(items)),("التصنيفات",items['التصنيف'].nunique()),("الموردون",items['المورد'].nunique()),("متوسط التغير",items['التغير الكلي %'].mean()/100),("مرتفعة",int((items.الاتجاه=='مرتفع').sum())),("منخفضة",int((items.الاتجاه=='منخفض').sum()))]
    for i,(label,val) in enumerate(kpis):
        c=1+i; ws.cell(4,c,label); ws.cell(5,c,val); ws.cell(4,c).fill=PatternFill("solid",fgColor=navy); ws.cell(4,c).font=Font(color=white,bold=True); ws.cell(4,c).alignment=Alignment(horizontal="center"); ws.cell(5,c).alignment=Alignment(horizontal="center"); ws.cell(5,c).font=Font(size=15,bold=True)
    ws["D5"].number_format="0.00%"
    top=items.nlargest(10,"التغير الكلي %")[["اسم الصنف","التغير الكلي %"]]; ws.append([]); ws.append(["أعلى الأصناف ارتفاعًا","نسبة التغير"])
    start=ws.max_row
    for row in top.itertuples(index=False):ws.append([row[0],row[1]/100])
    ws[f"B{start+1}:B{ws.max_row}"][0][0].number_format="0.00%"
    for r in range(start+1,ws.max_row+1):ws.cell(r,2).number_format="0.00%"
    ch=BarChart(); ch.title="أعلى 10 أصناف ارتفاعًا"; ch.add_data(Reference(ws,min_col=2,min_row=start,max_row=ws.max_row),titles_from_data=True); ch.set_categories(Reference(ws,min_col=1,min_row=start+1,max_row=ws.max_row)); ch.height=8; ch.width=15; ws.add_chart(ch,"D8")

    def add_sheet(name,df):
        sh=wb.create_sheet(name); sh.sheet_view.rightToLeft=True; cols=list(df.columns); sh.append(cols)
        for row in df.itertuples(index=False,name=None): sh.append([None if pd.isna(v) else v for v in row])
        for cell in sh[1]: cell.fill=PatternFill("solid",fgColor=navy); cell.font=Font(color=white,bold=True); cell.alignment=Alignment(horizontal="center")
        sh.freeze_panes="A2"; sh.auto_filter.ref=sh.dimensions
        thin=Side(style="thin",color="D1D5DB")
        for row in sh.iter_rows():
            for cell in row: cell.border=Border(bottom=thin); cell.alignment=Alignment(vertical="center",wrap_text=True)
        for i,col in enumerate(cols,1):
            width=min(max(len(str(col))+2,12),28); sh.column_dimensions[get_column_letter(i)].width=width
            if "%" in str(col):
                for r in range(2,sh.max_row+1):
                    if isinstance(sh.cell(r,i).value,(int,float)): sh.cell(r,i).value=sh.cell(r,i).value/100; sh.cell(r,i).number_format="0.00%"
        for m in months:
            if m in cols:
                ci=cols.index(m)+1
                if ci>1:
                    prev=ci-1; rng=f"{get_column_letter(ci)}2:{get_column_letter(ci)}{sh.max_row}"
                    sh.conditional_formatting.add(rng,CellIsRule(operator="greaterThan",formula=[f"{get_column_letter(prev)}2"],fill=PatternFill("solid",fgColor=red),font=Font(color="9B1C1C",bold=True)))
                    sh.conditional_formatting.add(rng,CellIsRule(operator="lessThan",formula=[f"{get_column_letter(prev)}2"],fill=PatternFill("solid",fgColor=green),font=Font(color="166534")))
        return sh

    display=["كود الصنف","اسم الصنف","التصنيف","الوحدة","المورد"]+months+["التغير الكلي %","التذبذب %","الاتجاه","درجة المخاطر","مستوى المخاطر","التوصية"]
    add_sheet("البيانات",items[display]); add_sheet("تحليل التصنيفات",cats); add_sheet("تحليل الموردين",sups)
    rec=items[["اسم الصنف","التصنيف","المورد","التغير الكلي %","درجة المخاطر","التوصية"]].sort_values("درجة المخاطر",ascending=False); add_sheet("التوصيات",rec)
    out=io.BytesIO(); wb.save(out); return out.getvalue()


st.markdown('<div class="hero"><h1>📊 برنامج تحليل تغير أسعار المشتريات V3</h1><p>تحليل الصنف والتصنيف والمورد، مركز اتخاذ القرار، وتصدير Excel احترافي.</p></div>',unsafe_allow_html=True)
with st.sidebar:
    st.header("إعدادات البيانات"); uploaded=st.file_uploader("اختر ملف Excel",type=["xlsx"]); st.caption("يتم تحليل أول خمسة أشهر من الصف 6، والأصفار تعامل كبيانات مفقودة.")
try:
    raw=uploaded.getvalue() if uploaded else DEFAULT_FILE.read_bytes(); base,months=read_xlsx_cached(raw); data=enrich(base,months)
except Exception as e: st.error(f"تعذر قراءة الملف: {e}"); st.stop()
with st.sidebar:
    cats_all=sorted(data["التصنيف"].dropna().astype(str).unique()); sups_all=sorted(data["المورد"].dropna().astype(str).unique())
    sc=st.multiselect("التصنيفات",cats_all,default=cats_all); ss=st.multiselect("الموردون",sups_all,default=sups_all); rf=st.multiselect("المخاطر",["مرتفع","متوسط","منخفض"],default=["مرتفع","متوسط","منخفض"])
filtered=data[data["التصنيف"].astype(str).isin(sc)&data["المورد"].astype(str).isin(ss)&data["مستوى المخاطر"].isin(rf)].copy()
for _col in ["اسم الصنف", "التصنيف", "المورد", "الوحدة"]:
    if _col in filtered.columns:
        filtered[_col] = filtered[_col].fillna("غير محدد").astype(str).str.strip()
if filtered.empty:st.warning("لا توجد بيانات تطابق الفلاتر");st.stop()
cat_sum=grouped(filtered,"التصنيف",months); sup_sum=grouped(filtered,"المورد",months)
cols=st.columns(7); vals=[("الأصناف",len(filtered)),("التصنيفات",filtered['التصنيف'].nunique()),("الموردون",filtered['المورد'].nunique()),("مرتفعة",(filtered.الاتجاه=='مرتفع').sum()),("منخفضة",(filtered.الاتجاه=='منخفض').sum()),("متوسط التغير",pct(filtered['التغير الكلي %'].mean())),("عالية المخاطر",(filtered['مستوى المخاطر']=='مرتفع').sum())]
for c,(l,v) in zip(cols,vals):c.metric(l,v)

tabs=st.tabs(["لوحة المدير","تحليل الصنف","تحليل التصنيف","تحليل المورد","مركز القرار","البيانات والتصدير"])
with tabs[0]:
    a,b=st.columns(2)
    with a: st.plotly_chart(px.bar(filtered.nlargest(10,'التغير الكلي %').sort_values('التغير الكلي %'),x='التغير الكلي %',y='اسم الصنف',orientation='h',title='أعلى 10 أصناف ارتفاعًا',text_auto='.2f'),use_container_width=True)
    with b: st.plotly_chart(px.bar(filtered.nsmallest(10,'التغير الكلي %').sort_values('التغير الكلي %',ascending=False),x='التغير الكلي %',y='اسم الصنف',orientation='h',title='أعلى 10 أصناف انخفاضًا',text_auto='.2f'),use_container_width=True)
    a,b=st.columns(2)
    with a:
        tr=filtered[months].mean().reset_index();tr.columns=['الشهر','متوسط السعر'];st.plotly_chart(px.line(tr,x='الشهر',y='متوسط السعر',markers=True,title='اتجاه متوسط الأسعار'),use_container_width=True)
    with b: st.plotly_chart(px.pie(filtered['الاتجاه'].value_counts().reset_index(),names='الاتجاه',values='count',hole=.5,title='توزيع الاتجاهات'),use_container_width=True)
with tabs[1]:
    item_names = sorted(filtered['اسم الصنف'].dropna().astype(str).unique(), key=lambda x: x.casefold())
    name=st.selectbox("اختر الصنف", item_names);r=filtered[filtered['اسم الصنف'].astype(str)==str(name)].iloc[0]
    cs=st.columns(6)
    for c,(l,v) in zip(cs,[("أول سعر",money(r['أول سعر'])),("آخر سعر",money(r['آخر سعر'])),("التغير",pct(r['التغير الكلي %'])),("المتوسط",money(r['متوسط السعر'])),("التذبذب",pct(r['التذبذب %'])),("المخاطر",f"{r['درجة المخاطر']:.1f}/100")]):c.metric(l,v)
    p=pd.DataFrame({'الشهر':months,'السعر':[r[m] for m in months]});p['التغير الشهري %']=p['السعر'].pct_change()*100
    a,b=st.columns(2)
    with a:st.plotly_chart(px.line(p,x='الشهر',y='السعر',markers=True,title=f'حركة سعر {name}'),use_container_width=True)
    with b:st.plotly_chart(px.bar(p,x='الشهر',y='التغير الشهري %',title='التغير الشهري %',text_auto='.2f'),use_container_width=True)
    st.info(f"التوصية: {r['التوصية']}")
with tabs[2]:
    cat=st.selectbox("اختر التصنيف",sorted(filtered['التصنيف'].dropna().astype(str).unique()));g=filtered[filtered['التصنيف'].astype(str)==cat]
    a,b=st.columns(2)
    with a:
        tr=g[months].mean().reset_index();tr.columns=['الشهر','متوسط السعر'];st.plotly_chart(px.line(tr,x='الشهر',y='متوسط السعر',markers=True,title=f'اتجاه {cat}'),use_container_width=True)
    with b:st.plotly_chart(px.bar(g.sort_values('التغير الكلي %'),x='التغير الكلي %',y='اسم الصنف',orientation='h',title='تغير الأصناف',text_auto='.2f'),use_container_width=True)
with tabs[3]:
    supplier=st.selectbox("اختر المورد",sorted(filtered['المورد'].dropna().astype(str).unique()));g=filtered[filtered['المورد'].astype(str)==supplier]
    cs=st.columns(5)
    for c,(l,v) in zip(cs,[("عدد الأصناف",len(g)),("متوسط التغير",pct(g['التغير الكلي %'].mean())),("مرتفعة",(g.الاتجاه=='مرتفع').sum()),("منخفضة",(g.الاتجاه=='منخفض').sum()),("متوسط المخاطر",f"{g['درجة المخاطر'].mean():.1f}")]):c.metric(l,v)
    st.plotly_chart(px.bar(g.sort_values('التغير الكلي %'),x='التغير الكلي %',y='اسم الصنف',orientation='h',title=f'أداء المورد: {supplier}',text_auto='.2f'),use_container_width=True)
with tabs[4]:
    urgent=filtered[filtered['التوصية'].str.contains('تفاوض|بديل',regex=True)].sort_values('درجة المخاطر',ascending=False)
    st.subheader("أصناف تحتاج تدخلًا عاجلًا")
    st.dataframe(urgent[["اسم الصنف","التصنيف","المورد","التغير الكلي %","درجة المخاطر","التوصية"]],use_container_width=True,hide_index=True)
    st.subheader("محاكاة خصم تفاوضي")
    discount=st.slider("نسبة الخصم المقترحة %",0,20,5); qty=st.number_input("الكمية الافتراضية لكل صنف",min_value=1,value=100)
    impact=((urgent['آخر سعر']*discount/100)*qty).sum(); st.metric("الوفر التقديري",money(impact))
with tabs[5]:
    display=["كود الصنف","اسم الصنف","التصنيف","الوحدة","المورد"]+months+["التغير الكلي %","التذبذب %","الاتجاه","درجة المخاطر","مستوى المخاطر","التوصية"]
    st.dataframe(filtered[display],use_container_width=True,height=520,hide_index=True)
    xlsx=make_excel(filtered,cat_sum,sup_sum,months)
    st.download_button("📥 تحميل تقرير Excel الاحترافي",xlsx,"تقرير_تحليل_الأسعار_V3.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.markdown('<div class="note"><b>منهجية الحساب:</b> نسبة التغير = (آخر سعر متاح ÷ أول سعر متاح − 1) × 100، مع استبعاد الأسعار الصفرية.</div>',unsafe_allow_html=True)
