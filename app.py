from __future__ import annotations

import io
import math
import os
import re
import statistics
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
DEFAULT_FILE = Path(__file__).with_name("مقارنة مشتريات جدة.xlsx")

st.set_page_config(page_title="تحليل مشتريات جدة", page_icon="📊", layout="wide")
st.markdown(
    """
    <style>
      html, body, [class*="css"] { direction: rtl; text-align: right; }
      .main .block-container { padding-top: 1.2rem; }
      div[data-testid="stMetric"] { background:#ffffff; border:1px solid #e5e7eb; border-radius:14px; padding:14px; box-shadow:0 2px 10px rgba(0,0,0,.04); }
      div[data-testid="stMetricLabel"] { font-weight:700; }
      .hero { background:linear-gradient(135deg,#17324d,#0e7490); color:#fff; border-radius:18px; padding:20px 24px; margin-bottom:14px; }
      .hero h1 { margin:0; font-size:31px; }
      .hero p { margin:7px 0 0; opacity:.9; }
      .note { background:#f8fafc; border-right:5px solid #0e7490; padding:12px 14px; border-radius:8px; }
    </style>
    """,
    unsafe_allow_html=True,
)


def col_index(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def read_xlsx_cached(source: bytes) -> Tuple[pd.DataFrame, List[str]]:
    """Read cached values from xlsx XML, including external-link formula results."""
    with zipfile.ZipFile(io.BytesIO(source)) as z:
        shared: List[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", NS):
                shared.append("".join((t.text or "") for t in si.findall(".//a:t", NS)))

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheet_el = wb.find(".//a:sheets/a:sheet", NS)
        sheet_name = sheet_el.attrib.get("name", "Sheet1") if sheet_el is not None else "Sheet1"
        sheet = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))

        rows: Dict[int, Dict[int, object]] = {}
        for row in sheet.findall(".//a:sheetData/a:row", NS):
            r = int(row.attrib["r"])
            rows[r] = {}
            for c in row.findall("a:c", NS):
                ref = c.attrib["r"]
                idx = col_index(ref)
                typ = c.attrib.get("t")
                v = c.find("a:v", NS)
                val: object = None
                if v is not None and v.text is not None:
                    raw = v.text
                    if typ == "s":
                        val = shared[int(raw)]
                    elif typ in ("str", "inlineStr"):
                        val = raw
                    elif typ == "b":
                        val = raw == "1"
                    else:
                        try:
                            val = float(raw)
                        except ValueError:
                            val = raw
                rows[r][idx] = val

    header_row = 6
    header = rows.get(header_row, {})
    base_cols = {2: "كود الصنف", 3: "اسم الصنف", 4: "التصنيف", 5: "الوحدة", 6: "المورد", 7: "السعر الأساسي"}
    month_cols: List[Tuple[int, str]] = []
    for c in range(9, 21):
        h = header.get(c)
        if h is not None and str(h).strip() != "":
            month_cols.append((c, str(int(h)) if isinstance(h, float) and h.is_integer() else str(h)))

    records = []
    for r in sorted(k for k in rows if k > header_row):
        vals = rows[r]
        name = vals.get(3)
        code = vals.get(2)
        if name in (None, "") or code in (None, ""):
            continue
        rec = {label: vals.get(idx) for idx, label in base_cols.items()}
        for idx, month in month_cols:
            x = vals.get(idx)
            try:
                x = float(x) if x is not None else math.nan
            except (TypeError, ValueError):
                x = math.nan
            # In this workbook zero represents no purchase / no valid price.
            rec[f"شهر {month}"] = x if x > 0 else math.nan
        records.append(rec)

    df = pd.DataFrame(records)
    for c in ["كود الصنف", "السعر الأساسي"]:
        if c in df:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    months = [f"شهر {m}" for _, m in month_cols]
    return df, months


def add_metrics(df: pd.DataFrame, months: List[str]) -> pd.DataFrame:
    out = df.copy()
    for m in months:
        out[m] = pd.to_numeric(out[m], errors="coerce")
    out["أول سعر"] = out[months].bfill(axis=1).iloc[:, 0]
    out["آخر سعر"] = out[months].ffill(axis=1).iloc[:, -1]
    out["التغير الكلي %"] = ((out["آخر سعر"] / out["أول سعر"]) - 1) * 100
    out.loc[(out["أول سعر"].isna()) | (out["آخر سعر"].isna()), "التغير الكلي %"] = math.nan
    out["متوسط السعر"] = out[months].mean(axis=1, skipna=True)
    out["أعلى سعر"] = out[months].max(axis=1, skipna=True)
    out["أقل سعر"] = out[months].min(axis=1, skipna=True)
    out["عدد الشهور المتاحة"] = out[months].count(axis=1)
    out["التذبذب %"] = out[months].std(axis=1, skipna=True) / out["متوسط السعر"] * 100
    out["الاتجاه"] = out["التغير الكلي %"].apply(lambda x: "مرتفع" if pd.notna(x) and x > .5 else ("منخفض" if pd.notna(x) and x < -.5 else "مستقر"))
    out["درجة المخاطر"] = (out["التغير الكلي %"].abs().fillna(0) * 2 + out["التذبذب %"].fillna(0) * 3).clip(0, 100)
    out["مستوى المخاطر"] = out["درجة المخاطر"].apply(lambda x: "مرتفع" if x >= 60 else ("متوسط" if x >= 30 else "منخفض"))
    for i in range(1, len(months)):
        prev, cur = months[i - 1], months[i]
        out[f"تغير {prev} ← {cur} %"] = (out[cur] / out[prev] - 1) * 100
    return out


def category_summary(df: pd.DataFrame, months: List[str]) -> pd.DataFrame:
    rows = []
    for cat, g in df.groupby("التصنيف", dropna=False):
        monthly = g[months].mean(skipna=True)
        valid = monthly.dropna()
        first = valid.iloc[0] if len(valid) else math.nan
        last = valid.iloc[-1] if len(valid) else math.nan
        change = (last / first - 1) * 100 if pd.notna(first) and first != 0 and pd.notna(last) else math.nan
        rows.append({
            "التصنيف": cat or "غير مصنف",
            "عدد الأصناف": len(g),
            "متوسط التغير %": g["التغير الكلي %"].mean(),
            "تغير متوسط سعر التصنيف %": change,
            "أصناف مرتفعة": int((g["الاتجاه"] == "مرتفع").sum()),
            "أصناف منخفضة": int((g["الاتجاه"] == "منخفض").sum()),
            "أصناف مستقرة": int((g["الاتجاه"] == "مستقر").sum()),
            "متوسط المخاطر": g["درجة المخاطر"].mean(),
            **{m: monthly.get(m, math.nan) for m in months},
        })
    return pd.DataFrame(rows).sort_values("متوسط التغير %", ascending=False)


def pct(x: float) -> str:
    return "—" if pd.isna(x) else f"{x:,.2f}%"


def money(x: float) -> str:
    return "—" if pd.isna(x) else f"{x:,.2f} ر.س"


st.markdown('<div class="hero"><h1>📊 برنامج تحليل تغير أسعار المشتريات</h1><p>تحليل احترافي للصنف والتصنيف خلال خمسة أشهر مع الرسوم البيانية ومؤشرات المخاطر.</p></div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("إعدادات البيانات")
    uploaded = st.file_uploader("اختر ملف Excel", type=["xlsx"])
    zero_note = st.checkbox("اعتبار السعر صفر كبيان مفقود", value=True, disabled=True)
    st.caption("البرنامج يقرأ القيم المخزنة داخل ملف Excel، حتى مع وجود روابط خارجية.")

try:
    raw = uploaded.getvalue() if uploaded else DEFAULT_FILE.read_bytes()
    base_df, months = read_xlsx_cached(raw)
    data = add_metrics(base_df, months)
except Exception as exc:
    st.error(f"تعذر قراءة الملف: {exc}")
    st.stop()

if len(months) < 2:
    st.error("لم يتم العثور على شهرين صالحين على الأقل في الصف 6.")
    st.stop()

with st.sidebar:
    categories = sorted(data["التصنيف"].dropna().astype(str).unique())
    suppliers = sorted(data["المورد"].dropna().astype(str).unique())
    selected_categories = st.multiselect("التصنيفات", categories, default=categories)
    selected_suppliers = st.multiselect("الموردون", suppliers, default=suppliers)
    risk_filter = st.multiselect("مستوى المخاطر", ["مرتفع", "متوسط", "منخفض"], default=["مرتفع", "متوسط", "منخفض"])

filtered = data[
    data["التصنيف"].astype(str).isin(selected_categories)
    & data["المورد"].astype(str).isin(selected_suppliers)
    & data["مستوى المخاطر"].isin(risk_filter)
].copy()

cats = category_summary(filtered, months) if not filtered.empty else pd.DataFrame()
valid_changes = filtered["التغير الكلي %"].dropna()

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("إجمالي الأصناف", f"{len(filtered):,}")
k2.metric("التصنيفات", f"{filtered['التصنيف'].nunique():,}")
k3.metric("مرتفعة", f"{(filtered['الاتجاه']=='مرتفع').sum():,}")
k4.metric("منخفضة", f"{(filtered['الاتجاه']=='منخفض').sum():,}")
k5.metric("متوسط التغير", pct(valid_changes.mean()))
k6.metric("عالية المخاطر", f"{(filtered['مستوى المخاطر']=='مرتفع').sum():,}")

if filtered.empty:
    st.warning("لا توجد بيانات تطابق الفلاتر المحددة.")
    st.stop()

summary_tab, item_tab, cat_tab, table_tab = st.tabs(["لوحة المدير", "تحليل الصنف", "تحليل التصنيف", "البيانات والتصدير"])

with summary_tab:
    c1, c2 = st.columns(2)
    top = filtered.dropna(subset=["التغير الكلي %"]).nlargest(10, "التغير الكلي %")
    bottom = filtered.dropna(subset=["التغير الكلي %"]).nsmallest(10, "التغير الكلي %")
    with c1:
        fig = px.bar(top.sort_values("التغير الكلي %"), x="التغير الكلي %", y="اسم الصنف", orientation="h", title="أعلى 10 أصناف ارتفاعًا", text_auto=".2f")
        fig.update_layout(height=430, yaxis_title="", xaxis_title="نسبة التغير %")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(bottom.sort_values("التغير الكلي %", ascending=False), x="التغير الكلي %", y="اسم الصنف", orientation="h", title="أعلى 10 أصناف انخفاضًا", text_auto=".2f")
        fig.update_layout(height=430, yaxis_title="", xaxis_title="نسبة التغير %")
        st.plotly_chart(fig, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        trend = filtered[months].mean().reset_index()
        trend.columns = ["الشهر", "متوسط السعر"]
        fig = px.line(trend, x="الشهر", y="متوسط السعر", markers=True, title="اتجاه متوسط أسعار جميع الأصناف")
        st.plotly_chart(fig, use_container_width=True)
    with c4:
        counts = filtered["الاتجاه"].value_counts().rename_axis("الاتجاه").reset_index(name="العدد")
        fig = px.pie(counts, names="الاتجاه", values="العدد", hole=.52, title="توزيع اتجاهات الأسعار")
        st.plotly_chart(fig, use_container_width=True)

    if not cats.empty:
        fig = px.bar(cats.sort_values("متوسط التغير %"), x="متوسط التغير %", y="التصنيف", orientation="h", title="متوسط نسبة تغير الأصناف حسب التصنيف", text_auto=".2f")
        fig.update_layout(height=max(420, 28 * len(cats)), yaxis_title="", xaxis_title="متوسط التغير %")
        st.plotly_chart(fig, use_container_width=True)

with item_tab:
    names = filtered.sort_values("اسم الصنف")["اسم الصنف"].tolist()
    selected_item = st.selectbox("اختر الصنف", names)
    item = filtered[filtered["اسم الصنف"] == selected_item].iloc[0]
    a, b, c, d, e, f = st.columns(6)
    a.metric("أول سعر", money(item["أول سعر"]))
    b.metric("آخر سعر", money(item["آخر سعر"]), pct(item["التغير الكلي %"]))
    c.metric("متوسط السعر", money(item["متوسط السعر"]))
    d.metric("أعلى سعر", money(item["أعلى سعر"]))
    e.metric("أقل سعر", money(item["أقل سعر"]))
    f.metric("درجة المخاطر", f"{item['درجة المخاطر']:.1f}/100")

    item_prices = pd.DataFrame({"الشهر": months, "السعر": [item[m] for m in months]})
    item_prices["التغير الشهري %"] = item_prices["السعر"].pct_change() * 100
    q1, q2 = st.columns(2)
    with q1:
        fig = px.line(item_prices, x="الشهر", y="السعر", markers=True, title=f"حركة سعر: {selected_item}")
        fig.update_traces(line_width=4, marker_size=10)
        st.plotly_chart(fig, use_container_width=True)
    with q2:
        fig = px.bar(item_prices.dropna(subset=["التغير الشهري %"]), x="الشهر", y="التغير الشهري %", title="نسبة التغير عن الشهر السابق", text_auto=".2f")
        st.plotly_chart(fig, use_container_width=True)

    details = pd.DataFrame({
        "البيان": ["الكود", "التصنيف", "الوحدة", "المورد", "الاتجاه", "مستوى المخاطر", "الشهور المتاحة", "التذبذب"],
        "القيمة": [item["كود الصنف"], item["التصنيف"], item["الوحدة"], item["المورد"], item["الاتجاه"], item["مستوى المخاطر"], int(item["عدد الشهور المتاحة"]), pct(item["التذبذب %"])],
    })
    st.dataframe(details, use_container_width=True, hide_index=True)

with cat_tab:
    selected_cat = st.selectbox("اختر التصنيف", sorted(filtered["التصنيف"].dropna().astype(str).unique()))
    g = filtered[filtered["التصنيف"].astype(str) == selected_cat]
    cat_row = cats[cats["التصنيف"].astype(str) == selected_cat].iloc[0]
    a, b, c, d, e = st.columns(5)
    a.metric("عدد الأصناف", f"{len(g):,}")
    b.metric("متوسط التغير", pct(g["التغير الكلي %"].mean()))
    c.metric("مرتفعة", f"{(g['الاتجاه']=='مرتفع').sum():,}")
    d.metric("منخفضة", f"{(g['الاتجاه']=='منخفض').sum():,}")
    e.metric("متوسط المخاطر", f"{g['درجة المخاطر'].mean():.1f}/100")

    cat_trend = g[months].mean().reset_index()
    cat_trend.columns = ["الشهر", "متوسط السعر"]
    c1, c2 = st.columns(2)
    with c1:
        fig = px.line(cat_trend, x="الشهر", y="متوسط السعر", markers=True, title=f"متوسط سعر تصنيف {selected_cat}")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(g.sort_values("التغير الكلي %"), x="التغير الكلي %", y="اسم الصنف", orientation="h", title="تغير كل صنف داخل التصنيف", text_auto=".2f")
        fig.update_layout(height=max(420, min(900, 25 * len(g))), yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

    heat = g.set_index("اسم الصنف")[months]
    heat_change = heat.pct_change(axis=1) * 100
    fig = px.imshow(heat_change, aspect="auto", title="الخريطة الحرارية لنسب التغير الشهري", labels={"x":"الشهر", "y":"الصنف", "color":"التغير %"})
    fig.update_layout(height=max(450, min(950, 24 * len(g))))
    st.plotly_chart(fig, use_container_width=True)

with table_tab:
    display_cols = ["كود الصنف", "اسم الصنف", "التصنيف", "الوحدة", "المورد"] + months + ["التغير الكلي %", "التذبذب %", "الاتجاه", "درجة المخاطر", "مستوى المخاطر"]
    st.dataframe(filtered[display_cols].style.format({m:"{:.2f}" for m in months} | {"التغير الكلي %":"{:.2f}%", "التذبذب %":"{:.2f}%", "درجة المخاطر":"{:.1f}"}), use_container_width=True, height=520)

    csv = filtered[display_cols].to_csv(index=False).encode("utf-8-sig")
    st.download_button("تحميل تحليل الأصناف CSV", csv, "تحليل_الأصناف.csv", "text/csv")
    if not cats.empty:
        csv_cat = cats.to_csv(index=False).encode("utf-8-sig")
        st.download_button("تحميل تحليل التصنيفات CSV", csv_cat, "تحليل_التصنيفات.csv", "text/csv")

st.markdown('<div class="note"><b>منهجية الحساب:</b> نسبة التغير الكلية = (آخر سعر متاح ÷ أول سعر متاح − 1) × 100. الأصفار تُستبعد لأنها تمثل عدم وجود سعر في الملف.</div>', unsafe_allow_html=True)
