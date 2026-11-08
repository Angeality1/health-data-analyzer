#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
تطبيق ويب لتحليل بيانات الوحدات الصحية
Health Data Analyzer - Web Application v4.0
مبني على النسخة Desktop المحسّنة
"""

import streamlit as st
import pandas as pd
import io
from datetime import datetime
from pathlib import Path
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.util import Pt
import re
import daily_slide

APP_VERSION = "4.2"
APP_UPDATED = "07/11/2026"
SUPPORT_PHONE = "01559600495"
SUPPORT_WA = "https://wa.me/201559600495"

# Page configuration
st.set_page_config(
    page_title="برنامج تحليل بيانات الوحدات الصحية",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS: surveillance-themed look (v4.2) — RTL, gradient backgrounds, cards, coloured sidebar
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
.stMarkdown, .stMarkdown *, label, button p, input, textarea, [data-baseweb="select"] *, .stCaption, h1, h2, h3 { font-family: 'Cairo', 'Segoe UI', Tahoma, sans-serif; }
.stApp {
    background:
        radial-gradient(circle at 10% 6%, rgba(27,153,139,.16), transparent 36%),
        radial-gradient(circle at 92% 18%, rgba(31,95,168,.16), transparent 40%),
        radial-gradient(circle at 50% 100%, rgba(232,113,10,.07), transparent 45%),
        linear-gradient(160deg, #eaf4fb 0%, #f8fcff 48%, #ebf7f2 100%);
}
.stApp::before {
    content: ""; position: fixed; inset: 0; pointer-events: none; z-index: 0; opacity: .9;
    background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='96' height='96'><g fill='none' stroke='%231f5fa8' stroke-opacity='.07' stroke-width='2'><path d='M48 38v20M38 48h20'/><circle cx='14' cy='14' r='4'/><circle cx='82' cy='76' r='3'/><path d='M8 80l6-8 6 8 6-8'/></g></svg>");
}
[data-testid="stMain"], .main { direction: rtl; text-align: right; }
.block-container { padding-top: 1.2rem; max-width: 1150px; position: relative; z-index: 1; }
header[data-testid="stHeader"] { background: transparent; }

/* ---- banner ---- */
.hero { position: relative; overflow: hidden; border-radius: 22px; padding: 28px 34px 56px; margin-bottom: 22px; color: #fff;
    background: linear-gradient(120deg, #0b3c5d 0%, #145a8d 48%, #1b998b 100%); box-shadow: 0 14px 34px rgba(11,60,93,.28); }
.hero::after { content: ""; position: absolute; width: 340px; height: 340px; left: -90px; top: -150px; border-radius: 50%;
    background: radial-gradient(circle, rgba(255,255,255,.20), transparent 70%); }
.hero-badge { display: inline-block; background: rgba(255,255,255,.18); border: 1px solid rgba(255,255,255,.35); padding: 4px 14px; border-radius: 999px; font-weight: 700; font-size: 14px; }
.hero h1 { color: #fff !important; text-align: right !important; font-size: 2.1rem; font-weight: 800; margin: 12px 0 4px; padding: 0; }
.hero-sub { margin: 0 0 12px; opacity: .93; font-size: 1.05rem; }
.chips span { display: inline-block; margin-left: 8px; background: rgba(0,0,0,.22); padding: 3px 12px; border-radius: 999px; font-size: 13px; }
.ecg { position: absolute; right: 0; left: 0; bottom: 4px; width: 100%; height: 54px; opacity: .85; }
.ecg polyline { fill: none; stroke: #8ff7de; stroke-width: 3; stroke-linejoin: round; stroke-dasharray: 1500; stroke-dashoffset: 1500; animation: ecgdash 5s linear infinite; filter: drop-shadow(0 0 4px #5ef0cf); }
@keyframes ecgdash { to { stroke-dashoffset: 0; } }

/* ---- cards & steps ---- */
div[class*="st-key-card_"] { background: rgba(255,255,255,.93); border: 1px solid #d3e4f1; border-radius: 18px; box-shadow: 0 8px 24px rgba(11,60,93,.09); padding: 1.1rem 1.4rem 1.3rem; margin-bottom: 18px; }
.step { display: flex; align-items: center; gap: 14px; margin: 2px 0 14px; direction: rtl; }
.step .num { flex: 0 0 auto; width: 44px; height: 44px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 20px; color: #fff;
    background: linear-gradient(135deg, #1f5fa8, #1b998b); box-shadow: 0 4px 12px rgba(31,95,168,.35); }
.step .st-title { font-weight: 800; font-size: 1.25rem; color: #0b3c5d; }
.step .st-sub { font-size: .92rem; color: #5b7a92; margin-top: 1px; }
div[data-testid="stMetric"] { background: linear-gradient(160deg, #ffffff, #f0f7fd); border: 1px solid #d3e4f1; border-top: 4px solid #1b998b; border-radius: 14px; padding: 10px 14px; box-shadow: 0 3px 10px rgba(11,60,93,.06); }
div[data-testid="stMetricLabel"] { color: #4d6b84; }
div[data-testid="stMetricValue"] { color: #0b3c5d; font-weight: 800; }
.stButton > button, .stDownloadButton > button, [data-testid="stButton"] button, [data-testid="stDownloadButton"] button { width: 100%; border: 0; border-radius: 12px; padding: 12px 10px; font-size: 18px; font-weight: 700; color: #fff;
    background: linear-gradient(120deg, #1f5fa8, #1b998b); box-shadow: 0 6px 16px rgba(31,95,168,.28); transition: transform .12s, box-shadow .12s; }
.stButton > button:hover, .stDownloadButton > button:hover { color: #fff; transform: translateY(-2px); box-shadow: 0 10px 22px rgba(31,95,168,.36); }
.success-box { padding: 22px; margin: 6px 0 12px; text-align: center; border-radius: 16px; color: #0b3c5d; font-weight: 700;
    background: linear-gradient(135deg, #d9f5ec, #e3f0fb); border: 1px solid #b4e3d3; box-shadow: 0 6px 18px rgba(27,153,139,.18); }
[data-testid="stFileUploader"] section { border: 2px dashed #5aa6d6; border-radius: 14px; background: #f3f9fe; }

/* ---- sidebar ---- */
section[data-testid="stSidebar"] { direction: rtl; text-align: right; background: linear-gradient(185deg, #0b3c5d 0%, #145a8d 58%, #1b998b 130%); }
section[data-testid="stSidebar"] * { color: #eaf6ff; }
.sb-card { background: rgba(255,255,255,.10); border: 1px solid rgba(255,255,255,.22); border-radius: 14px; padding: 12px 14px; margin-bottom: 14px; }
.sb-title { font-size: 1.2rem; font-weight: 800; margin-bottom: 6px; }
.sb-badge { display: inline-block; background: #1b998b; color: #fff !important; padding: 2px 11px; border-radius: 999px; font-weight: 700; font-size: 13px; margin-left: 6px; }
.sb-badge2 { display: inline-block; background: rgba(0,0,0,.25); padding: 2px 11px; border-radius: 999px; font-size: 13px; }
.sb-card ul { margin: 6px 0 0; padding-right: 18px; font-size: .93rem; line-height: 1.8; }
.sb-phone { font-size: 1.5rem; font-weight: 800; letter-spacing: 1px; direction: ltr; text-align: center; margin: 6px 0; }
.sb-wa { display: block; text-align: center; background: #25d366; color: #fff !important; font-weight: 800; padding: 8px; border-radius: 10px; text-decoration: none; }
.footer { text-align: center; color: #6c8aa1; font-size: .85rem; margin: 22px 0 6px; }
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

HERO_HTML = """
<div class="hero">
  <span class="hero-badge">🛡️ الترصد الوبائي</span>
  <h1>منظومة تحليل بيانات الترصد الوبائي</h1>
  <p class="hero-sub">الإدارة الصحية بالمحلة الكبرى ثان · القسم الوقائي · وحدة الوبائيات والترصد</p>
  <div class="chips"><span>الإصدار %VERSION%</span><span>آخر تحديث %UPDATED%</span><span>التحليل حسب محل السكن</span></div>
  <svg class="ecg" viewBox="0 0 800 100" preserveAspectRatio="none"><polyline points="0,60 90,60 120,60 140,22 160,96 182,40 202,60 330,60 360,60 380,14 402,98 424,50 444,60 560,60 590,60 610,28 630,92 652,52 672,60 800,60"/></svg>
</div>
""".replace("%VERSION%", APP_VERSION).replace("%UPDATED%", APP_UPDATED)

# Initialize session state
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False
if 'combined_data' not in st.session_state:
    st.session_state.combined_data = None
if 'report' not in st.session_state:
    st.session_state.report = None
if 'load_info' not in st.session_state:
    st.session_state.load_info = None
if 'health_units' not in st.session_state:
    st.session_state.health_units = []

def merge_related_units(df):
    """Merge related health units"""
    unit_mapping = {
        'وحدة بلهارسيا كفر حجازى': 'وحدة كفر حجازى',
        'وحدة بلهارسيا الابشيط': 'وحدة طب أسرة  الابشيط',
        'وحدة طب أسرة  شويقى سعد': 'وحدة عزبة شويقى سعد',
        'مستشفى تكامل الدواخلية': 'مركز طب اسرة الدواخلية',
        'مستشفى تكامل المعتمدية': 'مركز طب اسرة المعتمدية',
        'مستشفى تكامل صفط تراب': 'مركز طب اسرة صفط تراب',
        'مستشفى تكامل العامرية': 'مركز طب اسرة العامرية',
        'مستشفى تكامل الهياتم': 'مركز طب اسرة الهياتم',
        'وحدة بلهارسيا القيراطية': 'وحدة القيراطية'
    }
    squeeze = lambda t: re.sub(r'\s+', ' ', str(t)).strip()
    unit_mapping = {squeeze(k): v for k, v in unit_mapping.items()}
    df['HealthOffice'] = df['HealthOffice'].map(lambda t: squeeze(t) if pd.notna(t) else t).replace(unit_mapping)
    return df

def _nk(s):
    """مفتاح مقارنة للنصوص العربية: يزيل المسافات والرموز و«ال» ويوحّد الألف والياء والتاء المربوطة"""
    s = re.sub(r"[\s_\-\u00a0]+", "", str(s))
    s = re.sub("[أإآ]", "ا", s).replace("ى", "ي").replace("ة", "ه")
    return s.replace("ال", "")


def _pick_sheets(sheet_names):
    """يرجع (شيت محل السكن، شيت الحالات الخارجية، شيت مصدر الإبلاغ). يقبل: «محل سكن» و«محل السكن» و«محل_سكن» بأي مسافات."""
    res = ext = src = None
    for s in sheet_names:
        k = _nk(s)
        if k == "محلسكن" and res is None: res = s
        elif k.startswith("محلسكن") and "خارج" in k and ext is None: ext = s
        elif ("مصدر" in k and "ابلاغ" in k) and src is None: src = s
    return res, ext, src


def _read_sheet(xl, sheet):
    """يقرأ الشيت ويكتشف صف العناوين تلقائياً (حتى لو سبقه عنوان للصفحة أو صفوف فارغة)"""
    raw = pd.read_excel(xl, sheet_name=sheet, header=None, nrows=30)
    hdr = None
    for i in range(len(raw)):
        vals = [str(v).strip() for v in raw.iloc[i].tolist()]
        if "ResidenceDistrict" in vals and "HealthOffice" in vals:
            hdr = i
            break
    if hdr is None:
        return None
    df = pd.read_excel(xl, sheet_name=sheet, header=hdr)
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all")


def load_excel_files(uploaded_files):
    """Load and process Excel files (v4.0: sheet names and header row are detected automatically)"""
    required_cols = ['ResidenceDistrict', 'HealthOffice', 'Sex', 'disease',
                     'status', 'Outcome', 'Department', 'AgeGroup', 'Month', 'Week']

    dataframes = []
    st.session_state.load_notes = []
    for file in uploaded_files:
        try:
            xl = pd.ExcelFile(file)
            res, ext, src = _pick_sheets(xl.sheet_names)
            parts, used = [], []
            if res:
                parts.append(_read_sheet(xl, res)); used.append(res.strip())
            elif src:  # ملف الترصد المحلي: مصدر الإبلاغ (الداخلي) + شيت الحالات الخارجية
                parts.append(_read_sheet(xl, src)); used.append(src.strip())
            else:
                st.error(f"❌ الملف {file.name}: لم أجد شيت «محل سكن» (ولا «محل السكن»). الشيتات الموجودة: {', '.join(s.strip() for s in xl.sheet_names)}")
                return None
            if ext:
                parts.append(_read_sheet(xl, ext)); used.append(ext.strip())
            if any(p is None for p in parts):
                st.error(f"❌ الملف {file.name}: لم أجد صف العناوين (ResidenceDistrict / HealthOffice) في أول 30 صفاً من الشيت.")
                return None
            df = pd.concat(parts, ignore_index=True)

            missing = [c for c in required_cols if c not in df.columns]
            if missing:
                st.error(f"❌ الملف {file.name} ينقصه الأعمدة: {', '.join(missing)}")
                return None

            st.session_state.load_notes.append(f"{file.name}: الشيت المستخدم «{' + '.join(used)}» ({len(df)} سجل)")
            dataframes.append(df)

        except Exception as e:
            st.error(f"❌ خطأ في قراءة الملف {file.name}: {str(e)}")
            return None

    # Combine and process
    combined = pd.concat(dataframes, ignore_index=True)
    combined = merge_related_units(combined)
    district_data = combined[combined['ResidenceDistrict'].map(_nk) == _nk('محلة ثانى')].copy()

    return district_data

def fill_cover_slide(prs, unit_name, data):
    """Fill cover slide with improved formatting"""
    slide = prs.slides[0]
    
    min_date = data['reptdate_gen'].min() if 'reptdate_gen' in data.columns else None
    max_date = data['reptdate_gen'].max() if 'reptdate_gen' in data.columns else None
    
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text
            
            if 'وحده الوبائيات و الترصد' in text:
                shape.text_frame.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"الإدارة الصحية بالمحلة الكبرى ثان\nالقسم الوقائي\nوحده الوبائيات و الترصد \n{unit_name}"
                p.font.name = 'Times New Roman (Headings)'
                p.font.size = Pt(30)
                p.font.bold = True
                p.alignment = 2
                
            elif 'الادارة الصحية ثان' in text or 'من 1/1' in text:
                if pd.notna(min_date) and pd.notna(max_date):
                    start = pd.to_datetime(min_date).strftime("%d/%m/%Y")
                    end = pd.to_datetime(max_date).strftime("%d/%m/%Y")
                    
                    shape.text_frame.clear()
                    p = shape.text_frame.paragraphs[0]
                    p.text = f"الإدارة الصحية بالمحلة الكبرى ثان\nمن {start} إلى {end}"
                    p.font.name = 'Times New Roman (Headings)'
                    p.font.size = Pt(32)
                    p.font.bold = True
                    p.alignment = 1

def fill_info_slide(prs, unit_name, data, monitor, supervisor, director):
    """Fill info slide with improved formatting"""
    slide = prs.slides[1]
    
    min_date = data['reptdate_gen'].min() if 'reptdate_gen' in data.columns else None
    max_date = data['reptdate_gen'].max() if 'reptdate_gen' in data.columns else None
    
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text
            
            if 'وحده الوبائيات و الترصد بالوحده الصحية' in text:
                shape.text_frame.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"وحده الوبائيات و الترصد\n{unit_name}"
                p.font.name = 'Arial (Body)'
                p.font.size = Pt(18)
                p.font.bold = True
                p.alignment = 1
            
            elif 'الوحدة الصحية ب' in text:
                shape.text_frame.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = unit_name
                p.font.name = 'Arial (Body)'
                p.font.size = Pt(28)
                p.font.bold = True
                p.alignment = 1
            
            elif 'الفترة من' in text:
                if pd.notna(min_date) and pd.notna(max_date):
                    start_date = pd.to_datetime(min_date)
                    end_date = pd.to_datetime(max_date)
                    
                    months_ar = {
                        1: 'يناير', 2: 'فبراير', 3: 'مارس', 4: 'أبريل',
                        5: 'مايو', 6: 'يونيو', 7: 'يوليو', 8: 'أغسطس',
                        9: 'سبتمبر', 10: 'أكتوبر', 11: 'نوفمبر', 12: 'ديسمبر'
                    }
                    
                    start_str = f"{start_date.day} {months_ar[start_date.month]} {start_date.year}"
                    end_str = f"{end_date.day} {months_ar[end_date.month]} {end_date.year}"
                    
                    shape.text_frame.clear()
                    p = shape.text_frame.paragraphs[0]
                    p.text = f"الفترة من {start_str} وحتى {end_str}"
                    p.font.name = 'Arial (Body)'
                    p.font.size = Pt(28)
                    p.font.bold = True
                    p.alignment = 1
    
    # Update staff names by shape index
    if len(slide.shapes) > 8 and slide.shapes[8].has_text_frame:
        shape = slide.shapes[8]
        shape.text_frame.clear()
        p = shape.text_frame.paragraphs[0]
        p.text = f"أ./ {monitor}" if monitor else "أ./"
        p.font.name = 'Arial (Body)'
        p.font.size = Pt(18)
        p.font.bold = True
    
    if len(slide.shapes) > 9 and slide.shapes[9].has_text_frame:
        shape = slide.shapes[9]
        shape.text_frame.clear()
        p = shape.text_frame.paragraphs[0]
        p.text = f"د./ {supervisor}" if supervisor else "د./"
        p.font.name = 'Arial (Body)'
        p.font.size = Pt(18)
        p.font.bold = True
    
    if len(slide.shapes) > 10 and slide.shapes[10].has_text_frame:
        shape = slide.shapes[10]
        shape.text_frame.clear()
        p = shape.text_frame.paragraphs[0]
        p.text = f"د./ {director}" if director else "د./"
        p.font.name = 'Arial (Body)'
        p.font.size = Pt(18)
        p.font.bold = True

def fill_demographic_slide(prs, unit_name, population, villages):
    """Fill demographic slide with improved formatting"""
    slide = prs.slides[2]
    
    for shape in slide.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text
            
            if 'البيانات الديموغرافية بالوحده الصحية' in text:
                shape.text_frame.clear()
                p = shape.text_frame.paragraphs[0]
                p.text = f"البيانات الديموغرافية بـ{unit_name} بالإدارة الصحية المحلة ثان"
                p.font.name = 'Arial (Body)'
                p.font.size = Pt(18)
                p.font.bold = True
                p.alignment = 3
            
            elif 'عــــدد السكـــان التقديري' in text:
                shape.text_frame.clear()
                
                p1 = shape.text_frame.paragraphs[0]
                p1.text = f"عــــدد السكـــان التقديري : {population if population else '______'}"
                p1.font.name = 'Arial (Body)'
                p1.font.size = Pt(18)
                p1.font.bold = True
                
                p2 = shape.text_frame.add_paragraph()
                p2.text = f"\nالقرى والعزب التابعة : {villages if villages else '______'}"
                p2.font.name = 'Arial (Body)'
                p2.font.size = Pt(20)
                p2.font.bold = True

def update_chart(slide, categories, values, series_name):
    """Update chart data"""
    for shape in slide.shapes:
        if shape.has_chart:
            chart = shape.chart
            chart_data = CategoryChartData()
            chart_data.categories = categories
            chart_data.add_series(series_name, values)
            chart.replace_data(chart_data)
            break

def generate_report(data, unit_name, population, villages, monitor, supervisor, director):
    """Generate PowerPoint report with all improvements"""
    
    # Load template
    template_path = Path(__file__).parent / 'template.pptx'
    if not template_path.exists():
        st.error("❌ لم يتم العثور على قالب PowerPoint")
        return None
    
    try:
        prs = Presentation(str(template_path))
        
        # Use all available data (no month/quarter restriction)
        q1_data = data.copy()
        
        if len(q1_data) == 0:
            st.warning("⚠️ لا توجد بيانات في الملف المحدد")
            return None
        
        # Fill slides with improved formatting
        fill_cover_slide(prs, unit_name, q1_data)
        fill_info_slide(prs, unit_name, q1_data, monitor, supervisor, director)
        fill_demographic_slide(prs, unit_name, population, villages)
        
        # Fill all charts
        # Gender chart (slide 4)
        gender_counts = q1_data['Sex'].value_counts()
        categories = ['ذكر', 'أنثى']
        values = [int(gender_counts.get('ذكر', 0)), int(gender_counts.get('أنثى', 0))]
        update_chart(prs.slides[3], categories, values, 'العدد')
        
        # Department chart (slide 5)
        dept_counts = q1_data['Department'].value_counts()
        update_chart(prs.slides[4], dept_counts.index.tolist()[:10], 
                    [int(v) for v in dept_counts.values.tolist()[:10]], 'العدد')
        
        # Outcome chart (slide 6)
        outcome_counts = q1_data['Outcome'].value_counts()
        update_chart(prs.slides[5], outcome_counts.index.tolist()[:10],
                    [int(v) for v in outcome_counts.values.tolist()[:10]], 'العدد')
        
        # Status chart (slide 7)
        status_counts = q1_data['status'].value_counts()
        update_chart(prs.slides[6], status_counts.index.tolist()[:10],
                    [int(v) for v in status_counts.values.tolist()[:10]], 'العدد')
        
        # Disease chart (slide 8)
        disease_counts = q1_data['disease'].value_counts()
        update_chart(prs.slides[7], disease_counts.index.tolist()[:10],
                    [int(v) for v in disease_counts.values.tolist()[:10]], 'العدد')
        
        # Age chart (slide 9)
        age_counts = q1_data['AgeGroup'].value_counts()
        age_order = ['[0-2]', '[>2-5]', '[>5-15]', '[>15-35]', '[>35-50]', '[>50-65]', '[>65]']
        categories = [ag for ag in age_order if ag in age_counts.index]
        values = [int(age_counts.get(ag, 0)) for ag in categories]
        update_chart(prs.slides[8], categories, values, 'العدد')
        
        # Weekly chart (slide 10)
        week_counts = q1_data['Week'].value_counts().sort_index()
        # Sort weeks: put 52, 53 first (end of previous year), then rest in order
        all_weeks = []
        if 52 in week_counts.index:
            all_weeks.append(52)
        if 53 in week_counts.index:
            all_weeks.append(53)
        regular_weeks = sorted([w for w in week_counts.index if w not in (52, 53)])
        all_weeks.extend(regular_weeks)
        categories = [f"أسبوع {w}" for w in all_weeks]
        values = [int(week_counts.get(w, 0)) for w in all_weeks]
        update_chart(prs.slides[9], categories, values, 'العدد')
        
        # Monthly chart (slide 11)
        month_counts = q1_data['Month'].value_counts().sort_index()
        month_names = {1: 'يناير', 2: 'فبراير', 3: 'مارس', 4: 'أبريل',
                       5: 'مايو', 6: 'يونيو', 7: 'يوليو', 8: 'أغسطس',
                       9: 'سبتمبر', 10: 'أكتوبر', 11: 'نوفمبر', 12: 'ديسمبر'}
        available_months = sorted(month_counts.index.tolist())
        categories = [month_names.get(m, f"شهر {m}") for m in available_months]
        values = [int(month_counts.get(m, 0)) for m in available_months]
        update_chart(prs.slides[10], categories, values, 'العدد')

        # Daily analysis slide (v4.0): daily cases + registration date + delay curve, all on the same dates
        if 'reptdate_gen' in q1_data.columns:
            ok_daily, info_daily = daily_slide.add_daily_slide(prs, q1_data, after_index=9)
            if not ok_daily:
                st.warning(f"⚠️ لم تُضَف شريحة التحليل اليومي: {info_daily}")
        else:
            st.warning("⚠️ لم تُضَف شريحة التحليل اليومي: عمود reptdate_gen غير موجود في الملف")
        
        # Save to BytesIO
        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        
        return output
        
    except Exception as e:
        st.error(f"❌ خطأ في إنشاء التقرير: {str(e)}")
        return None

# Main App (v4.2)
ALL_LABEL = "المحلة ثان (كاملة)"


def stretch(fn, *args, **kwargs):
    """عرض كامل للعنصر: width='stretch' في النسخ الحديثة وuse_container_width في الأقدم"""
    try:
        return fn(*args, width="stretch", **kwargs)
    except Exception:
        return fn(*args, use_container_width=True, **kwargs)


def card(key):
    """بطاقة بإطار (تتوافق مع نسخ Streamlit الأقدم)"""
    try:
        return st.container(border=True, key=f"card_{key}")
    except TypeError:
        try:
            return st.container(border=True)
        except TypeError:
            return st.container()


def step(n, title, sub=""):
    st.markdown(f"<div class='step'><div class='num'>{n}</div><div><div class='st-title'>{title}</div>"
                + (f"<div class='st-sub'>{sub}</div>" if sub else "") + "</div></div>", unsafe_allow_html=True)


def period_of(df):
    if 'reptdate_gen' not in df.columns:
        return None, None
    d = pd.to_datetime(df['reptdate_gen'], errors='coerce', utc=True).dt.tz_localize(None)
    return d.min(), d.max()


def slice_unit(selected_unit):
    df = st.session_state.combined_data
    if selected_unit == ALL_LABEL:
        return df.copy(), "إدارة المحلة ثان كاملة"
    return df[df['HealthOffice'] == selected_unit].copy(), selected_unit


def render_sidebar():
    with st.sidebar:
        st.markdown(f"""
        <div class="sb-card">
            <div class="sb-title">🏥 تحليل بيانات الترصد</div>
            <span class="sb-badge">النسخة {APP_VERSION}</span><span class="sb-badge2">{APP_UPDATED}</span>
        </div>
        <div class="sb-card">
            <b>✨ المميزات</b>
            <ul>
                <li>يعمل على الجوال والكمبيوتر</li>
                <li>يقبل «محل سكن» و«محل السكن» ويكتشف صف العناوين</li>
                <li>دمج الوحدات المرتبطة</li>
                <li>معالجة الأسابيع 52 و 53</li>
                <li>شريحة التحليل اليومي: خطان (تاريخ الإبلاغ وتاريخ التسجيل) لكل الأيام حتى الجمعة</li>
                <li>معاينة سريعة قبل إنشاء التقرير</li>
            </ul>
        </div>
        <div class="sb-card">
            <b>📞 الدعم الفني</b>
            <div class="sb-phone">{SUPPORT_PHONE}</div>
            <a class="sb-wa" href="{SUPPORT_WA}" target="_blank">💬 مراسلة على واتساب</a>
        </div>
        """, unsafe_allow_html=True)


def show_preview(unit_data):
    """معاينة التحليل اليومي (خطان) قبل إنشاء التقرير"""
    if 'reptdate_gen' not in unit_data.columns:
        st.info("لا يوجد عمود reptdate_gen: لن تُضاف شريحة التحليل اليومي.")
        return
    fr = daily_slide.daily_frame(unit_data)
    if fr is None or not len(fr):
        st.info("لا توجد تواريخ إبلاغ صالحة للمعاينة.")
        return
    chart_df = pd.DataFrame({"تاريخ الإبلاغ": fr['report_n'].values, "تاريخ التسجيل": fr['insert_n'].values}, index=pd.to_datetime(fr['d']))
    try:
        import altair as alt
        names = ["تاريخ الإبلاغ", "تاريخ التسجيل"]
        long = chart_df.rename_axis("التاريخ").reset_index().melt("التاريخ", var_name="السلسلة", value_name="عدد الحالات")
        long["ترتيب"] = long["السلسلة"].map({names[0]: 0, names[1]: 1})
        color = alt.Color("السلسلة:N", scale=alt.Scale(domain=names, range=["#1F5FA8", "#1B998B"]), legend=alt.Legend(orient="bottom", title=None))
        base = alt.Chart(long).encode(
            x=alt.X("التاريخ:T", axis=alt.Axis(format="%d/%m", labelAngle=-45, title=None, tickCount=min(len(fr), 16))),
            y=alt.Y("عدد الحالات:Q", axis=alt.Axis(title="عدد الحالات")),
            color=color,
            strokeDash=alt.StrokeDash("السلسلة:N", scale=alt.Scale(domain=names, range=[[1, 0], [6, 4]]), legend=None),
            strokeWidth=alt.StrokeWidth("السلسلة:N", scale=alt.Scale(domain=names, range=[5, 2.5]), legend=None),
            order=alt.Order("ترتيب:Q"),
            tooltip=[alt.Tooltip("التاريخ:T", format="%d/%m/%Y"), "السلسلة:N", "عدد الحالات:Q"],
        )
        stretch(st.altair_chart, base.mark_line(point=True).properties(height=300))
    except Exception:
        st.line_chart(chart_df)
    diff = fr[fr['report_n'] != fr['insert_n']]
    if len(diff) == 0:
        st.success("✅ كل الأيام متطابقة: تاريخ التسجيل = تاريخ الإبلاغ.")
    else:
        st.warning(f"⚠️ يوجد {len(diff)} يوم بها اختلاف بين تاريخ الإبلاغ وتاريخ التسجيل (تسجيل متأخر أو متقدم).")
        with st.expander("عرض الأيام المختلفة"):
            stretch(st.dataframe, daily_slide.daily_table(diff), hide_index=True)
    fri = fr[fr['d'].dt.weekday == 4]
    st.caption(f"عدد الأيام في الفترة: {len(fr)} (منها {len(fri)} يوم جمعة، سُجّلت فيها {int(fri['report_n'].sum())} حالة).")


def main():
    st.markdown(HERO_HTML, unsafe_allow_html=True)
    render_sidebar()

    # Step 1: Upload Files
    with card("upload"):
        step(1, "رفع ملفات Excel", "ملف شهري واحد أو 3 ملفات (ربع سنوي) — يُقرأ شيت «محل سكن» (أو «محل السكن») تلقائياً")
        uploaded_files = st.file_uploader(
            "اختر ملفات Excel (شهر أو ربع سنوي)",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="ملف واحد (شهري) أو 3 ملفات (ربع سنوي). يُقرأ شيت «محل سكن» (أو «محل السكن») تلقائياً",
            label_visibility="collapsed"
        )

        if uploaded_files:
            st.success(f"✅ تم رفع {len(uploaded_files)} ملف")

            if stretch(st.button, "📊 تحميل وتحليل البيانات", key="load_data"):
                with st.spinner("جاري تحميل البيانات..."):
                    data = load_excel_files(uploaded_files)

                    if data is not None:
                        st.session_state.combined_data = data
                        st.session_state.report = None

                        # Get unique units and add "all" option
                        units = sorted(data['HealthOffice'].dropna().unique())
                        units.insert(0, ALL_LABEL)
                        st.session_state.health_units = units
                        st.session_state.data_loaded = True
                        st.session_state.load_info = {"notes": list(st.session_state.get('load_notes', []))}
                        st.balloons()

        info = st.session_state.get('load_info')
        if st.session_state.data_loaded and info:
            data = st.session_state.combined_data
            d0, d1 = period_of(data)
            st.markdown("**ملخص البيانات المحمّلة**")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("السجلات", f"{len(data):,}")
            c2.metric("الوحدات", len(st.session_state.health_units) - 1)
            c3.metric("الأسابيع", data['Week'].nunique())
            c4.metric("الفترة", f"{d0:%d/%m} – {d1:%d/%m}" if pd.notna(d0) and pd.notna(d1) else "—", help=f"من {d0:%d/%m/%Y} إلى {d1:%d/%m/%Y}" if pd.notna(d0) and pd.notna(d1) else None)
            for note in info["notes"]:
                st.caption("📄 " + note)

    # Step 2: Select Unit
    if st.session_state.data_loaded:
        with card("unit"):
            step(2, "اختيار الوحدة الصحية", "اختر وحدة معينة أو «المحلة ثان (كاملة)» للتحليل الشامل")
            selected_unit = st.selectbox(
                "اختر الوحدة:",
                st.session_state.health_units,
                label_visibility="collapsed"
            )
            unit_data, report_name = slice_unit(selected_unit)
            k1, k2, k3, k4 = st.columns(4)
            k1.metric("إجمالي السجلات", f"{len(unit_data):,}")
            k2.metric("عدد الأمراض", unit_data['disease'].nunique())
            k3.metric("الأقسام", unit_data['Department'].nunique())
            k4.metric("الفئات العمرية", unit_data['AgeGroup'].nunique())
            st.markdown("**معاينة التحليل اليومي** (كما ستظهر في الشريحة)")
            show_preview(unit_data)

        # Step 3: Additional Info
        with card("extra"):
            step(3, "البيانات الإضافية", "اختياري — تظهر في شرائح الغلاف والمعلومات والبيانات الديموغرافية")
            col1, col2 = st.columns(2)
            with col1:
                population = st.text_input("عدد السكان التقديري", placeholder="مثال: 15000")
                monitor = st.text_input("مراقب الترصد", placeholder="مثال: أحمد محمد")
                supervisor = st.text_input("مسئول الترصد", placeholder="مثال: د. محمد علي")
            with col2:
                villages = st.text_input("القرى والعزب التابعة", placeholder="مثال: الهياتم، دنوشر")
                director = st.text_input("مدير الوحدة", placeholder="مثال: د. فاطمة أحمد")

        # Step 4: Generate Report
        with card("report"):
            step(4, "إنشاء التقرير", "ملف PowerPoint جاهز للتنزيل")
            if stretch(st.button, "📄 إنشاء تقرير PowerPoint", key="generate", type="primary"):
                with st.spinner("جاري إنشاء التقرير... قد يستغرق دقيقة"):
                    report_file = generate_report(unit_data, report_name, population, villages, monitor, supervisor, director)
                    if report_file:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        st.session_state.report = {
                            "bytes": report_file.getvalue(),
                            "filename": f"تقرير_{report_name.replace('/', '_')}_{timestamp}.pptx",
                            "name": report_name,
                            "stats": (len(unit_data), unit_data['disease'].nunique(), unit_data['Department'].nunique(), unit_data['AgeGroup'].nunique()),
                        }

            rep = st.session_state.get("report")
            if rep:
                st.markdown(f"<div class='success-box'>✅ تم إنشاء التقرير بنجاح — {rep['name']}</div>", unsafe_allow_html=True)
                stretch(
                    st.download_button,
                    label="⬇️ تحميل التقرير PowerPoint",
                    data=rep["bytes"],
                    file_name=rep["filename"],
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
                st.caption("يبقى التقرير جاهزاً للتنزيل حتى لو غيّرت أي خيار، ولا حاجة لإعادة إنشائه.")

    st.markdown(f"<div class='footer'>برنامج تحليل بيانات الترصد · النسخة {APP_VERSION} · للدعم: {SUPPORT_PHONE}</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
