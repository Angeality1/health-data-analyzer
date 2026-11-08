# -*- coding: utf-8 -*-
"""شريحة «التحليل اليومي» لملف PowerPoint — مشتركة بين برنامج الويب (health-data-analyzer) وبرنامج الترصد المحلي (NEDSS_App).
v4.1: الرسم **خطان** على نفس تواريخ الفترة (كل الأيام بما فيها الجمعة حتى لو صفر): خط = الحالات حسب **تاريخ الإبلاغ**، وخط = الحالات حسب **تاريخ التسجيل (الإدخال)**؛
وخط ثالث رفيع منقّط (محور ثانوي) = متوسط التأخير بين التسجيل والإبلاغ بالأيام (يُلغى بوضع SHOW_DELAY = False). أسماء الأيام بالعربية في محور التواريخ وفي الجدول اليومي."""
import pandas as pd
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_LINE
from pptx.util import Pt
from pptx.oxml import parse_xml

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
DAYS_AR = {0: "الاثنين", 1: "الثلاثاء", 2: "الأربعاء", 3: "الخميس", 4: "الجمعة", 5: "السبت", 6: "الأحد"}       # weekday(): الاثنين = 0
DAYS_AR_SHORT = {0: "اثنين", 1: "ثلاثاء", 2: "أربعاء", 3: "خميس", 4: "جمعة", 5: "سبت", 6: "أحد"}
SHOW_DELAY = False                                 # الدكتور (7/11): الرسم خطان فقط (تاريخ الإبلاغ وتاريخ التسجيل). True يضيف خط متوسط التأخير على محور ثانوي
COL_REPORT, COL_INSERT, COL_DELAY = "1F5FA8", "1B998B", "7F7F7F"     # أزرق = الإبلاغ، أخضر = التسجيل، رمادي = التأخير


def daily_frame(data):
    """DataFrame بصف لكل يوم تقويمي: report_n, insert_n, avg_delay (None إن لم توجد حالات ذلك اليوم)"""
    rep = pd.to_datetime(data["reptdate_gen"], errors="coerce", utc=True).dt.tz_localize(None).dt.normalize()
    if "InsertDate" in data.columns: ins = pd.to_datetime(data["InsertDate"], errors="coerce", utc=True).dt.tz_localize(None).dt.normalize()
    else: ins = pd.Series(pd.NaT, index=data.index)
    ok = rep.notna()
    if not ok.any(): return None
    days = pd.date_range(rep[ok].min(), rep[ok].max(), freq="D")
    if "DurationBetweenInsertAndReportDate" in data.columns: delay = pd.to_numeric(data["DurationBetweenInsertAndReportDate"], errors="coerce")
    else: delay = (ins - rep).dt.days
    delay = delay.where(delay.notna(), (ins - rep).dt.days).clip(lower=0)
    df = pd.DataFrame({"d": days})
    df["report_n"] = df["d"].map(rep[ok].value_counts()).fillna(0).astype(int)
    df["insert_n"] = df["d"].map(ins.dropna().value_counts()).fillna(0).astype(int)
    avg = delay[ok].groupby(rep[ok]).mean()
    df["avg_delay"] = df["d"].map(avg).round(2)
    return df


def day_name(d, short=False):
    return (DAYS_AR_SHORT if short else DAYS_AR)[d.weekday()]


def day_label(d, long_period=False):
    """تسمية اليوم في المحور: «السبت» ثم التاريخ في سطر ثان (فترة قصيرة)، أو «سبت 31/10» (فترة طويلة)"""
    return f"{day_name(d, True)} {d:%d/%m}" if long_period else f"{day_name(d)}\n{d:%d/%m}"


def daily_table(df):
    """الجدول اليومي بالعربية: اليوم، التاريخ، الحالات حسب تاريخ الإبلاغ، الحالات حسب تاريخ التسجيل، متوسط التأخير (أيام)"""
    t = pd.DataFrame({"اليوم": [day_name(d) for d in df["d"]], "التاريخ": [f"{d:%d/%m/%Y}" for d in df["d"]],
                      "الحالات حسب تاريخ الإبلاغ": df["report_n"].astype(int), "الحالات حسب تاريخ التسجيل": df["insert_n"].astype(int),
                      "متوسط التأخير (أيام)": df["avg_delay"]})
    return t


def _style_series(ser, col, width_pt, marker, msize, dash=None):
    ser.smooth = False
    rgb = RGBColor.from_string(col)
    ser.format.line.color.rgb = rgb; ser.format.line.width = Pt(width_pt)
    if dash is not None: ser.format.line.dash_style = dash
    ser.marker.style = marker; ser.marker.size = msize
    ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = rgb; ser.marker.format.line.color.rgb = rgb


def _line_chart(slide, x, y, cx, cy, cats, rep, ins, delay, names, show_delay=True):
    cd = CategoryChartData(); cd.categories = cats
    cd.add_series(names[0], rep); cd.add_series(names[1], ins)
    if show_delay: cd.add_series(names[2], delay)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, x, y, cx, cy, cd); chart = gf.chart
    plot = chart.plots[0]
    sers = list(plot.series)
    # الخطان يتطابقان كثيراً (التسجيل في نفس يوم الإبلاغ) فيُرسم خط الإبلاغ عريضاً بعلامات كبيرة ويعلوه خط التسجيل أرفع ومتقطعاً بعلامات أصغر ليظهر الاثنان معاً
    _style_series(sers[0], COL_REPORT, 4.5, XL_MARKER_STYLE.CIRCLE, 11)
    _style_series(sers[1], COL_INSERT, 2.25, XL_MARKER_STYLE.DIAMOND, 7, MSO_LINE.DASH)
    if not show_delay: return chart
    _style_series(sers[2], COL_DELAY, 1.75, XL_MARKER_STYLE.CIRCLE, 4, MSO_LINE.DASH)
    # الخط الثالث على محور ثانوي (متوسط التأخير بالأيام)
    lc = plot._element
    ser3 = lc.findall("{%s}ser" % C_NS)[-1]; lc.remove(ser3)
    line2 = parse_xml('<c:lineChart xmlns:c="%s"><c:grouping val="standard"/><c:varyColors val="0"/></c:lineChart>' % C_NS)
    line2.append(ser3)
    for t in ('<c:marker xmlns:c="%s" val="1"/>', '<c:axId xmlns:c="%s" val="50010"/>', '<c:axId xmlns:c="%s" val="50020"/>'): line2.append(parse_xml(t % C_NS))
    lc.addnext(line2)
    area = lc.getparent(); last_ax = area.findall("{%s}valAx" % C_NS)[-1]
    cat2 = parse_xml('<c:catAx xmlns:c="%s"><c:axId val="50010"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50020"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' % C_NS)
    val2 = parse_xml('<c:valAx xmlns:c="%s" xmlns:a="%s"><c:axId val="50020"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:p><a:pPr><a:defRPr sz="1100" b="0"/></a:pPr><a:r><a:rPr lang="ar-EG" sz="1100" b="0"/><a:t>متوسط التأخير (أيام)</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:numFmt formatCode="0.00" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>' % (C_NS, A_NS))
    last_ax.addnext(cat2); cat2.addnext(val2)
    return chart


def add_daily_slide(prs, data, after_index=9, title="التحليل اليومي للحالات (تاريخ الإبلاغ وتاريخ التسجيل)"):
    """يضيف الشريحة بعد الشريحة رقم after_index (0-based؛ الافتراضي بعد شريحة الأسابيع). يرجع (ok, ملخص أو سبب)."""
    df = daily_frame(data)
    if df is None or not len(df): return False, "لا توجد تواريخ إبلاغ صالحة"
    ref = prs.slides[after_index]
    chart_shape = next(s for s in ref.shapes if getattr(s, "has_chart", False) and s.has_chart)
    slide = prs.slides.add_slide(ref.slide_layout)
    for ph in list(slide.placeholders):
        if ph.placeholder_format.type is not None and "TITLE" in str(ph.placeholder_format.type): ph.text_frame.text = title; ph.text_frame.paragraphs[0].font.size = Pt(26)
        else: ph._element.getparent().remove(ph._element)
    ref_title = next((s for s in ref.shapes if s.is_placeholder), None)
    if ref_title is not None:
        t = next((p for p in slide.placeholders), None)
        if t is not None: t.left, t.top, t.width, t.height = ref_title.left, ref_title.top, ref_title.width, ref_title.height
    long_period = len(df) > 31
    compact = len(df) > 14                                       # أكثر من أسبوعين: «خميس 29/10» في سطر واحد ومائلة رأسياً حتى لا تُقتطع التسمية
    cats = [day_label(d, compact) for d in df["d"]]
    delay_vals = [None if pd.isna(v) else float(v) for v in df["avg_delay"]]
    chart = _line_chart(slide, chart_shape.left, chart_shape.top, chart_shape.width, chart_shape.height, cats,
                        [int(v) for v in df["report_n"]], [int(v) for v in df["insert_n"]], delay_vals,
                        ("الحالات حسب تاريخ الإبلاغ", "الحالات حسب تاريخ التسجيل (الإدخال)", "متوسط التأخير بين التسجيل والإبلاغ (أيام)"), show_delay=SHOW_DELAY)
    chart.has_legend = True; chart.legend.position = XL_LEGEND_POSITION.BOTTOM; chart.legend.include_in_layout = False; chart.legend.font.size = Pt(11)
    chart.font.size = Pt(10)
    from pptx.chart.axis import ValueAxis, CategoryAxis
    area = chart._chartSpace.chart.plotArea
    prim_v = next(v for v in area.findall("{%s}valAx" % C_NS) if v.find("{%s}axId" % C_NS).get("val") != "50020")
    prim_c = next(v for v in area.findall("{%s}catAx" % C_NS) if v.find("{%s}axId" % C_NS).get("val") != "50010")
    ca = CategoryAxis(prim_c); ca.tick_labels.font.size = Pt(11 if len(df) <= 7 else 10 if len(df) <= 14 else 9 if len(df) <= 31 else 8); ca.has_major_gridlines = False
    if compact:
        bp = prim_c.find("{%s}txPr" % C_NS).find("{%s}bodyPr" % A_NS); bp.set("rot", "-5400000"); bp.set("vert", "horz")
    if long_period:                                              # فترة طويلة: عرض كل n تسمية
        skip = parse_xml('<c:tickLblSkip xmlns:c="%s" val="%d"/>' % (C_NS, 2 if len(df) <= 62 else 3)); nm = prim_c.find("{%s}noMultiLvlLbl" % C_NS)
        nm.addprevious(skip) if nm is not None else prim_c.append(skip)
    va = ValueAxis(prim_v); va.has_title = True; va.axis_title.text_frame.text = "عدد الحالات"; va.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
    va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9)
    va.tick_labels.font.size = Pt(10); va.minimum_scale = 0
    chart.has_title = False
    # نقل الشريحة لتأتي بعد after_index
    lst = prs.slides._sldIdLst; ids = list(lst); new = ids[-1]; lst.remove(new); lst.insert(after_index + 1, new)
    fri = df[df["d"].dt.weekday == 4]
    return True, {"days": int(len(df)), "fridays": int(len(fri)), "fridays_with_cases": int((fri["report_n"] > 0).sum()), "max_delay": None if df["avg_delay"].dropna().empty else float(df["avg_delay"].max())}
