# -*- coding: utf-8 -*-
"""شريحة «التحليل اليومي» لملف PowerPoint — مشتركة بين برنامج الويب (health-data-analyzer) وبرنامج الترصد المحلي (NEDSS_App).
رسم واحد بمحور تواريخ واحد (كل أيام الفترة بما فيها الجمعة حتى لو صفر): أعمدة = الحالات حسب تاريخ الإبلاغ، أعمدة = الحالات حسب تاريخ التسجيل (InsertDate)،
ومنحنى = متوسط الفرق بين تاريخ التسجيل وتاريخ الإبلاغ بالأيام (محور ثانوي). كل السلاسل على نفس التواريخ بالضبط."""
import copy
import pandas as pd
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Pt
from pptx.oxml import parse_xml

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
DAYS_AR = {0: "الاثنين", 1: "الثلاثاء", 2: "الأربعاء", 3: "الخميس", 4: "الجمعة", 5: "السبت", 6: "الأحد"}


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


def _combo_chart(slide, x, y, cx, cy, cats, bar1, bar2, line, names):
    cd = CategoryChartData(); cd.categories = cats
    cd.add_series(names[0], bar1); cd.add_series(names[1], bar2); cd.add_series(names[2], line)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, cd); chart = gf.chart
    plot = chart.plots[0]; bar = plot._element
    sers = bar.findall("{%s}ser" % C_NS); ser3 = sers[-1]; bar.remove(ser3)
    for tag in ("invertIfNegative",):
        for e in ser3.findall("{%s}%s" % (C_NS, tag)): ser3.remove(e)
    # ألوان الأعمدة
    for ser, col in ((sers[0], "1F77B4"), (sers[1], "9DB7D5")):
        for e in ser.findall("{%s}spPr" % C_NS): ser.remove(e)
        sp = parse_xml('<c:spPr xmlns:c="%s" xmlns:a="%s"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></c:spPr>' % (C_NS, A_NS, col))
        ser.insert(list(ser).index(ser.find("{%s}tx" % C_NS)) + 1, sp)
    # سلسلة المنحنى: خط برتقالي بعلامات
    for e in ser3.findall("{%s}spPr" % C_NS): ser3.remove(e)
    sp = parse_xml('<c:spPr xmlns:c="%s" xmlns:a="%s"><a:ln w="31750" cap="rnd"><a:solidFill><a:srgbClr val="E8710A"/></a:solidFill><a:round/></a:ln></c:spPr>' % (C_NS, A_NS))
    mk = parse_xml('<c:marker xmlns:c="%s" xmlns:a="%s"><c:symbol val="circle"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="E8710A"/></a:solidFill></c:spPr></c:marker>' % (C_NS, A_NS))
    tx = ser3.find("{%s}tx" % C_NS); i = list(ser3).index(tx) + 1; ser3.insert(i, sp); ser3.insert(i + 1, mk)
    ser3.append(parse_xml('<c:smooth xmlns:c="%s" val="0"/>' % C_NS))
    line = parse_xml('<c:lineChart xmlns:c="%s"><c:grouping val="standard"/><c:varyColors val="0"/></c:lineChart>' % C_NS)
    line.append(ser3)
    for t in ('<c:marker xmlns:c="%s" val="1"/>', '<c:axId xmlns:c="%s" val="50010"/>', '<c:axId xmlns:c="%s" val="50020"/>'): line.append(parse_xml(t % C_NS))
    bar.addnext(line)
    area = bar.getparent(); last_ax = area.findall("{%s}valAx" % C_NS)[-1]
    cat2 = parse_xml('<c:catAx xmlns:c="%s"><c:axId val="50010"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="1"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50020"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' % C_NS)
    val2 = parse_xml('<c:valAx xmlns:c="%s" xmlns:a="%s"><c:axId val="50020"/><c:scaling><c:orientation val="minMax"/><c:min val="0"/></c:scaling><c:delete val="0"/><c:axPos val="r"/><c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:p><a:pPr><a:defRPr sz="1100" b="0"/></a:pPr><a:r><a:rPr lang="ar-EG" sz="1100" b="0"/><a:t>متوسط التأخير (أيام)</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:numFmt formatCode="0.00" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="50010"/><c:crosses val="max"/><c:crossBetween val="between"/></c:valAx>' % (C_NS, A_NS))
    last_ax.addnext(cat2); cat2.addnext(val2)
    return chart


def add_daily_slide(prs, data, after_index=9, title="التحليل اليومي للحالات (حسب تاريخ الإبلاغ والتسجيل)"):
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
    cats = [f"{d:%d/%m}" for d in df["d"]]
    line_vals = [None if pd.isna(v) else float(v) for v in df["avg_delay"]]
    chart = _combo_chart(slide, chart_shape.left, chart_shape.top, chart_shape.width, chart_shape.height, cats,
                         [int(v) for v in df["report_n"]], [int(v) for v in df["insert_n"]], line_vals,
                         ("الحالات حسب تاريخ الإبلاغ", "الحالات حسب تاريخ التسجيل", "متوسط التأخير بين التسجيل والإبلاغ (أيام)"))
    chart.has_legend = True; chart.legend.position = XL_LEGEND_POSITION.BOTTOM; chart.legend.include_in_layout = False; chart.legend.font.size = Pt(11)
    chart.font.size = Pt(10)
    from pptx.chart.axis import ValueAxis, CategoryAxis
    area = chart._chartSpace.chart.plotArea
    prim_v = next(v for v in area.findall("{%s}valAx" % C_NS) if v.find("{%s}axId" % C_NS).get("val") != "50020")
    prim_c = next(v for v in area.findall("{%s}catAx" % C_NS) if v.find("{%s}axId" % C_NS).get("val") != "50010")
    ca = CategoryAxis(prim_c); ca.tick_labels.font.size = Pt(9 if len(df) <= 31 else 8); ca.has_major_gridlines = False
    if len(df) > 31:                                              # فترة طويلة: عرض كل n تسمية
        skip = parse_xml('<c:tickLblSkip xmlns:c="%s" val="%d"/>' % (C_NS, 2 if len(df) <= 62 else 3)); nm = prim_c.find("{%s}noMultiLvlLbl" % C_NS)
        nm.addprevious(skip) if nm is not None else prim_c.append(skip)
    va = ValueAxis(prim_v); va.has_title = True; va.axis_title.text_frame.text = "عدد الحالات"; va.axis_title.text_frame.paragraphs[0].font.size = Pt(11)
    va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = __import__("pptx.dml.color", fromlist=["RGBColor"]).RGBColor(0xD9, 0xD9, 0xD9)
    va.tick_labels.font.size = Pt(10)
    chart.has_title = False
    # نقل الشريحة لتأتي بعد after_index
    lst = prs.slides._sldIdLst; ids = list(lst); new = ids[-1]; lst.remove(new); lst.insert(after_index + 1, new)
    fri = df[df["d"].dt.weekday == 4]
    return True, {"days": int(len(df)), "fridays": int(len(fri)), "fridays_with_cases": int((fri["report_n"] > 0).sum()), "max_delay": None if df["avg_delay"].dropna().empty else float(df["avg_delay"].max())}
