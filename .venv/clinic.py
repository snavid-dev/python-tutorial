import docx
from docx.shared import Pt, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


# تنظیمات برای راست‌چین کردن متن‌ها
def set_rtl_run(run):
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    run.font.rtl = True


def add_rtl_paragraph(doc, text, bold=False, size=11, alignment=WD_ALIGN_PARAGRAPH.RIGHT):
    p = doc.add_paragraph()
    p.alignment = alignment
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.rtl = True
    return p


def create_pain_clinic_dossier():
    doc = docx.Document()

    # تنظیم حاشیه‌های صفحه A4
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

    # --- فوتر (پایین تمام صفحات) ---
    # نکته: چون سکشن‌ها تعریف شده‌اند، از سکشن آخر استفاده می‌کنیم یا داخل حلقه بالا
    section = doc.sections[0]
    footer = section.footer
    footer_p = footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_p.add_run(
        "\n________________________            ________________________            ________________________\n")
    run.font.size = Pt(10)
    run_text = footer_p.add_run(
        "   تاریخ                                           امضای داکتر                                           امضای مریض   ")
    run_text.font.name = 'Arial'
    run_text.font.size = Pt(10)
    run_text.font.rtl = True

    # ================= صفحه ۱: دوسیه عمومی =================
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = title_p.add_run("دوسیه جامع مریضان کلینیک درد")
    run_title.bold = True
    run_title.font.size = Pt(16)
    run_title.font.name = 'Arial'

    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = subtitle_p.add_run("کلینیک درد – ایران کلینیک هرات")
    run_sub.bold = True
    run_sub.font.size = Pt(14)
    run_sub.font.name = 'Arial'

    doc.add_paragraph()

    # جدول مشخصات
    add_rtl_paragraph(doc, "مشخصات عمومی مریض:", bold=True, size=12)

    table = doc.add_table(rows=5, cols=4)
    table.style = 'Table Grid'

    data = [
        ("شماره دوسیه:", "", "تاریخ:", ""),
        ("نام و تخلص:", "", "نام پدر:", ""),
        ("سن:", "", "جنس:", ""),
        ("شماره تماس:", "", "آدرس:", ""),
        ("شغل:", "", "وضعیت تأهل:", "")
    ]

    for r_idx, row_data in enumerate(data):
        row = table.rows[r_idx]
        cells = row.cells
        # چیدمان ستون‌ها برای فارسی

        p = cells[3].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(row_data[0])
        set_rtl_run(r)

        p = cells[2].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(row_data[1])
        set_rtl_run(r)

        p = cells[1].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(row_data[2])
        set_rtl_run(r)

        p = cells[0].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(row_data[3])
        set_rtl_run(r)

    doc.add_paragraph()

    # شرح حال
    add_rtl_paragraph(doc, "شرح حال و ارزیابی:", bold=True, size=12)
    add_rtl_paragraph(doc, "شکایت اصلی: ____________________________________________________________________")
    add_rtl_paragraph(doc, "محل درد: _____________________________   مدت درد: _____________________________")

    doc.add_paragraph()
    p_vas = doc.add_paragraph()
    p_vas.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_vas = p_vas.add_run("شدت درد (VAS):")
    r_vas.bold = True
    r_vas.font.name = 'Arial'
    r_vas.font.rtl = True

    p_scale = doc.add_paragraph()
    p_scale.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_scale = p_scale.add_run(
        "0  -----  1  -----  2  -----  3  -----  4  -----  5  -----  6  -----  7  -----  8  -----  9  -----  10")
    r_scale.bold = True

    p_scale_lbl = doc.add_paragraph()
    p_scale_lbl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_scale_lbl = p_scale_lbl.add_run(
        "(بدون درد)                                                                                             (بدترین درد)")
    r_scale_lbl.font.size = Pt(8)

    doc.add_paragraph()

    # سوابق
    add_rtl_paragraph(doc, "سوابق پزشکی:", bold=True)
    p_hist = doc.add_paragraph()
    p_hist.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_hist = p_hist.add_run("☐ سابقه دیابت      ☐ فشار خون      ☐ بیماری قلبی      ☐ سابقه جراحی: __________________")
    run_hist.font.name = 'Arial'
    run_hist.font.rtl = True

    doc.add_paragraph()

    # تشخیص
    add_rtl_paragraph(doc, "تشخیص نهایی:", bold=True)
    add_rtl_paragraph(doc, "___________________________________________________________________________________")
    add_rtl_paragraph(doc, "___________________________________________________________________________________")

    doc.add_paragraph()

    # پلن درمانی
    add_rtl_paragraph(doc, "انتخاب نوع تداوی (Treatment Plan):", bold=True)
    p_plan = doc.add_paragraph()
    p_plan.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_plan = p_plan.add_run(
        "☐ PRP Therapy\n\n☐ تزریق ژل (Gel Injection)\n\n☐ تزریق کورتون (Cortisone)\n\n☐ تزریق ارتروپویتین (Erythropoietin)")
    run_plan.font.name = 'Arial'
    run_plan.font.size = Pt(12)
    run_plan.font.rtl = True

    doc.add_page_break()

    # ================= صفحه ۲: PRP =================
    title_p2 = doc.add_paragraph()
    title_p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t2 = title_p2.add_run("دوسیه اختصاصی تزریق PRP")
    run_t2.bold = True
    run_t2.font.size = Pt(14)
    run_t2.font.name = 'Arial'

    doc.add_paragraph()

    table_prp = doc.add_table(rows=2, cols=4)
    table_prp.style = 'Table Grid'

    lbls = ["محل تزریق:", "نوع PRP:", "تعداد جلسات:", "فاصله جلسات:"]
    for i, txt in enumerate(reversed(lbls)):
        p = table_prp.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(txt)
        r.bold = True
        set_rtl_run(r)

    for cell in table_prp.rows[1].cells:
        cell.text = "\n"

    doc.add_paragraph()
    add_rtl_paragraph(doc, "جدول پیگیری جلسات:", bold=True)

    track_table = doc.add_table(rows=5, cols=6)
    track_table.style = 'Table Grid'

    headers = ["جلسه", "تاریخ تزریق", "دوز (ml)", "درد قبل از تزریق", "بهبود (1 ماه)", "بهبود (3 ماه)"]
    for i, h in enumerate(reversed(headers)):
        cell = track_table.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        set_rtl_run(r)
        tc_pr = cell._element.get_or_add_tcPr()
        shd = docx.oxml.shared.OxmlElement("w:shd")
        shd.set(docx.oxml.ns.qn("w:fill"), "E7E7E7")  # رنگ خاکستری روشن
        tc_pr.append(shd)

    for i in range(1, 5):
        cell = track_table.rows[i].cells[5]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(str(i))

    doc.add_paragraph()
    add_rtl_paragraph(doc, "یادداشت داکتر / نام داکتر:")
    add_rtl_paragraph(doc, "___________________________________________________________________________")

    doc.add_page_break()

    # ================= صفحه ۳: ژل / کورتون =================
    title_p3 = doc.add_paragraph()
    title_p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t3 = title_p3.add_run("دوسیه اختصاصی تزریق ژل / کورتون")
    run_t3.bold = True
    run_t3.font.size = Pt(14)
    run_t3.font.name = 'Arial'

    doc.add_paragraph()

    info_table = doc.add_table(rows=2, cols=2)
    info_table.style = 'Table Grid'

    c1 = info_table.rows[0].cells[1]
    p = c1.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("نوع تزریق:   ☐ ژل      ☐ کورتون")
    set_rtl_run(r)

    c0 = info_table.rows[0].cells[0]
    p = c0.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("نام دارو (Brand): _____________")
    set_rtl_run(r)

    c1 = info_table.rows[1].cells[1]
    p = c1.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("محل تزریق: ______________")
    set_rtl_run(r)

    c0 = info_table.rows[1].cells[0]
    p = c0.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("دوز تزریقی: ______________")
    set_rtl_run(r)

    doc.add_paragraph()

    add_rtl_paragraph(doc, "جدول ثبت تزریقات:", bold=True)

    track_table_3 = doc.add_table(rows=4, cols=5)
    track_table_3.style = 'Table Grid'

    headers_3 = ["تاریخ تزریق", "درد قبل", "درد بعد (1 ماه)", "درصد بهبود", "عوارض جانبی"]
    for i, h in enumerate(reversed(headers_3)):
        cell = track_table_3.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        set_rtl_run(r)
        tc_pr = cell._element.get_or_add_tcPr()
        shd = docx.oxml.shared.OxmlElement("w:shd")
        shd.set(docx.oxml.ns.qn("w:fill"), "E7E7E7")  # رنگ خاکستری روشن
        tc_pr.append(shd)

    # --- اینجا جایی بود که کد شما از تابع بیرون زده بود (تصحیح شده) ---
    doc.add_paragraph()
    add_rtl_paragraph(doc, "توضیحات تکمیلی داکتر:")
    doc.add_paragraph("_" * 70)

    doc.add_page_break()

    # ================= صفحه ۴: ارتروپویتین =================
    title_p4 = doc.add_paragraph()
    title_p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t4 = title_p4.add_run("دوسیه اختصاصی تزریق ارتروپویتین")
    run_t4.bold = True
    run_t4.font.size = Pt(14)
    run_t4.font.name = 'Arial'

    doc.add_paragraph()

    box_table = doc.add_table(rows=1, cols=1)
    box_table.style = 'Table Grid'
    p_ind = box_table.rows[0].cells[0].paragraphs[0]
    p_ind.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r_ind = p_ind.add_run("اندیکاسیون (علت) تزریق: __________________________________________________")
    set_rtl_run(r_ind)

    doc.add_paragraph()

    spec_table = doc.add_table(rows=1, cols=3)
    spec_table.style = 'Table Grid'

    labels_4 = ["محل تزریق", "دوز دارو", "تعداد جلسات"]
    for i, txt in enumerate(reversed(labels_4)):
        p = spec_table.rows[0].cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(txt + "\n\n")
        r.bold = True
        set_rtl_run(r)

    doc.add_paragraph()

    add_rtl_paragraph(doc, "جدول پیگیری درمان:", bold=True)

    track_table_4 = doc.add_table(rows=5, cols=6)
    track_table_4.style = 'Table Grid'

    headers_4 = ["جلسه", "تاریخ", "درد قبل", "درد (1 ماه)", "درد (3 ماه)", "عوارض / نتیجه"]
    for i, h in enumerate(reversed(headers_4)):
        cell = track_table_4.rows[0].cells[i]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h)
        r.bold = True
        set_rtl_run(r)
        tc_pr = cell._element.get_or_add_tcPr()
        shd = docx.oxml.shared.OxmlElement("w:shd")
        shd.set(docx.oxml.ns.qn("w:fill"), "E7E7E7")  # رنگ خاکستری روشن
        tc_pr.append(shd)

    for i in range(1, 5):
        cell = track_table_4.rows[i].cells[5]
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(str(i))

    doc.add_paragraph()
    add_rtl_paragraph(doc, "دستورات بعدی داکتر:")
    doc.add_paragraph("_" * 70)

    # ذخیره فایل
    file_path = "Pain_Clinic_File_Herat.docx"
    doc.save(file_path)
    print(f"فایل با موفقیت ساخته شد: {file_path}")


# اجرا
if __name__ == "__main__":
    create_pain_clinic_dossier()