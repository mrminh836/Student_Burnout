"""
Exports all Master Word (.docx) and PowerPoint (.pptx) files (both VI and EN) to PDF
using Microsoft Office COM Automation, updates Word TOC/Page fields, and renders PNG previews
for visual quality inspection.
"""

import os
import time
import shutil
import win32com.client
import pymupdf

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_FOLDER = r"c:\Users\mrmin\Desktop\drive-download-20261105T132746Z-1-001"
PREVIEW_DIR = os.path.join(BASE_DIR, "previews")
os.makedirs(PREVIEW_DIR, exist_ok=True)


def export_single_word_to_pdf(docx_p, pdf_p):
    if not os.path.exists(docx_p):
        return
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(os.path.abspath(docx_p))
        doc.Fields.Update()
        doc.Save()
        doc.ExportAsFixedFormat(os.path.abspath(pdf_p), 17)
        doc.Close(False)
        print("Exported Word PDF:", os.path.basename(pdf_p))
    except Exception as e:
        print("Word export warning:", e)
    finally:
        try:
            word.Quit()
        except Exception:
            pass
    time.sleep(0.5)


def export_single_ppt_to_pdf(pptx_p, pdf_p):
    if not os.path.exists(pptx_p):
        return
    ppt = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = ppt.Presentations.Open(os.path.abspath(pptx_p), True, False, True)
        pres.SaveAs(os.path.abspath(pdf_p), 32)
        pres.Close()
        print("Exported PowerPoint PDF:", os.path.basename(pdf_p))
    except Exception as e:
        print("PPT export warning:", e)
    finally:
        try:
            ppt.Quit()
        except Exception:
            pass
    time.sleep(0.5)


def render_previews():
    pdf_report = os.path.join(SRC_FOLDER, "SIC_AI_Capstone Project_Final Report_VI.pdf")
    if os.path.exists(pdf_report):
        doc = pymupdf.open(pdf_report)
        print(f"Final Report VI total pages: {len(doc)}")
        for p_idx in [0, 1, 2, 3, 5, 7]:
            if p_idx < len(doc):
                pix = doc[p_idx].get_pixmap(dpi=130)
                pix.save(os.path.join(PREVIEW_DIR, f"report_vi_page_{p_idx+1:02d}.png"))
        doc.close()

    pdf_slides = os.path.join(SRC_FOLDER, "SIC_AI_Capstone Project_Presentation Slide Template_VI.pdf")
    if os.path.exists(pdf_slides):
        doc = pymupdf.open(pdf_slides)
        print(f"Presentation VI total slides: {len(doc)}")
        for s_idx in [0, 1, 2, 5, 6, 9, 10, 12, 13]:
            if s_idx < len(doc):
                pix = doc[s_idx].get_pixmap(dpi=130)
                pix.save(os.path.join(PREVIEW_DIR, f"slide_vi_{s_idx+1:02d}.png"))
        doc.close()

    for lang, sub_drv, sub_ws in [
        ("VI", "Ban_Tieng_Viet_Full", "SIC_Documents_VI"),
        ("EN", "Ban_Tieng_Anh_Full", "SIC_Documents_EN"),
    ]:
        for prefix in [
            "SIC_AI_Capstone Project_Final Report",
            "SIC_AI_Capstone Project_Action Plan",
            "SIC_AI_Capstone Project_Presentation Slide Template",
        ]:
            src_pdf = os.path.join(SRC_FOLDER, f"{prefix}_{lang}.pdf")
            if os.path.exists(src_pdf):
                shutil.copy2(src_pdf, os.path.join(SRC_FOLDER, sub_drv, f"{prefix}.pdf"))
                shutil.copy2(src_pdf, os.path.join(BASE_DIR, sub_ws, f"{prefix}.pdf"))


if __name__ == "__main__":
    for docx_name, pdf_name in [
        ("SIC_AI_Capstone Project_Final Report_VI.docx", "SIC_AI_Capstone Project_Final Report_VI.pdf"),
        ("SIC_AI_Capstone Project_Final Report_EN.docx", "SIC_AI_Capstone Project_Final Report_EN.pdf"),
        ("SIC_AI_Capstone Project_Action Plan_VI.docx", "SIC_AI_Capstone Project_Action Plan_VI.pdf"),
        ("SIC_AI_Capstone Project_Action Plan_EN.docx", "SIC_AI_Capstone Project_Action Plan_EN.pdf"),
    ]:
        export_single_word_to_pdf(os.path.join(SRC_FOLDER, docx_name), os.path.join(SRC_FOLDER, pdf_name))

    for pptx_name, pdf_name in [
        ("SIC_AI_Capstone Project_Presentation Slide Template_VI.pptx", "SIC_AI_Capstone Project_Presentation Slide Template_VI.pdf"),
        ("SIC_AI_Capstone Project_Presentation Slide Template_EN.pptx", "SIC_AI_Capstone Project_Presentation Slide Template_EN.pdf"),
    ]:
        export_single_ppt_to_pdf(os.path.join(SRC_FOLDER, pptx_name), os.path.join(SRC_FOLDER, pdf_name))

    render_previews()
    print("PDF EXPORT AND VISUAL PREVIEW RENDERING COMPLETE!")
