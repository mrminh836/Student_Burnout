"""
Builds FINAL_CAPSTONE/02_Final_Evaluation_Sheet.docx
Strictly conforms to official Samsung Innovation Campus (SIC) template,
cleans student self-evaluation from panel score fields and TRL column,
standardizes project title and team name.
"""
import os
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

BASE_DIR = r"c:\Users\mrmin\Desktop\Sudent_burnout"
SRC_TEMPLATE = os.path.join(BASE_DIR, "templates_backup", "SIC_AI_Capstone Project_Evaluation Sheet.docx")
OUT_PATH = os.path.join(BASE_DIR, "FINAL_CAPSTONE", "02_Final_Evaluation_Sheet.docx")


def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def build_evaluation_sheet():
    doc = Document(SRC_TEMPLATE)

    # Table 0: Score Table
    t0 = doc.tables[0]
    # Row 0: COURSE | AI Course | EVALUATOR | Name & Signature
    t0.cell(0, 1).text = "AI Course (Samsung Innovation Campus)"
    t0.cell(0, 3).text = ""  # Left blank for Evaluator Name & Signature

    # Row 1: PROJECT TITLE | ... | TEAM NAME | ...
    t0.cell(1, 1).text = (
        "Xây dựng, huấn luyện và so sánh các mô hình học máy cho bài toán dự đoán tỷ lệ bỏ học "
        "dựa trên các chỉ số về căng thẳng và sự tham gia học tập"
    )
    t0.cell(1, 3).text = "Nhóm 03: Đỗ Quang Minh (Trưởng nhóm), Vũ Quang Huy"

    # Ensure all evaluator score cells and comment in Table 0 are clean for printing
    # t0 has 30 rows. Evaluator comment is at row 29.
    if len(t0.rows) > 29:
        t0.cell(29, 0).text = ""

    # Table 1: Technology Readiness Level Assessment (TRL 1 to 9)
    # Strictly leave column 2 (LEVEL) blank as required for evaluation panels!
    t1 = doc.tables[1]
    for row_idx in range(1, len(t1.rows)):
        if len(t1.rows[row_idx].cells) >= 3:
            t1.rows[row_idx].cells[2].text = ""

    # Polish formatting
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                set_cell_margins(cell, top=60, bottom=60, left=100, right=100)
                for p in cell.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    p.paragraph_format.space_after = Pt(1)
                    for run in p.runs:
                        run.font.name = "Arial"
                        if run.font.size is None or run.font.size > Pt(10.5):
                            run.font.size = Pt(9.5)

    doc.save(OUT_PATH)
    print(f"Generated successfully: {OUT_PATH}")


if __name__ == "__main__":
    build_evaluation_sheet()

