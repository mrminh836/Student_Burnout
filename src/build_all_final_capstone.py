"""
Orchestrates building all deliverables for FINAL_CAPSTONE:
1. 01_Final_Action_Plan.docx
2. 02_Final_Evaluation_Sheet.docx
3. 03_Final_Report.docx
4. 04_Final_Presentation.pptx
5. 05_Final_WBS.xlsx
6. 06_Consistency_Audit.xlsx
7. 07_Change_Log.md (already generated)
"""
import os
import sys

BASE_DIR = r"c:\Users\mrmin\Desktop\Sudent_burnout"
SRC_DIR = os.path.join(BASE_DIR, "src")
SCRATCH_DIR = r"C:\Users\mrmin\.gemini\antigravity\brain\66192d60-1d52-48bc-8312-555617a75eaf\scratch"
sys.path.insert(0, SRC_DIR)
sys.path.insert(0, SCRATCH_DIR)

print("--- 1. Building 01_Final_Action_Plan.docx ---")
import build_final_capstone_action_plan
build_final_capstone_action_plan.build_action_plan()

print("--- 2. Building 02_Final_Evaluation_Sheet.docx ---")
import build_final_capstone_evaluation_sheet
build_final_capstone_evaluation_sheet.build_evaluation_sheet()

print("--- 3. Building 05_Final_WBS.xlsx ---")
import build_final_capstone_wbs
build_final_capstone_wbs.build_wbs()

print("--- 4. Building 04_Final_Presentation.pptx (9 Slides) ---")
import build_final_capstone_presentation_9slides
build_final_capstone_presentation_9slides.build_9slide_presentation()

print("--- 5. Building 03_Final_Report.docx ---")
import build_final_capstone_report
build_final_capstone_report.build_final_report()

print("--- 6. Updating 06_Consistency_Audit.xlsx ---")
import build_audit

print("\nAll FINAL_CAPSTONE deliverables built successfully!")

