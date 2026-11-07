"""
Exports FINAL_SHORT_PRESENTATION PPTX decks to PDF and 1920x1080 PNG images
for full visual inspection.
"""

import os
import win32com.client

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE_DIR, "FINAL_SHORT_PRESENTATION")
IMG_DIR = os.path.join(OUT_DIR, "Slide_Images")
os.makedirs(IMG_DIR, exist_ok=True)

PPTX_SHORT = os.path.join(OUT_DIR, "SIC_AI_Capstone_Short_Final.pptx")
PDF_SHORT = os.path.join(OUT_DIR, "SIC_AI_Capstone_Short_Final.pdf")
PPTX_BACKUP = os.path.join(OUT_DIR, "Q&A_Backup_Slides.pptx")
PDF_BACKUP = os.path.join(OUT_DIR, "Q&A_Backup_Slides.pdf")

ppt_app = win32com.client.Dispatch("PowerPoint.Application")
try:
    # 1. Export Short Deck (9 slides)
    pres1 = ppt_app.Presentations.Open(PPTX_SHORT, WithWindow=False)
    pres1.SaveAs(PDF_SHORT, 32)
    for idx, slide in enumerate(pres1.Slides, start=1):
        png_path = os.path.join(IMG_DIR, f"slide_{idx:02d}.png")
        slide.Export(png_path, "PNG", 1920, 1080)
        print(f"Rendered slide_{idx:02d}.png")
    pres1.Close()

    # 2. Export Backup Deck (6 slides)
    pres2 = ppt_app.Presentations.Open(PPTX_BACKUP, WithWindow=False)
    pres2.SaveAs(PDF_BACKUP, 32)
    for idx, slide in enumerate(pres2.Slides, start=1):
        png_path = os.path.join(IMG_DIR, f"backup_{idx:02d}.png")
        slide.Export(png_path, "PNG", 1920, 1080)
        print(f"Rendered backup_{idx:02d}.png")
    pres2.Close()
finally:
    ppt_app.Quit()

print("ALL PDFs AND PNGs RENDERED SUCCESSFULLY!")

