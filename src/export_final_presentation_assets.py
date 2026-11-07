"""
Exports the streamlined 10-Slide presentation to:
1. FINAL_PRESENTATION/SIC_AI_Capstone_Final.pdf
2. SIC_Documents_VI/SIC_AI_Capstone Project_Presentation Slide Template.pdf
3. FINAL_PRESENTATION/Slide_Images/slide_01.png to slide_10.png (1920x1080 HD)
"""

import os
import glob
import win32com.client

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FINAL_DIR = os.path.join(BASE_DIR, "FINAL_PRESENTATION")
VI_DIR = os.path.join(BASE_DIR, "SIC_Documents_VI")
PPTX_PATH = os.path.join(FINAL_DIR, "SIC_AI_Capstone_Final.pptx")
PDF_FINAL = os.path.join(FINAL_DIR, "SIC_AI_Capstone_Final.pdf")
PDF_VI = os.path.join(VI_DIR, "SIC_AI_Capstone Project_Presentation Slide Template.pdf")
IMG_DIR = os.path.join(FINAL_DIR, "Slide_Images")
os.makedirs(IMG_DIR, exist_ok=True)

# Clean old slide_*.png files before re-exporting
for old_png in glob.glob(os.path.join(IMG_DIR, "slide_*.png")):
    try:
        os.remove(old_png)
    except OSError:
        pass

ppt_app = win32com.client.Dispatch("PowerPoint.Application")
try:
    pres = ppt_app.Presentations.Open(PPTX_PATH, WithWindow=False)
    print(f"Opened {PPTX_PATH} with {pres.Slides.Count} slides.")

    pres.SaveAs(PDF_FINAL, 32)
    try:
        pres.SaveAs(PDF_VI, 32)
    except Exception as e:
        print("VI PDF save note:", e)

    for idx, slide in enumerate(pres.Slides, start=1):
        png_name = f"slide_{idx:02d}.png"
        png_path = os.path.join(IMG_DIR, png_name)
        slide.Export(png_path, "PNG", 1920, 1080)
        print(f"Exported Slide {idx:02d} -> {png_name}")

    pres.Close()
finally:
    ppt_app.Quit()

print("EXPORT COMPLETE!")
