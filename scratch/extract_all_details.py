import os
import sys
import docx
import pypdf
import pptx
import zipfile

sys.stdout.reconfigure(encoding='utf-8')

wipo_dir = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\bài báo sáng chế\WIPO 26-27.8.2026-20260826T025931Z-1-001\WIPO 26-27.8.2026"

out_dir = r"C:\Users\Admin\.gemini\antigravity-ide\brain\2664ec95-b734-4052-ac21-ea79e5aed358\scratch"
os.makedirs(out_dir, exist_ok=True)

# 1. Parse PDF: Aug 26 PM - MST preliminary analysis VN strategic technologies.pdf
pdf_path = os.path.join(wipo_dir, "Aug 26 PM - MST preliminary analysis VN strategic technologies.pdf")
print("=== ANALYSIS OF STRATEGIC TECHNOLOGIES PDF ===")
reader = pypdf.PdfReader(pdf_path)
for i, page in enumerate(reader.pages):
    txt = page.extract_text()
    if txt and any(k in txt.lower() for k in ['viet nam', 'technology', 'patent', 'strategic', 'ai', 'semiconductor', 'biotech', 'robotics']):
        print(f"\n--- Page {i+1} ---")
        print(txt[:1000])

# 2. Check Scopus docx file
scopus_docx = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\bài báo sáng chế\Xu hướng sáng chế trong nền kinh tế số Dữ liệu từ scopus.docx"
print("\n=== SCOPUS DOCX FILE CHECK ===")
try:
    doc = docx.Document(scopus_docx)
    print("Successfully opened with python-docx!")
    for p in doc.paragraphs[:20]:
        if p.text.strip():
            print(p.text)
except Exception as e:
    print(f"Error opening with python-docx: {e}")
    # Try zipfile inspection if docx is xml based
    try:
        with zipfile.ZipFile(scopus_docx, 'r') as z:
            print("Zip contents:", z.namelist()[:10])
    except Exception as e2:
        print(f"Not a valid zip file either: {e2}")

