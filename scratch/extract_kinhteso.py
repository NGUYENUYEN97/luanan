import fitz, os, sys, json
sys.stdout.reconfigure(encoding='utf-8')

folder = r'G:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\04_BaiBao_TrongNuoc'

targets = [
    'Phát triển kinh tế số tại Việt Nam Thực trạng và giải pháp.pdf',
    'Các yếu tố ảnh hưởng đến phát triển kinh tế số ở Việt Nam.pdf',
    'KINH TẾ SỐ CỦA MỸ VÀ MỘT SỐ HÀM Ý CHO VIỆT NAM.pdf',
    'Hoàn thiện pháp luật về trí tuệ nhân tạo trong nền kinh tế số tại Việt Nam.pdf',
    'Chính sách phát triển nền kinh tế số ở Việt Nam.pdf',
    'EVFTA and Digital Economy in Vietnam VIE.pdf',
    'Intellectual property rights and control in the digital economy  Examining the expansion of M-Pesa.pdf'
]

results = []

for fname in targets:
    fp = os.path.join(folder, fname)
    if not os.path.exists(fp):
        continue
    try:
        doc = fitz.open(fp)
        
        # Read first 2 pages for Abstract, Intro, Methods
        head_text = ""
        for i in range(min(2, doc.page_count)):
            head_text += doc[i].get_text()
            
        # Read last 2 pages for Conclusion, Results
        tail_text = ""
        for i in range(max(0, doc.page_count - 2), doc.page_count):
            if i >= 2: # don't overlap if very short
                tail_text += doc[i].get_text()
                
        doc.close()
        results.append({"file": fname, "head": head_text[:2000], "tail": tail_text[-2000:]})
    except Exception as e:
        print(f"Error on {fname}: {e}")

with open(r'G:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\extract_kinhteso.json', 'w', encoding='utf-8') as f:
    json.dump(results, f, ensure_ascii=False, indent=2)
    
print("Extraction complete.")
