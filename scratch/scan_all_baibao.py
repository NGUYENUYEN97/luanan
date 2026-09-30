import os, fitz, sys
sys.stdout.reconfigure(encoding='utf-8')

folder = r'G:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\04_BaiBao_TrongNuoc'
files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]

output = r'G:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\scan_all_baibao.txt'
os.makedirs(os.path.dirname(output), exist_ok=True)

results = []
count_shtt = 0

for f in files:
    fp = os.path.join(folder, f)
    try:
        doc = fitz.open(fp)
        text = ""
        # scan first 2 pages
        for i in range(min(2, doc.page_count)):
            text += doc[i].get_text()
        doc.close()
        
        # Check if text looks like an article on SHTT
        text_lower = text.lower()
        if 'sở hữu trí tuệ' in text_lower or 'shtt' in text_lower:
            # get the first few non-empty lines as title
            lines = [l.strip() for l in text.split('\n') if len(l.strip()) > 10]
            title_guess = lines[0] if lines else "Unknown"
            results.append(f"FILE: {f}\nTITLE GUESS: {title_guess}\n")
            count_shtt += 1
    except Exception as e:
        results.append(f"FILE: {f}\nERROR: {str(e)}\n")

with open(output, 'w', encoding='utf-8') as out:
    out.write(f"Total PDFs: {len(files)}\n")
    out.write(f"SHTT PDFs: {count_shtt}\n")
    out.write("="*40 + "\n")
    for r in results:
        out.write(r + "-"*40 + "\n")

print(f"Scanned {len(files)} PDFs. Found {count_shtt} related to SHTT.")
