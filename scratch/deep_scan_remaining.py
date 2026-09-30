import json
import os
import re
import fitz

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\remaining.json", "r", encoding="utf-8") as f:
    remaining = json.load(f)

base_dir = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\05_BaiBao_QuocTe"
doi_pattern = re.compile(r'10\.\d{4,9}/[-._;()/:A-Za-z0-9]+', re.IGNORECASE)

findings = []
for item in remaining:
    fn = item["filename"]
    fp = os.path.join(base_dir, fn)
    all_dois = []
    first_page_text = ""
    last_page_text = ""
    
    if fn.endswith('.pdf'):
        try:
            doc = fitz.open(fp)
            first_page_text = doc[0].get_text() if len(doc) > 0 else ""
            last_page_text = doc[-1].get_text() if len(doc) > 0 else ""
            
            # search all pages for doi
            for p in range(len(doc)):
                t = doc[p].get_text()
                for d in doi_pattern.findall(t):
                    d = d.rstrip('.,;:)')
                    if d not in all_dois and not d.endswith(('jpg','png','gif','svg')):
                        all_dois.append(d)
        except Exception as e:
            all_dois = [f"ERR: {e}"]
            
    findings.append({
        "filename": fn,
        "all_dois": all_dois[:5],
        "first_page_snippet": first_page_text[:400].replace('\n', ' ')
    })

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\remaining_analyzed.json", "w", encoding="utf-8") as out:
    json.dump(findings, out, ensure_ascii=False, indent=2)

print("Done scanning entire files.")
