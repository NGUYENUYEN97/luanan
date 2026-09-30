import os
import re
import fitz # PyMuPDF
import json

base_dir = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\05_BaiBao_QuocTe"
files = [f for f in os.listdir(base_dir) if os.path.isfile(os.path.join(base_dir, f))]

doi_pattern = re.compile(r'10\.\d{4,9}/[-._;()/:A-Za-z0-9]+', re.IGNORECASE)

results = []

for f in sorted(files):
    if f.endswith('.csv') or f.endswith('.txt'):
        continue
    filepath = os.path.join(base_dir, f)
    item = {
        "filename": f,
        "type": "pdf" if f.endswith('.pdf') else "html" if f.endswith('.html') else "other",
        "dois": [],
        "sample_text": ""
    }
    
    if f.endswith('.pdf'):
        try:
            doc = fitz.open(filepath)
            num_pages = len(doc)
            item["pages"] = num_pages
            full_head = ""
            for p in range(min(num_pages, 4)):
                full_head += doc[p].get_text() + "\n"
            
            found_dois = doi_pattern.findall(full_head)
            clean_dois = []
            for d in found_dois:
                d = d.rstrip('.,;:)')
                if d not in clean_dois and not d.endswith(('jpg', 'png', 'gif', 'svg')):
                    clean_dois.append(d)
            item["dois"] = clean_dois
            item["sample_text"] = full_head[:2500]
        except Exception as e:
            item["error"] = str(e)
            
    elif f.endswith('.html'):
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as hf:
                html_text = hf.read()
                plain_text = re.sub(r'<[^>]+>', ' ', html_text)
                plain_text = ' '.join(plain_text.split())
                found_dois = doi_pattern.findall(plain_text)
                item["dois"] = list(set([d.rstrip('.,;:)') for d in found_dois]))
                item["sample_text"] = plain_text[:2500]
        except Exception as e:
            item["error"] = str(e)
            
    results.append(item)

out_path = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\metadata_all.json"
with open(out_path, "w", encoding="utf-8") as out:
    json.dump(results, out, ensure_ascii=False, indent=2)

print(f"Successfully processed {len(results)} files and wrote to {out_path}")
