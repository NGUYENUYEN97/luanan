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
    item = {"filename": f, "type": "pdf" if f.endswith('.pdf') else "html" if f.endswith('.html') else "other"}
    text_sample = ""
    dois = []
    
    if f.endswith('.pdf'):
        try:
            doc = fitz.open(filepath)
            num_pages = len(doc)
            item["pages"] = num_pages
            # Read first 2 pages for metadata & DOI
            for p in range(min(num_pages, 3)):
                page_text = doc[p].get_text()
                text_sample += page_text + "\n"
            
            # Find DOIs
            found_dois = doi_pattern.findall(text_sample)
            # clean DOIs (remove trailing punctuation)
            clean_dois = []
            for d in found_dois:
                d = d.rstrip('.,;:)')
                if d not in clean_dois and not d.endswith(('jpg', 'png', 'gif')):
                    clean_dois.append(d)
            item["dois"] = clean_dois
            item["first_500_chars"] = text_sample[:500].replace('\n', ' ')
        except Exception as e:
            item["error"] = str(e)
            
    elif f.endswith('.html'):
        try:
            with open(filepath, 'r', encoding='utf-8', errors='ignore') as hf:
                html_text = hf.read()
                # strip basic html tags
                plain_text = re.sub(r'<[^>]+>', ' ', html_text)
                plain_text = ' '.join(plain_text.split())
                found_dois = doi_pattern.findall(plain_text)
                item["dois"] = list(set([d.rstrip('.,;:)') for d in found_dois]))
                item["first_500_chars"] = plain_text[:500]
        except Exception as e:
            item["error"] = str(e)
            
    results.append(item)

print(f"Total processed: {len(results)}")
for r in results:
    dois_str = ", ".join(r.get("dois", [])) if r.get("dois") else "NO_DOI"
    print(f"[{r['filename']}] -> DOI: {dois_str}")

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\extracted_dois.json", "w", encoding="utf-8") as out:
    json.dump(results, out, ensure_ascii=False, indent=2)
