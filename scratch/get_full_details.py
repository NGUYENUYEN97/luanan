import json
import os
import fitz

base_dir = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\05_BaiBao_QuocTe"

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\metadata_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

extracted_full = []

for idx, item in enumerate(data):
    fn = item["filename"]
    fp = os.path.join(base_dir, fn)
    info = {"index": idx + 1, "filename": fn, "dois": item.get("dois", [])}
    
    if fn.endswith('.pdf'):
        try:
            doc = fitz.open(fp)
            meta = doc.metadata
            info["doc_title"] = meta.get("title", "")
            info["doc_author"] = meta.get("author", "")
            info["page_count"] = len(doc)
            first_p = doc[0].get_text() if len(doc) > 0 else ""
            second_p = doc[1].get_text() if len(doc) > 1 else ""
            info["text_p1_p2"] = (first_p + "\n" + second_p)[:1500]
        except Exception as e:
            info["error"] = str(e)
            
    elif fn.endswith('.html'):
        try:
            with open(fp, 'r', encoding='utf-8', errors='ignore') as hf:
                info["text_html"] = hf.read()[:1500]
        except Exception as e:
            info["error"] = str(e)
            
    extracted_full.append(info)

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\extracted_full_details.json", "w", encoding="utf-8") as out:
    json.dump(extracted_full, out, ensure_ascii=False, indent=2)

print("Saved extracted_full_details.json")
