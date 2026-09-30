import json
import re

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\extracted_full_details.json", "r", encoding="utf-8") as f:
    items = json.load(f)

print(f"Total items: {len(items)}")

out_records = []
for it in items:
    fn = it["filename"]
    dois = it.get("dois", [])
    p_text = it.get("text_p1_p2", "")
    
    # Try to extract DOI if not already found
    # Look for doi.org or 10.xxxx
    doi = dois[0] if dois else ""
    if not doi:
        m = re.search(r'10\.\d{4,9}/[^\s"<>]+', p_text)
        if m:
            doi = m.group(0).rstrip('.,;:)')
            
    out_records.append({
        "index": it["index"],
        "filename": fn,
        "doi": doi,
        "doc_title": it.get("doc_title", ""),
        "doc_author": it.get("doc_author", ""),
        "text_sample": p_text[:600].replace('\n', ' ')
    })

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\all_records_summary.json", "w", encoding="utf-8") as out:
    json.dump(out_records, out, ensure_ascii=False, indent=2)

print("Saved all_records_summary.json")
