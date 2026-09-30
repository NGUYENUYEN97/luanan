import json
import re

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\metadata_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

print(f"Total entries: {len(data)}")
has_doi = 0
no_doi = []
for idx, item in enumerate(data):
    fn = item["filename"]
    dois = item.get("dois", [])
    if dois:
        has_doi += 1
        print(f"[{idx+1}] {fn} -> DOI: {dois[0]}")
    else:
        no_doi.append((idx+1, fn, item.get("sample_text", "")[:300].replace("\n", " ")))

print(f"\nTotal with DOI found directly: {has_doi} / {len(data)}")
print(f"Total without direct DOI: {len(no_doi)}")
print("\nItems without direct DOI:")
for idx, fn, sample in no_doi:
    print(f"--- #{idx} {fn} ---")
    print(f"Sample: {sample[:200]}")
