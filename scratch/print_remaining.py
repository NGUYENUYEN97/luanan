import json

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\remaining_analyzed.json", "r", encoding="utf-8") as f:
    items = json.load(f)

for idx, it in enumerate(items):
    fn = it["filename"]
    dois = it.get("all_dois", [])
    snip = it.get("first_page_snippet", "")[:120]
    print(f"[{idx+1}] {fn} | DOIs: {dois} | Head: {snip}")
