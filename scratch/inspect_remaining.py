import json
import re

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\metadata_all.json", "r", encoding="utf-8") as f:
    data = json.load(f)

remaining = []
for idx, item in enumerate(data):
    if not item.get("dois"):
        remaining.append({
            "index": idx + 1,
            "filename": item["filename"],
            "type": item["type"],
            "sample_head": item.get("sample_text", "")[:1000]
        })

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\remaining.json", "w", encoding="utf-8") as out:
    json.dump(remaining, out, ensure_ascii=False, indent=2)

print(f"Wrote {len(remaining)} items to remaining.json")
