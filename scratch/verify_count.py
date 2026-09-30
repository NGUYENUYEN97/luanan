import json

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\all_records_summary.json", "r", encoding="utf-8") as f:
    items = json.load(f)

print(f"Total items in all_records_summary: {len(items)}")
filenames = [it['filename'] for it in items]
print("Unique filenames count:", len(set(filenames)))
