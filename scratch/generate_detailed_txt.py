import json
import re

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\all_records_summary.json", "r", encoding="utf-8") as f:
    items = json.load(f)

lines = []
for it in items:
    lines.append(f"==================================================")
    lines.append(f"INDEX: {it['index']}")
    lines.append(f"FILENAME: {it['filename']}")
    lines.append(f"DOI: {it['doi']}")
    lines.append(f"DOC TITLE: {it['doc_title']}")
    lines.append(f"DOC AUTHOR: {it['doc_author']}")
    lines.append(f"SNIPPET: {it['text_sample']}")
    lines.append("")

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\all_57_papers_detailed.txt", "w", encoding="utf-8") as out:
    out.write("\n".join(lines))

print("Wrote all_57_papers_detailed.txt")
