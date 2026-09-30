import os
import sys

sys.stdout.reconfigure(encoding='utf-8')

summary_path = r"C:\Users\Admin\.gemini\antigravity-ide\brain\2664ec95-b734-4052-ac21-ea79e5aed358\scratch\wipo_materials_summary.txt"

with open(summary_path, "r", encoding="utf-8") as f:
    content = f.read()

sections = content.split("==================================================")

for i, sec in enumerate(sections):
    lines = [l.strip() for l in sec.strip().split("\n") if l.strip()]
    if not lines:
        continue
    header = lines[0]
    print(f"\n--- SECTION {i+1}: {header} ---")
    # Print sample lines
    for line in lines[1:25]:
        print(f"  {line}")
    if len(lines) > 25:
        print(f"  ... (+ {len(lines)-25} more lines)")
