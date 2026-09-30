import os
import re

base_dir = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\05_BaiBao_QuocTe"

lines = []
for hf in [
    "9. Insight into IP Risk&Opportunity in the Era of Digital Economy-君伦律师事务所.html",
    "Intellectual property challenges for the digital economy.html",
    "Strengthening intellectual property policy in Vietnam’s digital economy.html"
]:
    fp = os.path.join(base_dir, hf)
    with open(fp, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
        content = re.sub(r'<script.*?</script>', '', content, flags=re.DOTALL | re.IGNORECASE)
        content = re.sub(r'<style.*?</style>', '', content, flags=re.DOTALL | re.IGNORECASE)
        plain = re.sub(r'<[^>]+>', ' ', content)
        plain = ' '.join(plain.split())
        lines.append(f"=== {hf} ===")
        lines.append(plain[:1500])
        lines.append("\n" + "="*50 + "\n")

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\clean_html.txt", "w", encoding="utf-8") as out:
    out.write("\n".join(lines))

print("Wrote clean_html.txt")
