import os
import re

base_dir = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\01_NotebookLM_Inputs\05_BaiBao_QuocTe"
html_files = [
    "9. Insight into IP Risk&Opportunity in the Era of Digital Economy-君伦律师事务所.html",
    "Intellectual property challenges for the digital economy.html",
    "Strengthening intellectual property policy in Vietnam’s digital economy.html"
]

out_lines = []
for hf in html_files:
    fp = os.path.join(base_dir, hf)
    with open(fp, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read()
        title_m = re.search(r'<title>(.*?)</title>', content, re.IGNORECASE | re.DOTALL)
        title = title_m.group(1).strip() if title_m else "NO_TITLE"
        url_m = re.search(r'<meta[^>]+property=["\']og:url["\'][^>]+content=["\']([^"\']+)["\']', content, re.IGNORECASE)
        if not url_m:
            url_m = re.search(r'<link[^>]+rel=["\']canonical["\'][^>]+href=["\']([^"\']+)["\']', content, re.IGNORECASE)
        url = url_m.group(1) if url_m else "NO_URL"
        
        auth_m = re.search(r'<meta[^>]+name=["\']author["\'][^>]+content=["\']([^"\']+)["\']', content, re.IGNORECASE)
        author = auth_m.group(1) if auth_m else "NO_AUTHOR"
        
        plain = re.sub(r'<[^>]+>', ' ', content)
        plain = ' '.join(plain.split())
        
        out_lines.append(f"FILE: {hf}")
        out_lines.append(f"TITLE: {title}")
        out_lines.append(f"URL: {url}")
        out_lines.append(f"AUTHOR: {author}")
        out_lines.append(f"SNIPPET: {plain[:500]}")
        out_lines.append("-" * 50)

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\html_info.txt", "w", encoding="utf-8") as out:
    out.write("\n".join(out_lines))

print("Wrote html_info.txt successfully")
