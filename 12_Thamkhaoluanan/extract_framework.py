import fitz
import os
import glob
import re

pdf_files = glob.glob("*.pdf")

keywords = [r"khung\s+lý\s+thuyết", r"khung\s+nghiên\s+cứu", r"khung\s+phân\s+tích", r"mô\s+hình\s+nghiên\s+cứu", r"mô\s+hình\s+lý\s+thuyết"]

with open("output.txt", "w", encoding="utf-8") as f:
    for pdf_file in pdf_files:
        f.write(f"=== File: {pdf_file} ===\n")
        try:
            doc = fitz.open(pdf_file)
            toc = doc.get_toc()
            f.write("--- Table of Contents (TOC) excerpts matching keywords ---\n")
            found_in_toc = False
            for item in toc:
                lvl, title, page = item
                title_lower = title.lower()
                if any(re.search(kw, title_lower) for kw in keywords):
                    f.write(f"Page {page}: {title}\n")
                    found_in_toc = True
            
            if not found_in_toc:
                f.write("No matching headings found in TOC.\n")
            
            f.write("--- Text excerpts ---\n")
            matches_found = 0
            for page_num in range(min(150, doc.page_count)):
                page = doc.load_page(page_num)
                text = page.get_text()
                for kw in keywords:
                    for match in re.finditer(kw, text.lower()):
                        if matches_found >= 10:
                            break
                        start = max(0, match.start() - 150)
                        end = min(len(text), match.end() + 150)
                        excerpt = text[start:end].replace('\n', ' ').strip()
                        f.write(f"Page {page_num+1}: ...{excerpt}...\n")
                        matches_found += 1
                    if matches_found >= 10:
                        break
                if matches_found >= 10:
                    break
            doc.close()
        except Exception as e:
            f.write(f"Error reading {pdf_file}: {e}\n")
        f.write("\n")
