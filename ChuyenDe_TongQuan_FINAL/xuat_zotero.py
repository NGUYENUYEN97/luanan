# Xuất danh mục tài liệu được trích dẫn trong chuyên đề sang định dạng nhập được vào Zotero.
# Chạy: python xuat_zotero.py   ->  ../Zotero_ChuyenDe_TongQuan.json (CSL JSON, khuyến nghị) và ../Zotero_ChuyenDe_TongQuan.ris
# Mỗi mục ghi khóa trong refs.py và số thứ tự [n] trong danh mục của bản .docx vào trường Ghi chú (note).
import json
import os
import re
import sys
import unicodedata

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
from refs import REFS  # noqa: E402

MD_FILES = ["01_MoDau.md", "02_Phan1_111.md", "03_Phan1_112.md", "04_Phan1_113.md", "05_Phan1_12_13.md", "07_KetLuan.md"]

# ------------------------------------------------------------------ đánh số giống build_chuyen_de.py
VN_ORDER = "aăâbcdđeêfghijklmnoôơpqrstuưvwxyz"
TONE_MARKS = {"̀", "́", "̃", "̉", "̣"}


def vn_key(s):
    out = []
    for ch in s.lower():
        dec = unicodedata.normalize("NFD", ch)
        base = unicodedata.normalize("NFC", "".join(c for c in dec if c not in TONE_MARKS))
        tone = sum(ord(c) for c in dec if c in TONE_MARKS)
        if base in VN_ORDER:
            out.append((VN_ORDER.index(base), tone))
        elif base.isdigit():
            out.append((-1, int(base)))
        else:
            out.append((-2, 0))
    return out


def read_md(name):
    out = []
    for line in open(os.path.join(HERE, name), encoding="utf-8").read().split("\n"):
        out.append(read_md(line[9:].strip()) if line.startswith("!INCLUDE ") else line)
    return "\n".join(out)


text = "\n".join(read_md(f) for f in MD_FILES)
cited = []
for m in re.finditer(r"\[@([^\]]+)\]", text):
    for k in m.group(1).split(";"):
        k = k.strip().lstrip("@")
        if k not in cited:
            cited.append(k)
vi = sorted([k for k in cited if REFS[k][0] == "vi"], key=lambda k: vn_key(REFS[k][1]))
en = sorted([k for k in cited if REFS[k][0] == "en"], key=lambda k: REFS[k][1].lower())
NUM = {k: i + 1 for i, k in enumerate(vi + en)}

# ------------------------------------------------------------------ tách chuỗi tài liệu thành trường
ORG_HINT = ("Organization", "Organisation", "OECD", "UNCTAD", "UNECE", "Fund", "Office", "Commission", "Group", "Quốc hội",
            "Chính phủ", "Thủ tướng", "Bộ ", "Cục ", "School", "World Bank", "Nations", "APEC")


def person(name, lang):
    name = name.strip()
    if any(h in name for h in ORG_HINT):
        return {"literal": name}
    parts = name.split()
    if lang == "vi" or not re.search(r"\b[A-Z]\.", name):
        # tên Việt Nam (có hoặc không dấu) hoặc tên viết đầy đủ: họ đứng đầu
        return {"family": parts[0], "given": " ".join(parts[1:])}
    # dạng "Goldfarb A." hoặc "Nicholson J.R."
    return {"family": " ".join(p for p in parts if not re.fullmatch(r"([A-Z]\.)+", p)),
            "given": " ".join(p for p in parts if re.fullmatch(r"([A-Z]\.)+", p))}


def parse(key):
    lang, _, s = REFS[key]
    item = {"id": key, "note": f"Khóa refs.py: {key}; số thứ tự trong danh mục chuyên đề: [{NUM[key]}]", "language": lang}
    m = re.match(r"^(.*?) \((\d{4})[a-z]?\), (.*)$", s)
    if not m:
        item.update(type="document", title=s)
        return item
    who, year, rest = m.groups()
    item["issued"] = {"date-parts": [[int(year)]]}
    eds = "(eds.)" in who or "(ed.)" in who
    who = who.replace("(eds.)", "").replace("(ed.)", "").strip()
    etal = who.endswith(" và cs")
    who = re.sub(r" và cs$", "", who)
    names = [person(n, lang) for n in re.split(r", (?=\S)", who) if n.strip()]
    item["editor" if eds else "author"] = names
    if etal:
        item["note"] += "; danh sách tác giả gốc có thêm người (và cs)"
    q = re.match(r"^“(.*?)”(?: \[(.*?)\])?, (.*)$", rest)
    if q:  # bài báo hoặc chương sách
        item["title"] = q.group(1)
        if q.group(2):
            item["note"] += f"; tên dịch: {q.group(2)}"
        tail = q.group(3)
        j = re.match(r"^(in .*?, )?\*(.*?)\*(.*)$", tail)
        if j:
            item["container-title"] = j.group(2)
            item["type"] = "chapter" if j.group(1) else "article-journal"
            after = j.group(3).strip(" ,.")
            vi_ = re.search(r"(\d+)\((\w+)\)", after) or re.search(r"^\((\w+)\)", after)
            if vi_ and vi_.re.pattern.startswith("(\\d"):
                item["volume"], item["issue"] = vi_.group(1), vi_.group(2)
            elif vi_:
                item["issue"] = vi_.group(1)
            else:
                v = re.match(r"^(\d+)\b", after)
                if v:
                    item["volume"] = v.group(1)
            pg = re.search(r"(\d+-\d+|\b\d{6}\b)\s*\.?$", after)
            if pg:
                item["page"] = pg.group(1)
            if "quanlynhanuoc.vn" in after:
                item["URL"] = "https://www.quanlynhanuoc.vn"
                item["type"] = "article-magazine"
        else:
            item.update(type="article", title=q.group(1))
        return item
    b = re.match(r"^\*(.*?)\*,? ?(.*)$", rest)
    title, tail = (b.group(1), b.group(2).strip(" .")) if b else (rest, "")
    item["title"] = title
    if "Luận án tiến sĩ" in tail:
        item["type"] = "thesis"
        item["genre"] = tail.split(",")[0]
        item["publisher"] = ", ".join(x.strip() for x in tail.split(",")[1:-1]) or tail
        item["publisher-place"] = tail.split(",")[-1].strip()
    elif any(w in tail for w in ("Working Paper", "Policy Paper", "Report for", "Technical Notes", "Discussion Paper", "Working Papers")):
        item["type"] = "report"
        item["publisher"] = tail
    elif any(n.get("literal", "").startswith(("Quốc hội", "Chính phủ", "Thủ tướng", "Bộ Chính trị")) for n in names):
        item["type"] = "legislation"
        item["publisher"] = tail
    else:
        item["type"] = "book"
        parts = [p.strip() for p in tail.split(",") if p.strip()]
        if len(parts) >= 2 and re.fullmatch(r"[A-Z]{2}", parts[-1]):
            parts = parts[:-2] + [parts[-2] + ", " + parts[-1]]
        if len(parts) >= 2:
            item["publisher"], item["publisher-place"] = ", ".join(parts[:-1]), parts[-1]
        elif parts:
            item["publisher"] = parts[0]
    return item


items = [parse(k) for k in vi + en]
json.dump(items, open(os.path.join(ROOT, "Zotero_ChuyenDe_TongQuan.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)

RIS_TY = {"article-journal": "JOUR", "article-magazine": "MGZN", "chapter": "CHAP", "thesis": "THES", "report": "RPRT",
          "legislation": "GOVDOC", "book": "BOOK", "article": "GEN", "document": "GEN"}
out = []
for it in items:
    r = [f"TY  - {RIS_TY[it['type']]}"]
    for role, tag in (("author", "AU"), ("editor", "A2")):
        for p in it.get(role, []):
            r.append(f"{tag}  - " + (p["literal"] if "literal" in p else f"{p['family']}, {p['given']}".rstrip(", ")))
    r.append(f"TI  - {it['title']}")
    if "issued" in it:
        r.append(f"PY  - {it['issued']['date-parts'][0][0]}")
    for k, tag in (("container-title", "T2"), ("volume", "VL"), ("issue", "IS"), ("publisher", "PB"), ("publisher-place", "CY"), ("URL", "UR")):
        if it.get(k):
            r.append(f"{tag}  - {it[k]}")
    if it.get("page"):
        sp, _, ep = it["page"].partition("-")
        r.append(f"SP  - {sp}")
        if ep:
            r.append(f"EP  - {ep}")
    r.append(f"LA  - {it['language']}")
    r.append(f"N1  - {it['note']}")
    r.append("ER  - ")
    out.append("\n".join(r))
open(os.path.join(ROOT, "Zotero_ChuyenDe_TongQuan.ris"), "w", encoding="utf-8").write("\n\n".join(out) + "\n")
print(f"Đã xuất {len(items)} tài liệu: Zotero_ChuyenDe_TongQuan.json, Zotero_ChuyenDe_TongQuan.ris")
