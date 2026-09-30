import urllib.request
import json
import urllib.parse

def check_crossref(title):
    try:
        url = "https://api.crossref.org/works?query.bibliographic=" + urllib.parse.quote(title) + "&rows=3"
        req = urllib.request.Request(url, headers={'User-Agent': 'ResearchBot/1.0 (mailto:scholar@thesis.edu)'})
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            items = data.get('message', {}).get('items', [])
            res = []
            for it in items:
                res.append({
                    "title": it.get("title", [""])[0],
                    "doi": it.get("DOI", ""),
                    "score": it.get("score", 0)
                })
            return res
    except Exception as e:
        return [{"error": str(e)}]

queries = [
    "Legal Protection of Intellectual Property Rights in the Digital Industry: A Review of Legal Developments and Implementation Challenges",
    "Creative Economy as a Driver of Economic Growth in the Digital Era",
    "Protection of Intellectual Property Rights in the Global Processes of the Digital Economy Development",
    "Balancing Innovation and IP Protection in the Digital Economy: The Role of Copyright Law",
    "Intellectual Property and Economic Development: Catalysts for Innovation and Growth",
    "The Evolution of Intellectual Property Rights in the Digital Age"
]

results = {}
for q in queries:
    results[q] = check_crossref(q)

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\crossref_results.json", "w", encoding="utf-8") as f:
    json.dump(results, f, ensure_ascii=False, indent=2)

print("Saved crossref_results.json")
