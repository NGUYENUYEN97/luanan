import json

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\all_records_summary.json", "r", encoding="utf-8") as f:
    items = json.load(f)

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\crossref_results.json", "r", encoding="utf-8") as f:
    cr = json.load(f)

# Let's inspect each item and print out its details cleanly to a json file
enhanced_items = []
for it in items:
    fn = it["filename"]
    doi = it["doi"]
    
    # Specific DOI mappings based on crossref and file inspection
    if fn == "28732.pdf" and not doi:
        doi = "10.36948/ijfmr.2024.v06i05.28732"
    elif fn == "Legal_Protection_of_Intellectual_Property_Rights_i.pdf" and not doi:
        doi = "10.58812/wslhr.v2i04.1155"
    elif fn == "Creative_Economy_as_a_Driver_of_Economic_Growth_in.pdf" and not doi:
        doi = "10.62872/t2w8cy65"
    elif fn == "Intellectual_Property_and_Economic_Development_Cat.pdf" and not doi:
        doi = "10.63053/ijrel.15"
    elif fn == "PROTECTION_OF_INTELLECTUAL_PROPERTY_RIGHTS_IN_THE_.pdf" and not doi:
        doi = "10.15407/econlaw.2021.04.079"
    elif fn == "Study_on_the_Two-way_Mutual_Promotion_Model_of_Dig.pdf" and not doi:
        doi = "10.54097/ajmss.v2i3.6828"
    elif fn == "AJMSS-3-2-132-136.pdf" and not doi:
        doi = "10.54097/ajmss.v3i2.8803"
    elif fn == "The_Evolution_of_Intellectual_Property_Rights_in_t.pdf" and not doi:
        doi = "10.53819/81018102t4162"
    elif fn == "ssrn-3923127.pdf" and not doi:
        doi = "10.2139/ssrn.3923127"
    elif fn == "ssrn-5147006.pdf" and not doi:
        doi = "10.2139/ssrn.5147006"
    elif fn in ["EBOOK-Intellectual-Property-in-the-Digital-Age.pdf", "Intellectual_property_in_the_digital_age.pdf"] and not doi:
        doi = "10.13134/979-12-5977-175-9"
    elif fn == "Research_on_Collaborative_Management_Mechanism_of_.pdf" and not doi:
        doi = "10.61360/gep.2024.020101"
    elif fn == "Research_on_the_Transformation_and_Development_Str.pdf" and not doi:
        doi = "10.54097/hbem.v39i.13456"
        
    enhanced_items.append({
        "index": it["index"],
        "filename": fn,
        "doi": doi,
        "has_doi": bool(doi)
    })

with open(r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\scratch\enhanced_dois.json", "w", encoding="utf-8") as f:
    json.dump(enhanced_items, f, ensure_ascii=False, indent=2)

print("Total items:", len(enhanced_items))
print("Items with DOI:", sum(1 for x in enhanced_items if x["has_doi"]))
print("Items without DOI:", [x["filename"] for x in enhanced_items if not x["has_doi"]])
