import pandas as pd
import numpy as np
import os
import re

# File paths
BASE_DIR = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo"
INPUT_EXCEL = os.path.join(BASE_DIR, "00_Admin_Project_Status", "SOURCE_INVENTORY.xlsx")
OUT_DIR = os.path.join(BASE_DIR, "03_Literature_Screening")
os.makedirs(OUT_DIR, exist_ok=True)

OUT_MATRIX = os.path.join(OUT_DIR, "LITERATURE_SCREENING_MATRIX.xlsx")
OUT_CORE = os.path.join(OUT_DIR, "CORE_REFERENCE_LIST.md")
OUT_SEC = os.path.join(OUT_DIR, "SECONDARY_REFERENCE_LIST.md")
OUT_REJ = os.path.join(OUT_DIR, "REJECTED_REFERENCE_LIST.md")
OUT_CH_MAP = os.path.join(OUT_DIR, "CHAPTER_SOURCE_MAP.xlsx")
OUT_PR_MAP = os.path.join(OUT_DIR, "PRODUCT_SOURCE_MAP.xlsx")
OUT_LEGAL = os.path.join(OUT_DIR, "LEGAL_VALIDITY_CHECKLIST.md")
OUT_EVIDENCE = os.path.join(BASE_DIR, "00_Admin_Project_Status", "EVIDENCE_MATRIX.xlsx")

# Load data
df = pd.read_excel(INPUT_EXCEL)

# Initialize scoring columns
criteria = [
    'Relevance to thesis topic', 'Relevance to Chapter 1', 'Relevance to Chapter 2',
    'Relevance to Chapter 3', 'Relevance to Chapter 4', 'Theoretical value',
    'Legal/policy value', 'Empirical value', 'Methodological value', 'Recency',
    'Source authority', 'Usefulness for domestic articles', 'Usefulness for international article',
    'Usefulness for institutional project', 'Risk of weak/non-academic source'
]
for c in criteria:
    df[c] = 0

def safe_str(val):
    return str(val).lower() if pd.notna(val) else ""

# Keyword definitions
ch1_kw = ['tổng quan', 'review', 'nghiên cứu trước', 'overview', 'literature']
ch2_kw = ['lý thuyết', 'khái niệm', 'vai trò', 'đặc điểm', 'theory', 'concept', 'bản chất']
ch3_kw = ['thực trạng', 'đánh giá', 'số liệu', 'việt nam', 'status', 'current']
ch4_kw = ['giải pháp', 'định hướng', 'hoàn thiện', 'chính sách', 'solution', 'policy recommendation']

for idx, row in df.iterrows():
    title = safe_str(row.get('Document title'))
    concepts = safe_str(row.get('Main concepts'))
    src_type = safe_str(row.get('Source type'))
    rec_ch = safe_str(row.get('Recommended chapter'))
    rec_pr = safe_str(row.get('Recommended product'))
    usable_data = safe_str(row.get('Usable data'))
    year = row.get('Year')
    
    text_blob = f"{title} {concepts} {rec_ch} {src_type}".lower()
    
    # 1. Relevance to thesis topic
    topic_kw = ['sở hữu trí tuệ', 'kinh tế số', 'quản lý nhà nước', 'shtt', 'ip', 'digital economy', 'intellectual property']
    match_count = sum([1 for kw in topic_kw if kw in text_blob])
    df.at[idx, 'Relevance to thesis topic'] = min(5, match_count + (2 if 'shtt' in text_blob or 'sở hữu trí tuệ' in text_blob else 0))
    if df.at[idx, 'Relevance to thesis topic'] == 0: df.at[idx, 'Relevance to thesis topic'] = 1

    # 2. Relevance to Chapter 1
    if any(kw in text_blob for kw in ch1_kw) or '1' in rec_ch:
        df.at[idx, 'Relevance to Chapter 1'] = 4 if 'review' in text_blob else 3
    
    # 3. Relevance to Chapter 2
    if any(kw in text_blob for kw in ch2_kw) or '2' in rec_ch or 'book' in src_type or 'sách' in src_type:
        df.at[idx, 'Relevance to Chapter 2'] = 5 if 'book' in src_type else 4
    
    # 4. Relevance to Chapter 3
    if any(kw in text_blob for kw in ch3_kw) or '3' in rec_ch or 'thực trạng' in text_blob:
        df.at[idx, 'Relevance to Chapter 3'] = 5 if 'số liệu' in usable_data or 'data' in usable_data else 4
        
    # 5. Relevance to Chapter 4
    if any(kw in text_blob for kw in ch4_kw) or '4' in rec_ch:
        df.at[idx, 'Relevance to Chapter 4'] = 5 if 'giải pháp' in text_blob else 4

    # 6. Theoretical value
    if 'academic' in src_type or 'tạp chí' in src_type or 'journal' in src_type or 'book' in src_type:
        df.at[idx, 'Theoretical value'] = 5
    elif 'report' in src_type:
        df.at[idx, 'Theoretical value'] = 3
    else:
        df.at[idx, 'Theoretical value'] = 1

    # 7. Legal/policy value
    if 'legal' in src_type or 'văn bản' in src_type or 'luật' in src_type or 'nghị định' in src_type:
        df.at[idx, 'Legal/policy value'] = 5
    elif 'policy' in src_type or 'báo cáo' in src_type:
        df.at[idx, 'Legal/policy value'] = 4
    else:
        df.at[idx, 'Legal/policy value'] = 1

    # 8. Empirical value
    if pd.notna(usable_data) and len(str(usable_data)) > 5:
        df.at[idx, 'Empirical value'] = 5 if 'số liệu' in usable_data.lower() or '%' in usable_data else 3
    
    # 9. Methodological value
    if 'phương pháp' in text_blob or 'method' in text_blob or 'mô hình' in text_blob or 'model' in text_blob:
        df.at[idx, 'Methodological value'] = 5
    else:
        df.at[idx, 'Methodological value'] = 2

    # 10. Recency
    try:
        y = int(year)
        if y >= 2024: df.at[idx, 'Recency'] = 5
        elif y >= 2021: df.at[idx, 'Recency'] = 4
        elif y >= 2018: df.at[idx, 'Recency'] = 3
        elif y >= 2015: df.at[idx, 'Recency'] = 2
        else: df.at[idx, 'Recency'] = 1
    except:
        df.at[idx, 'Recency'] = 3 # Default if unknown

    # 11. Source authority
    if 'scopus' in text_blob or 'isi' in text_blob or 'wipo' in text_blob or 'world bank' in text_blob or 'chính phủ' in text_blob or 'bộ' in text_blob or 'quốc hội' in text_blob:
        df.at[idx, 'Source authority'] = 5
    elif 'tạp chí' in text_blob or 'journal' in text_blob:
        df.at[idx, 'Source authority'] = 4
    elif 'báo' in src_type and ('news' in src_type or 'thanh niên' in text_blob or 'vnexpress' in text_blob or 'tuổi trẻ' in text_blob):
        df.at[idx, 'Source authority'] = 1
    else:
        df.at[idx, 'Source authority'] = 3

    # 12. Usefulness for domestic articles
    if ('trong nước' in text_blob or 'việt nam' in text_blob) and df.at[idx, 'Theoretical value'] >= 4:
        df.at[idx, 'Usefulness for domestic articles'] = 5
    elif df.at[idx, 'Legal/policy value'] >= 4:
        df.at[idx, 'Usefulness for domestic articles'] = 4
    
    # 13. Usefulness for international article
    if 'english' in safe_str(row.get('Language')).lower() or 'scopus' in text_blob or 'quốc tế' in text_blob:
        df.at[idx, 'Usefulness for international article'] = 5
    elif df.at[idx, 'Theoretical value'] == 5 and df.at[idx, 'Recency'] >= 4:
        df.at[idx, 'Usefulness for international article'] = 3
        
    # 14. Usefulness for institutional project
    if df.at[idx, 'Relevance to Chapter 3'] >= 4 or df.at[idx, 'Relevance to Chapter 4'] >= 4 or df.at[idx, 'Legal/policy value'] >= 4:
        df.at[idx, 'Usefulness for institutional project'] = 5
        
    # 15. Risk of weak/non-academic source
    if df.at[idx, 'Source authority'] <= 2 and df.at[idx, 'Legal/policy value'] <= 2:
        df.at[idx, 'Risk of weak/non-academic source'] = 5
    elif pd.isna(row.get('Author/Agency')):
        df.at[idx, 'Risk of weak/non-academic source'] = 4
    else:
        df.at[idx, 'Risk of weak/non-academic source'] = 1

# Classification Logic
df['Total Score'] = df[criteria[:-1]].sum(axis=1) # Sum all except risk
df['Classification'] = 'Loại (Rejected)'
df['Reject Reason'] = ''

for idx, row in df.iterrows():
    if row['Risk of weak/non-academic source'] >= 4:
        df.at[idx, 'Classification'] = 'Loại (Rejected)'
        df.at[idx, 'Reject Reason'] = 'Nguồn không có tính học thuật / pháp lý rõ ràng hoặc rủi ro cao.'
    elif row['Total Score'] >= 35 and (row['Relevance to thesis topic'] >= 4 or row['Theoretical value'] == 5 or row['Legal/policy value'] == 5):
        df.at[idx, 'Classification'] = 'Lõi (Core)'
    elif row['Total Score'] >= 20:
        df.at[idx, 'Classification'] = 'Nền (Secondary)'
    else:
        df.at[idx, 'Classification'] = 'Loại (Rejected)'
        df.at[idx, 'Reject Reason'] = 'Điểm relevance và giá trị học thuật thấp.'

# Export Matrix
df.to_excel(OUT_MATRIX, index=False)

# Export Core
core_df = df[df['Classification'] == 'Lõi (Core)']
with open(OUT_CORE, 'w', encoding='utf-8') as f:
    f.write("# CORE REFERENCE LIST\n\n")
    for _, row in core_df.iterrows():
        f.write(f"- **{row['Document title']}** ({row['Year']}) - *{row['Author/Agency']}*\n")
        f.write(f"  - **Type:** {row['Source type']} | **Score:** {row['Total Score']}\n")
        f.write(f"  - **Main Concepts:** {row.get('Main concepts', 'N/A')}\n\n")

# Export Secondary
sec_df = df[df['Classification'] == 'Nền (Secondary)']
with open(OUT_SEC, 'w', encoding='utf-8') as f:
    f.write("# SECONDARY REFERENCE LIST\n\n")
    for _, row in sec_df.iterrows():
        f.write(f"- **{row['Document title']}** ({row['Year']}) - *{row['Author/Agency']}*\n")
        f.write(f"  - **Type:** {row['Source type']} | **Score:** {row['Total Score']}\n")

# Export Rejected
rej_df = df[df['Classification'] == 'Loại (Rejected)']
with open(OUT_REJ, 'w', encoding='utf-8') as f:
    f.write("# REJECTED REFERENCE LIST\n\n")
    for _, row in rej_df.iterrows():
        f.write(f"- **{row['Document title']}**\n")
        f.write(f"  - **Reason:** {row['Reject Reason']}\n\n")

# Export Chapter Map (Only Core and Sec)
ch_map = df[df['Classification'].isin(['Lõi (Core)', 'Nền (Secondary)'])][['Source ID', 'Document title', 'Relevance to Chapter 1', 'Relevance to Chapter 2', 'Relevance to Chapter 3', 'Relevance to Chapter 4']]
ch_map.to_excel(OUT_CH_MAP, index=False)

# Export Product Map (Only Core and Sec)
pr_map = df[df['Classification'].isin(['Lõi (Core)', 'Nền (Secondary)'])][['Source ID', 'Document title', 'Usefulness for domestic articles', 'Usefulness for international article', 'Usefulness for institutional project']]
pr_map.to_excel(OUT_PR_MAP, index=False)

# Export Legal Validity
legal_df = df[df['Legal/policy value'] >= 4]
with open(OUT_LEGAL, 'w', encoding='utf-8') as f:
    f.write("# LEGAL VALIDITY CHECKLIST\n\n")
    f.write("*(Cần kiểm tra hiệu lực, sửa đổi, thay thế của các văn bản pháp quy sau)*\n\n")
    f.write("| Tên văn bản | Năm | Cơ quan ban hành | Tình trạng hiệu lực |\n")
    f.write("|---|---|---|---|\n")
    for _, row in legal_df.iterrows():
        f.write(f"| {row['Document title']} | {row['Year']} | {row['Author/Agency']} | [ ] Chưa kiểm tra |\n")

# Update Evidence Matrix (Mock up an update by creating/updating the file)
try:
    with pd.ExcelWriter(OUT_EVIDENCE, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        core_df[['Source ID', 'Document title', 'Author/Agency', 'Year', 'Usable claims', 'Usable data']].to_excel(writer, sheet_name='Core_Evidence', index=False)
except Exception as e:
    # Fallback if file doesn't exist or is corrupted
    core_df[['Source ID', 'Document title', 'Author/Agency', 'Year', 'Usable claims', 'Usable data']].to_excel(OUT_EVIDENCE, index=False)

print("Screening completed successfully.")
