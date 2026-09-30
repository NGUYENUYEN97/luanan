import os
import pandas as pd
from datetime import datetime

# 1. Define paths
BASE_DIR = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\05_Data_Governance"
FOLDERS = [
    "raw_survey_data",
    "raw_interview_data",
    "cleaned_data",
    "coding_data",
    "analysis_outputs"
]

# Create directories
os.makedirs(BASE_DIR, exist_ok=True)
for folder in FOLDERS:
    os.makedirs(os.path.join(BASE_DIR, folder), exist_ok=True)

# 2. Create DATA_GOVERNANCE_LOG.md
log_content = """# DATA GOVERNANCE LOG

Tài liệu này ghi nhận mọi thao tác quản trị, làm sạch và biến đổi dữ liệu nhằm đảm bảo tính minh bạch và khả năng lặp lại (reproducibility) của nghiên cứu.

## Nguyên tắc cốt lõi:
1. **Bất khả xâm phạm dữ liệu thô:** Dữ liệu trong thư mục `raw_survey_data/` và `raw_interview_data/` là dữ liệu gốc, **tuyệt đối không được ghi đè hay chỉnh sửa trực tiếp**.
2. **Ghi chép minh bạch:** Mọi bước xử lý từ file raw sang file cleaned phải được ghi log tại đây.
3. **Ẩn danh:** Không ghi thông tin định danh chuyên gia vào bản thảo nếu chưa được phép, sử dụng mã CG01, CG02...
4. **Trích dẫn:** Mỗi bảng/hình tạo từ dữ liệu phải trích dẫn rõ file nguồn và biến được dùng.

---

## Lịch sử xử lý dữ liệu

| Ngày (YYYY-MM-DD) | Thao tác / Sự kiện | Người thực hiện | File Input | File Output | Ghi chú |
| :--- | :--- | :--- | :--- | :--- | :--- |
| """ + datetime.now().strftime("%Y-%m-%d") + """ | Khởi tạo Hệ thống Quản trị Dữ liệu Nghiên cứu | AI Assistant | - | Toàn bộ thư mục 05_Data_Governance | Đã thiết lập thành công. Chờ học viên nạp dữ liệu. |
"""

with open(os.path.join(BASE_DIR, "DATA_GOVERNANCE_LOG.md"), "w", encoding="utf-8") as f:
    f.write(log_content)

# 3. Create DATA_DICTIONARY.xlsx
data_dict_columns = ["Variable Code", "Variable Name", "Description", "Scale/Type", "Values/Coding", "Source/Reference"]
data_dict_data = [
    ["LF1", "Bảo vệ tài sản trí tuệ mới", "Đánh giá về tính hiệu quả của khung pháp lý đối với nội dung số, AI", "Ordinal (Likert 5)", "1: Hoàn toàn không đồng ý -> 5: Hoàn toàn đồng ý", "Survey"],
    ["LF2", "Chế tài răn đe", "Đánh giá mức độ răn đe của chế tài xử phạt trên môi trường mạng", "Ordinal (Likert 5)", "1: Hoàn toàn không đồng ý -> 5: Hoàn toàn đồng ý", "Survey"],
    ["-", "...", "...", "...", "...", "..."]
]
df_data_dict = pd.DataFrame(data_dict_data, columns=data_dict_columns)
df_data_dict.to_excel(os.path.join(BASE_DIR, "DATA_DICTIONARY.xlsx"), index=False)

# 4. Create INTERVIEW_CODEBOOK.xlsx
codebook_columns = ["Code (Mã chủ đề)", "Theme (Chủ đề lớn)", "Description (Mô tả nội dung)", "Example Quote (Ví dụ trích dẫn)"]
codebook_data = [
    ["LF_GAP", "Khoảng trống pháp lý", "Các ý kiến về sự thiếu hụt hoặc lạc hậu của luật pháp so với công nghệ", "[CG01]: Luật hiện nay chưa theo kịp sự phát triển của AI..."],
    ["OS_OVERLAP", "Chồng chéo quản lý", "Các ý kiến phàn nàn về sự đùn đẩy trách nhiệm giữa các bộ ngành", ""],
    ["-", "...", "...", ""]
]
df_codebook = pd.DataFrame(codebook_data, columns=codebook_columns)
df_codebook.to_excel(os.path.join(BASE_DIR, "INTERVIEW_CODEBOOK.xlsx"), index=False)

# 5. Create EXPERT_ANONYMIZATION_LIST.xlsx
expert_columns = ["Expert Code", "Real Name", "Position/Title", "Organization", "Date of Interview", "Consent Given", "Notes"]
expert_data = [
    ["CG01", "[Tên thật - Cần bảo mật]", "Phó Cục trưởng", "Cục SHTT", "2023-10-15", "Yes", "Đồng ý trích dẫn chức danh, không trích dẫn tên"],
    ["CG02", "[Tên thật - Cần bảo mật]", "Giám đốc Pháp chế", "Tập đoàn X", "2023-10-20", "Yes", "Yêu cầu ẩn danh hoàn toàn"],
    ["...", "...", "...", "...", "...", "...", "..."]
]
df_expert = pd.DataFrame(expert_data, columns=expert_columns)
df_expert.to_excel(os.path.join(BASE_DIR, "EXPERT_ANONYMIZATION_LIST.xlsx"), index=False)

print("Data Governance System initialized successfully.")
