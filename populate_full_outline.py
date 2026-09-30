import os
import pandas as pd
from docx import Document

BASE_DIR = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\07_Detailed_Outline"
os.makedirs(BASE_DIR, exist_ok=True)

# ==========================================
# 1. FULL EXCEL MATRICES
# ==========================================

# 1. EVIDENCE PLAN
evidence_columns = ["Chương", "Mục", "Luận điểm chính", "Evidence type", "Nguồn cụ thể"]
evidence_data = [
    ["Chương 1", "1.1. Tình hình nghiên cứu thế giới", "Các quốc gia đang chuyển đổi số quản trị SHTT.", "Literature", "WIPO, USPTO reports."],
    ["Chương 1", "1.2. Tình hình nghiên cứu tại VN", "Nghiên cứu trong nước tập trung pháp lý, thiếu công nghệ.", "Literature", "Các bài báo tạp chí Luật học VN."],
    ["Chương 1", "1.3. Khoảng trống nghiên cứu", "Chưa có khung phân tích định lượng về hiệu quả QLNN.", "Synthesis", "Đánh giá của tác giả."],
    ["Chương 2", "2.1. SHTT trong KT số", "KT số làm mờ biên giới, tài sản số dễ sao chép.", "Theory", "Sách chuyên khảo về kinh tế số."],
    ["Chương 2", "2.2. QLNN về SHTT", "Chuyển từ quản trị truyền thống sang Networked Governance.", "Theory", "Lý thuyết quản trị công mới."],
    ["Chương 2", "2.3. Kinh nghiệm quốc tế", "Mô hình TQ và HQ ứng dụng AI, liên thông dữ liệu.", "Case Study", "Báo cáo KIPO, CNIPA."],
    ["Chương 3", "3.1. Thực trạng pháp lý", "Luật SHTT 2022 chưa quy định rõ trách nhiệm nền tảng số.", "Legal", "Luật SHTT 2022, NĐ 17/2023."],
    ["Chương 3", "3.2. Bộ máy QLNN", "Chồng chéo giữa KHCN, TT&TT, Công an.", "Data/Interview", "Phỏng vấn CG01, CG02."],
    ["Chương 3", "3.3. Thực thi và Xử phạt", "Vi phạm gia tăng, tỷ lệ xử lý thấp.", "Data/Survey", "Biến survey LF1-LF4, Báo cáo Thanh tra."],
    ["Chương 4", "4.1. Quan điểm, định hướng", "Bảo vệ SHTT phải đồng hành khuyến khích sáng tạo.", "Policy", "NQ 52-NQ/TW."],
    ["Chương 4", "4.2. Giải pháp pháp lý", "Hoàn thiện cơ chế Safe Harbor và Notice-and-Takedown.", "Proposal", "Dựa trên Chương 3."],
    ["Chương 4", "4.3. Giải pháp công nghệ", "Xây dựng nền tảng giám sát vi phạm bằng AI.", "Proposal", "Dựa trên mô hình HQ/TQ."],
]
pd.DataFrame(evidence_data, columns=evidence_columns).to_excel(os.path.join(BASE_DIR, "OUTLINE_WITH_EVIDENCE_PLAN.xlsx"), index=False)

# 2. DATA USE MAP
data_map_columns = ["Biến khảo sát/Mã Phỏng vấn", "Sử dụng ở Chương/Mục", "Mục đích sử dụng"]
data_map_data = [
    ["LF1, LF2, LF3", "Mục 3.1", "Thống kê mô tả đánh giá mức độ răn đe của pháp luật hiện hành."],
    ["OS1, OS2, OS3", "Mục 3.2", "Đánh giá sự phối hợp liên ngành trong bộ máy QLNN."],
    ["TI1, TI2, TI3", "Mục 3.3", "Phân tích thực trạng ứng dụng công nghệ số trong QLNN."],
    ["CG01, CG02", "Mục 3.1, 3.2", "Giải thích nguyên nhân cốt lõi của sự đùn đẩy trách nhiệm."],
    ["CG03, CG04", "Mục 4.2", "Tham vấn tính khả thi của các quy định pháp luật mới."],
    ["Kết quả Hồi quy/SEM", "Tiểu kết Chương 3", "Xác định nhân tố tác động mạnh nhất đến Hiệu quả QLNN."],
]
pd.DataFrame(data_map_data, columns=data_map_columns).to_excel(os.path.join(BASE_DIR, "DATA_USE_MAP.xlsx"), index=False)

# 3. FIGURE TABLE REGISTER
fig_columns = ["Mã số", "Loại (Hình/Bảng)", "Tên", "Nguồn", "Vị trí trong luận án"]
fig_data = [
    ["Hình 1.1", "Hình", "Sơ đồ khung nghiên cứu", "Tác giả tổng hợp", "Cuối mục 1.3"],
    ["Hình 2.1", "Hình", "Mô hình quản trị mạng lưới trong SHTT", "Tác giả xây dựng dựa trên lý thuyết X", "Mục 2.2"],
    ["Bảng 3.1", "Bảng", "Thống kê số vụ vi phạm bản quyền trực tuyến (2020-2025)", "Thanh tra Bộ KHCN", "Mục 3.3"],
    ["Bảng 3.2", "Bảng", "Kết quả phân tích nhân tố (EFA)", "Dữ liệu khảo sát", "Cuối Chương 3"],
    ["Hình 4.1", "Hình", "Sơ đồ hệ thống liên thông dữ liệu SHTT quốc gia", "Tác giả đề xuất", "Mục 4.3"],
]
pd.DataFrame(fig_data, columns=fig_columns).to_excel(os.path.join(BASE_DIR, "FIGURE_TABLE_REGISTER.xlsx"), index=False)

# 4. LEGAL ARGUMENT MATRIX
legal_columns = ["Luận điểm trong luận án", "Văn bản pháp luật liên quan", "Điều khoản cụ thể", "Ghi chú/Đánh giá"]
legal_data = [
    ["Quy định nền tảng trung gian chưa rõ", "Luật SHTT sửa đổi 2022", "Điều 198b", "Mới chỉ quy định nguyên tắc chung, chưa có quy trình cụ thể gỡ bỏ."],
    ["Xử lý vi phạm còn nhẹ", "Nghị định 99/2013/NĐ-CP, NĐ 17/2023", "Chương xử phạt vi phạm hành chính", "Mức phạt tiền chưa đủ sức răn đe so với lợi nhuận thu được từ vi phạm online."],
    ["Khuyến khích số hoá", "Luật Công nghệ thông tin, NQ 52", "Toàn văn", "Cơ sở định hướng cho Chương 4."],
]
pd.DataFrame(legal_data, columns=legal_columns).to_excel(os.path.join(BASE_DIR, "LEGAL_ARGUMENT_MATRIX.xlsx"), index=False)

# 5. THESIS COHERENCE MAP
coherence_columns = ["Chương 1 (Vấn đề nghiên cứu)", "Chương 2 (Khung lý thuyết)", "Chương 3 (Bằng chứng thực tiễn)", "Chương 4 (Giải pháp tương ứng)"]
coherence_data = [
    ["Nghiên cứu trước chưa làm rõ trách nhiệm nền tảng số.", "Lý thuyết về trách nhiệm gián tiếp (Safe Harbor).", "Thực tiễn các vụ kiện Facebook, Shopee tại VN (Survey LF).", "Đề xuất quy trình Notice-and-Takedown chuẩn."],
    ["Thiếu công cụ đo lường hiệu quả quản lý bằng công nghệ.", "Mô hình quản lý công điện tử (E-government).", "Dữ liệu từ Bộ ngành rời rạc, chưa liên thông (Interview CG01).", "Xây dựng Hệ thống giám sát bằng AI dùng chung."],
]
pd.DataFrame(coherence_data, columns=coherence_columns).to_excel(os.path.join(BASE_DIR, "THESIS_COHERENCE_MAP.xlsx"), index=False)


# ==========================================
# 2. GENERATE FULL DETAILED WORD OUTLINE
# ==========================================

def add_outline_section(doc, level_text, title, objective, argument, content, refs, laws, survey, interview, tables, table_src, citation_pos, products, priority, risk, self_write):
    doc.add_heading(f"{level_text} {title}", level=2 if len(level_text.split('.'))==2 else 3)
    p = doc.add_paragraph()
    p.add_run("1. Mục tiêu: ").bold = True; p.add_run(objective + "\n")
    p.add_run("2. Luận điểm chính: ").bold = True; p.add_run(argument + "\n")
    p.add_run("3. Nội dung cần viết:\n").bold = True; p.add_run(content + "\n")
    p.add_run("4. Nguồn tài liệu tham khảo: ").bold = True; p.add_run(refs + "\n")
    p.add_run("5. Văn bản pháp quy: ").bold = True; p.add_run(laws + "\n")
    p.add_run("6. Dữ liệu khảo sát: ").bold = True; p.add_run(survey + "\n")
    p.add_run("7. Dữ liệu phỏng vấn: ").bold = True; p.add_run(interview + "\n")
    p.add_run("8. Bảng/hình dự kiến: ").bold = True; p.add_run(tables + "\n")
    p.add_run("9. Nguồn của bảng/hình: ").bold = True; p.add_run(table_src + "\n")
    p.add_run("10. Vị trí trích dẫn: ").bold = True; p.add_run(citation_pos + "\n")
    p.add_run("11. Sản phẩm liên quan: ").bold = True; p.add_run(products + "\n")
    p.add_run("12. Mức độ ưu tiên: ").bold = True; p.add_run(priority + "\n")
    p.add_run("13. Rủi ro: ").bold = True; p.add_run(risk + "\n")
    p.add_run("14. Nội dung học viên tự viết: ").bold = True; p.add_run(self_write + "\n")
    doc.add_paragraph("-" * 60)

doc = Document()
doc.add_heading('BẢN THIẾT KẾ THI CÔNG LUẬN ÁN (FULL BLUEPRINT)', 0)

# CHƯƠNG 1
doc.add_heading('CHƯƠNG 1: TỔNG QUAN TÌNH HÌNH NGHIÊN CỨU', level=1)

add_outline_section(doc, "1.1.", "Tình hình nghiên cứu trên thế giới", 
                    "Chỉ ra dòng chảy lý thuyết quốc tế về SHTT trong kỷ nguyên số.", 
                    "Thế giới đang chuyển từ quản lý cứng nhắc sang quản trị mạng lưới.", 
                    "- Điểm lại các nghiên cứu về bản quyền trên Internet.\n- Các nghiên cứu về nền tảng xuyên biên giới.",
                    "LITERATURE_SCREENING_MATRIX", "Không", "Không", "Không", "Bảng 1.1: Tổng hợp lý thuyết", "Tác giả tổng hợp", "Xuyên suốt", "Chuyên đề Tổng quan", "Cao", "Thiếu tài liệu mới", "Bình luận, so sánh các trường phái.")

add_outline_section(doc, "1.2.", "Tình hình nghiên cứu tại Việt Nam", 
                    "Làm rõ những gì học giả trong nước đã làm.", 
                    "Việt Nam chủ yếu nghiên cứu luật thực định, thiếu nghiên cứu thực nghiệm.", 
                    "- Nghiên cứu về Luật SHTT.\n- Nghiên cứu về vi phạm trên MXH.",
                    "Tạp chí Luật học, Nghiên cứu lập pháp", "Không", "Không", "Không", "Không", "Không", "Xuyên suốt", "Chuyên đề Tổng quan", "Cao", "Bỏ sót bài báo quan trọng", "Chỉ ra điểm hạn chế của các bài báo trước.")

add_outline_section(doc, "1.3.", "Khoảng trống nghiên cứu và mục tiêu, nhiệm vụ", 
                    "Khẳng định lý do vì sao luận án này cần thiết.", 
                    "Chưa có nghiên cứu hệ thống về hiệu quả QLNN đối với SHTT số bằng phương pháp hỗn hợp.", 
                    "- Khoảng trống lý thuyết.\n- Khoảng trống thực tiễn.\n- Mục tiêu tổng quát & cụ thể.",
                    "Tự suy luận từ 1.1 và 1.2", "Không", "Không", "Không", "Hình 1.1: Khung logic nghiên cứu", "Tác giả", "Cuối mục", "Phần mở đầu luận án", "Rất cao", "Lập luận không thuyết phục", "Xây dựng lập luận chặt chẽ bảo vệ tính mới.")

# CHƯƠNG 2
doc.add_heading('CHƯƠNG 2: CƠ SỞ LÝ LUẬN VÀ PHÁP LÝ VỀ QLNN ĐỐI VỚI QUYỀN SHTT TRONG KINH TẾ SỐ', level=1)

add_outline_section(doc, "2.1.", "Tổng quan về quyền SHTT trong nền kinh tế số", 
                    "Định hình đối tượng quản lý.", 
                    "Tài sản trí tuệ trong kinh tế số bị số hóa, vượt biên giới, khó kiểm soát.", 
                    "- Khái niệm kinh tế số.\n- Đặc điểm tài sản trí tuệ số.",
                    "Giáo trình SHTT, WIPO", "Luật SHTT", "Không", "Không", "Hình 2.1: Vòng đời TSTT số", "Tác giả", "Mục 2.1", "Bài báo trong nước 1", "Cao", "Định nghĩa lan man", "Đúc kết khái niệm dùng riêng cho luận án.")

add_outline_section(doc, "2.2.", "Nội dung Quản lý nhà nước về SHTT trong kinh tế số", 
                    "Cung cấp khung đo lường cho Chương 3.", 
                    "QLNN không chỉ là xử phạt mà là kiến tạo môi trường và hạ tầng số.", 
                    "- Ban hành pháp luật.\n- Tổ chức bộ máy.\n- Thanh tra, xử lý vi phạm.\n- Tuyên truyền, hợp tác.",
                    "Giáo trình QLNN", "Các quy định về chức năng bộ ngành", "Không", "Không", "Bảng 2.1: Khung tiêu chí đánh giá hiệu quả", "Tác giả", "Mục 2.2", "Chuyên đề 1", "Rất cao", "Lạc đề sang dân sự", "Chốt chuẩn các tiêu chí khảo sát.")

add_outline_section(doc, "2.3.", "Kinh nghiệm quốc tế và bài học cho Việt Nam", 
                    "Tìm kiếm mô hình mẫu.", 
                    "Trung Quốc và Hàn Quốc đã áp dụng thành công AI trong QLNN về SHTT.", 
                    "- Mô hình Hàn Quốc (KIPO).\n- Mô hình Trung Quốc (Tòa án Internet).\n- Bài học rút ra.",
                    "Báo cáo quốc tế", "Luật của HQ, TQ", "Không", "Không", "Bảng 2.2: So sánh kinh nghiệm quốc tế", "Tác giả", "Mục 2.3", "Chuyên đề 1", "Cao", "Thông tin cũ", "Tự rút ra bài học phù hợp với VN.")

# CHƯƠNG 3
doc.add_heading('CHƯƠNG 3: THỰC TRẠNG QLNN VỀ QUYỀN SHTT TRONG KINH TẾ SỐ Ở VIỆT NAM', level=1)

add_outline_section(doc, "3.1.", "Thực trạng khung pháp lý và chính sách", 
                    "Đánh giá luật hiện hành.", 
                    "Luật có nhiều tiến bộ nhưng chậm hơn công nghệ.", 
                    "- Quy định về TSTT mới.\n- Chế tài xử lý.",
                    "Báo cáo ngành", "Luật SHTT 2022", "Survey LF1-LF4", "CG01, CG02", "Bảng thống kê tỷ lệ hài lòng", "Data Analysis", "Phân tán", "Bài báo 1, Chuyên đề 2", "Rất cao", "Số liệu survey không đạt", "Bình luận các bất cập từ số liệu.")

add_outline_section(doc, "3.2.", "Thực trạng tổ chức bộ máy và nguồn nhân lực", 
                    "Chỉ ra điểm nghẽn về con người.", 
                    "Thiếu nhân lực công nghệ cao tại cơ quan QLNN, phân quyền chồng chéo.", 
                    "- Cơ cấu tổ chức Bộ KHCN, TT&TT.\n- Năng lực cán bộ.",
                    "Báo cáo Bộ", "Quy định chức năng", "Survey OS, HR", "CG03, CG04", "Biểu đồ đánh giá năng lực", "Data Analysis", "Phân tán", "Bài báo 2", "Cao", "Khó lấy số liệu", "Tổng hợp nhận định từ phỏng vấn sâu.")

add_outline_section(doc, "3.3.", "Thực trạng thanh tra, kiểm tra và xử lý vi phạm", 
                    "Đánh giá đầu ra của QLNN.", 
                    "Vi phạm tràn lan, xử lý thủ công, thiếu công cụ kỹ thuật.", 
                    "- Các vụ việc nổi cộm trên mạng.\n- Hoạt động của nền tảng trung gian.",
                    "Báo cáo Thanh tra", "Nghị định xử phạt", "Survey TI, ME", "CG05", "Bảng 3.1: Số vụ xử phạt", "Bộ KHCN", "Phân tán", "Bài báo quốc tế", "Rất cao", "Số liệu không minh bạch", "Chạy mô hình phân tích định lượng (Regression/SEM).")

# CHƯƠNG 4
doc.add_heading('CHƯƠNG 4: QUAN ĐIỂM VÀ GIẢI PHÁP HOÀN THIỆN QLNN VỀ SHTT', level=1)

add_outline_section(doc, "4.1.", "Quan điểm, định hướng hoàn thiện", 
                    "Dẫn đường cho giải pháp.", 
                    "Chuyển từ tiền kiểm sang hậu kiểm bằng công nghệ.", 
                    "- Bối cảnh mới (AI, Blockchain).\n- Quan điểm của Đảng và Nhà nước.",
                    "Nghị quyết TW", "Các văn kiện Đảng", "Không", "Không", "Không", "Không", "Đầu chương", "Đề tài cấp cơ sở", "Trung bình", "Lan man", "Viết súc tích, bám sát chiến lược quốc gia.")

add_outline_section(doc, "4.2.", "Giải pháp hoàn thiện khung pháp lý", 
                    "Lấp đầy khoảng trống ở Chương 3.", 
                    "Cần ban hành Nghị định riêng về SHTT trên môi trường mạng.", 
                    "- Hoàn thiện quy định về AI/Big Data.\n- Sửa đổi mức phạt.",
                    "Tham khảo Chương 2", "Đề xuất sửa luật", "Không", "CG01, CG02", "Không", "Không", "Mục 4.2", "Bài báo 3", "Rất cao", "Đề xuất thiếu thực tế", "Sử dụng phỏng vấn chuyên gia để bảo vệ tính khả thi.")

add_outline_section(doc, "4.3.", "Giải pháp đổi mới tổ chức bộ máy và ứng dụng công nghệ", 
                    "Giải quyết vấn đề vận hành.", 
                    "Thành lập trung tâm dữ liệu SHTT quốc gia dùng chung.", 
                    "- Cơ chế phối hợp liên ngành.\n- Giải pháp hạ tầng công nghệ số.",
                    "Báo cáo CĐS quốc gia", "Quyết định của Thủ tướng", "Không", "CG03, CG04", "Hình 4.1: Hệ thống liên thông", "Tác giả đề xuất", "Mục 4.3", "Bài báo 2, Chuyên đề 3", "Rất cao", "Vượt quá nguồn lực quốc gia", "Thiết kế phân kỳ đầu tư, lộ trình áp dụng.")

doc.save(os.path.join(BASE_DIR, "DETAILED_THESIS_OUTLINE.docx"))
print("Tạo bản thiết kế thi công FULL thành công.")
