import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document()

# Margin configuration: Top/Bottom/Right = 2cm (0.79 in), Left = 3cm (1.18 in)
sections = doc.sections
for section in sections:
    section.top_margin = Inches(0.79)
    section.bottom_margin = Inches(0.79)
    section.left_margin = Inches(1.18)
    section.right_margin = Inches(0.79)

style_normal = doc.styles['Normal']
style_normal.font.name = 'Times New Roman'
style_normal.font.size = Pt(13)
style_normal.font.color.rgb = RGBColor(0, 0, 0)
style_normal.paragraph_format.line_spacing = 1.3
style_normal.paragraph_format.space_after = Pt(4)

def add_title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(15)
    run.font.bold = True
    return p

def add_h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.font.bold = True
    return p

def add_h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.font.bold = True
    return p

def add_h3(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.italic = True
    return p

def add_p(text, bold_prefix='', italic=False):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.first_line_indent = Inches(0.4)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Times New Roman'
        r_pre.font.size = Pt(13)
        r_pre.font.bold = True
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    run.font.italic = italic
    return p

def add_bullet(text, bold_prefix=''):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.line_spacing = 1.3
    p.paragraph_format.space_after = Pt(3)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.name = 'Times New Roman'
        r_pre.font.size = Pt(13)
        r_pre.font.bold = True
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(13)
    return p

# TITLE SECTION
add_title('ĐỀ CƯƠNG LUẬN ÁN TIẾN SĨ\nNGÀNH QUẢN LÝ KINH TẾ')
add_title('TÊN ĐỀ TÀI:\nQUẢN LÝ NHÀ NƯỚC VỀ QUYỀN SỞ HỮU TRÍ TUỆ TRONG NỀN KINH TẾ SỐ Ở VIỆT NAM')

# PHẦN MỞ ĐẦU
add_h1('PHẦN MỞ ĐẦU')

# 1. TÍNH CẤP THIẾT
add_h2('1. Tính cấp thiết của đề tài')
add_p('Quyền sở hữu trí tuệ đóng vai trò là công cụ vĩ mô quan trọng nhằm tạo lập môi trường kinh doanh minh bạch, bảo vệ quyền và lợi ích hợp pháp của các chủ thể, bảo đảm trật tự thị trường, cân bằng lợi ích của chủ thể quyền và xã hội, thực hiện các cam kết quốc tế, hỗ trợ phổ biến tri thức, đồng thời thúc đẩy cạnh tranh và chuyển giao công nghệ. Nhận thức rõ tầm quan trọng này, Đảng và Nhà nước ta đã ban hành nhiều chủ trương, định hướng chiến lược. Văn kiện Đại hội XIII của Đảng nhấn mạnh mô hình tăng trưởng dựa trên khoa học công nghệ, đổi mới sáng tạo và chuyển đổi số. Nghị quyết số 52-NQ/TW của Bộ Chính trị xác định chủ động tham gia cuộc Cách mạng công nghiệp lần thứ tư; Nghị quyết số 57-NQ/TW của Bộ Chính trị tạo bước đột phá về phát triển khoa học, công nghệ, đổi mới sáng tạo và chuyển đổi số quốc gia. Cụ thể hóa các định hướng lớn này, Thủ tướng Chính phủ đã ban hành Chiến lược Sở hữu trí tuệ đến năm 2030 (Quyết định số 1068/QĐ-TTg năm 2019), Chiến lược quốc gia phát triển kinh tế số và xã hội số đến năm 2025, định hướng đến năm 2030 (Quyết định số 411/QĐ-TTg) và Chương trình Chuyển đổi số quốc gia đến năm 2025, định hướng đến năm 2030 (Quyết định số 749/QĐ-TTg năm 2020).')

add_p('Trong bối cảnh nền kinh tế số phát triển mạnh mẽ, các giao dịch kinh doanh chuyển dịch nhanh chóng lên môi trường số. Điều này đặt ra những yêu cầu mới đối với quản lý nhà nước về quyền sở hữu trí tuệ, đặc biệt là trong việc hỗ trợ doanh nghiệp tiếp cận và sử dụng hiệu quả hệ thống sở hữu trí tuệ để bảo vệ kết quả đổi mới sáng tạo, nâng cao năng lực cạnh tranh. Luận án tiếp cận vấn đề dưới góc độ Quản lý kinh tế, tập trung xem xét vai trò của quản lý nhà nước về quyền sở hữu trí tuệ đối với đổi mới sáng tạo của doanh nghiệp.')

add_p('Về lý luận, các nghiên cứu về quản lý nhà nước đối với quyền sở hữu trí tuệ chủ yếu tập trung vào thể chế, tổ chức thực hiện và bảo vệ quyền, trong khi các nghiên cứu về đổi mới sáng tạo thường tiếp cận từ nguồn lực nội bộ của doanh nghiệp. Mối quan hệ giữa quản lý nhà nước, mức độ doanh nghiệp tiếp cận và sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo trong điều kiện kinh tế số chưa được làm rõ đầy đủ. Đây là khoảng trống lý luận mà luận án tập trung nghiên cứu.')

add_p('Tuy nhiên, thực tiễn công tác quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số ở Việt Nam hiện nay đang bộc lộ nhiều hạn chế: khuôn khổ pháp lý chưa hoàn toàn thích ứng với các dịch vụ số; công tác tổ chức thực hiện và hỗ trợ xác lập, thương mại hóa quyền chưa đạt hiệu quả cao; hoạt động thực thi và xử lý vi phạm trên môi trường số còn gặp nhiều khó khăn. Xuất phát từ những yêu cầu lý luận và thực tiễn đó, nghiên cứu sinh lựa chọn triển khai đề tài: "Quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số ở Việt Nam" làm đề tài luận án Tiến sĩ chuyên ngành Quản lý kinh tế.')

# 2. MỤC ĐÍCH VÀ NHIỆM VỤ NGHIÊN CỨU
add_h2('2. Mục đích và nhiệm vụ nghiên cứu')
add_h3('2.1. Mục đích nghiên cứu')
add_p('Mục đích của luận án là làm rõ cơ sở lý luận và thực tiễn của quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số; đánh giá thực trạng ở Việt Nam giai đoạn 2019–2025; làm rõ vai trò và mối quan hệ giữa chất lượng quản lý nhà nước theo đánh giá của doanh nghiệp, mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo của doanh nghiệp; từ đó đề xuất hệ thống giải pháp hoàn thiện đến năm 2035.')

add_h3('2.2. Nhiệm vụ nghiên cứu')
add_bullet('Hệ thống hóa cơ sở lý luận, nội dung và tiêu chí đánh giá quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số; làm rõ mối quan hệ giữa chất lượng quản lý nhà nước, mức độ doanh nghiệp tiếp cận và sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo của doanh nghiệp.', bold_prefix='- Thứ nhất: ')
add_bullet('Phân tích kinh nghiệm quốc tế về quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số và rút ra các bài học kinh nghiệm cho Việt Nam.', bold_prefix='- Thứ hai: ')
add_bullet('Phân tích, đánh giá thực trạng quản lý nhà nước về quyền sở hữu trí tuệ ở Việt Nam giai đoạn 2019–2025 theo bốn nội dung quản lý và tiêu chí đánh giá; kiểm định mối quan hệ giữa chất lượng quản lý nhà nước theo đánh giá của doanh nghiệp, mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo của doanh nghiệp.', bold_prefix='- Thứ ba: ')
add_bullet('Đề xuất quan điểm, định hướng và hệ thống giải pháp hoàn thiện quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số ở Việt Nam đến năm 2035.', bold_prefix='- Thứ tư: ')

# 3. CÂU HỎI NGHIÊN CỨU
add_h2('3. Câu hỏi nghiên cứu')
add_p('Để thực hiện các mục tiêu nghiên cứu, luận án tập trung giải quyết 04 câu hỏi nghiên cứu trọng tâm sau:')
add_bullet('Quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số gồm những nội dung nào và được đánh giá theo những tiêu chí nào?', bold_prefix='- Câu hỏi 1: ')
add_bullet('Thực trạng quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số ở Việt Nam giai đoạn 2019–2025 như thế nào?', bold_prefix='- Câu hỏi 2: ')
add_bullet('Chất lượng quản lý nhà nước về quyền sở hữu trí tuệ theo đánh giá của doanh nghiệp có mối quan hệ như thế nào với mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo của doanh nghiệp?', bold_prefix='- Câu hỏi 3: ')
add_bullet('Cần thực hiện những giải pháp nào để hoàn thiện quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số ở Việt Nam đến năm 2035?', bold_prefix='- Câu hỏi 4: ')

# 4. ĐỐI TƯỢNG VÀ PHẠM VI NGHIÊN CỨU
add_h2('4. Đối tượng và phạm vi nghiên cứu')
add_h3('4.1. Đối tượng nghiên cứu và đối tượng khảo sát')
add_p('Đối tượng nghiên cứu của luận án là quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số ở Việt Nam.')

add_p('Đối tượng khảo sát được thiết kế tách biệt theo hai phương pháp nghiên cứu:')
add_bullet('Đối tượng phục vụ nghiên cứu định lượng (khảo sát bằng bảng hỏi): Các doanh nghiệp hoạt động trong những ngành có mức độ số hóa và sử dụng tài sản trí tuệ tương đối cao tại Việt Nam. Mẫu khảo sát bao gồm cả doanh nghiệp đã và chưa sử dụng hệ thống sở hữu trí tuệ, đã và chưa thực hiện đổi mới sáng tạo trong thời kỳ nghiên cứu. Đơn vị phân tích là doanh nghiệp; người trả lời là lãnh đạo hoặc nhân sự am hiểu hoạt động đổi mới sáng tạo và sở hữu trí tuệ của doanh nghiệp.', bold_prefix='- ')
add_bullet('Đối tượng phục vụ nghiên cứu định tính (phỏng vấn sâu): Cán bộ quản lý nhà nước về sở hữu trí tuệ ở trung ương và địa phương, các chuyên gia pháp lý, đại diện tổ chức sở hữu công nghiệp, đại diện các tổ chức trung gian (viện nghiên cứu, trường đại học, sàn giao dịch công nghệ) và một số doanh nghiệp tiêu biểu.', bold_prefix='- ')

add_h3('4.2. Phạm vi nghiên cứu')
add_p('Phạm vi về nội dung: Luận án nghiên cứu bốn nội dung quản lý nhà nước chủ yếu: (1) Xây dựng chiến lược, chính sách và pháp luật; (2) Tổ chức bộ máy và phối hợp quản lý; (3) Tổ chức thực hiện và hỗ trợ xác lập, khai thác, chuyển giao, thương mại hóa quyền; (4) Kiểm tra, giám sát, thực thi và xử lý vi phạm. Chuyển đổi số, nhân lực, tài chính và hạ tầng dữ liệu được xem xét như các điều kiện bảo đảm hoạt động quản lý. Luận án tập trung vào nhóm quyền sở hữu trí tuệ nghiên cứu chính (gồm sáng chế liên quan đến công nghệ số, nhãn hiệu được sử dụng trong hoạt động kinh doanh số, quyền tác giả đối với phần mềm và nội dung số) và nhóm nghiên cứu bổ trợ (gồm bí mật kinh doanh; dữ liệu và sản phẩm do trí tuệ nhân tạo tạo ra được xem xét như những vấn đề mới đặt ra đối với quản lý nhà nước). Luận án đồng thời phân tích mối quan hệ giữa chất lượng quản lý nhà nước theo đánh giá của doanh nghiệp, mức độ doanh nghiệp tiếp cận và sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo của doanh nghiệp.', bold_prefix='- ')
add_p('Phạm vi về thời gian: Dữ liệu thứ cấp được sử dụng để phân tích thực trạng quản lý nhà nước trong giai đoạn 2019–2025 (bao quát từ khi ban hành Chiến lược SHTT quốc gia 2019 và triển khai Chương trình Chuyển đổi số quốc gia 2020 theo Quyết định 749/QĐ-TTg). Dữ liệu sơ cấp được thu thập trong thời gian thực hiện luận án; kết quả đổi mới sáng tạo của doanh nghiệp được khảo sát trong ba năm gần nhất trước thời điểm điều tra. Định hướng và giải pháp được đề xuất đến năm 2035.', bold_prefix='- ')
add_p('Phạm vi về không gian: Nghiên cứu thực tiễn trên phạm vi cả nước Việt Nam và tham khảo kinh nghiệm của một số quốc gia tiêu biểu.', bold_prefix='- ')

# 5. PHƯƠNG PHÁP NGHIÊN CỨU
add_h2('5. Cơ sở phương pháp luận và phương pháp nghiên cứu')
add_p('Luận án vận dụng phương pháp luận duy vật biện chứng và duy vật lịch sử, kết hợp các tiếp cận thể chế, tiếp cận hệ thống và tiếp cận hệ thống đổi mới quốc gia.')

add_p('Luận án sử dụng phương pháp nghiên cứu hỗn hợp, trong đó nghiên cứu định tính giữ vai trò chủ đạo; nghiên cứu định lượng được sử dụng để bổ sung bằng chứng và kiểm tra các mối quan hệ trong khung phân tích:')
add_bullet('Nghiên cứu định tính: Được thực hiện thông qua tổng quan tài liệu, phân tích chính sách và văn bản pháp luật, phân tích số liệu thứ cấp, so sánh kinh nghiệm quốc tế và phỏng vấn sâu bán cấu trúc. Đối tượng phỏng vấn dự kiến gồm 20–25 cán bộ quản lý nhà nước, chuyên gia, đại diện tổ chức sở hữu công nghiệp, tổ chức trung gian và doanh nghiệp. Kết quả nghiên cứu định tính được sử dụng để đánh giá thực trạng, xác định hạn chế và nguyên nhân, hoàn thiện bảng hỏi và giải thích kết quả khảo sát.', bold_prefix='- ')
add_bullet('Nghiên cứu định lượng: Được thực hiện thông qua khảo sát khoảng 300–350 doanh nghiệp thuộc các ngành có mức độ số hóa và sử dụng tài sản trí tuệ tương đối cao. Người trả lời là lãnh đạo hoặc nhân sự am hiểu hoạt động sở hữu trí tuệ và đổi mới sáng tạo của doanh nghiệp. Thang đo Likert 5 mức được sử dụng để thu thập đánh giá của doanh nghiệp về chất lượng chính sách, thủ tục, dịch vụ hỗ trợ và thực thi quyền. Các chỉ báo về mức độ tiếp cận, sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo được xây dựng từ tổng quan tài liệu, hướng dẫn của OECD Oslo Manual, phỏng vấn chuyên gia và khảo sát thử. Các thang đo được thiết kế để phân biệt rõ giữa đánh giá của doanh nghiệp về chất lượng quản lý nhà nước và mức độ doanh nghiệp thực tế tiếp cận, sử dụng các thành tố của hệ thống sở hữu trí tuệ. Dữ liệu khảo sát được phân tích bằng thống kê mô tả, kiểm định độ tin cậy của thang đo (Cronbach’s Alpha), phân tích nhân tố khám phá (EFA), phân tích tương quan và hồi quy đa biến. Vai trò trung gian của mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ được kiểm định bằng phương pháp bootstrap. Quy mô, ngành, tuổi doanh nghiệp, hình thức sở hữu, mức độ số hóa và hoạt động R&D được xem xét như các yếu tố kiểm soát. Kết quả nghiên cứu định tính, định lượng và dữ liệu thứ cấp được đối chiếu, tổng hợp để đánh giá quản lý nhà nước về quyền sở hữu trí tuệ và đề xuất giải pháp.', bold_prefix='- ')

# 6. ĐÓNG GÓP MỚI
add_h2('6. Những đóng góp mới của luận án')
add_p('Luận án cụ thể hóa hệ thống tiêu chí đánh giá quản lý nhà nước về quyền sở hữu trí tuệ trong điều kiện kinh tế số, tập trung vào khả năng thích ứng của chính sách, tính minh bạch và chất lượng hỗ trợ doanh nghiệp tiếp cận, khai thác hệ thống sở hữu trí tuệ phục vụ đổi mới sáng tạo.', bold_prefix='Thứ nhất: ')
add_p('Luận án xây dựng và kiểm định khung phân tích về mối quan hệ giữa chất lượng quản lý nhà nước về quyền sở hữu trí tuệ theo đánh giá của doanh nghiệp, mức độ doanh nghiệp tiếp cận và sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo; đồng thời kiểm định vai trò trung gian của mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ.', bold_prefix='Thứ hai: ')
add_p('Luận án xác định các nhóm nội dung quản lý nhà nước có ảnh hưởng đáng kể đến mức độ doanh nghiệp tiếp cận và sử dụng hệ thống sở hữu trí tuệ phục vụ đổi mới sáng tạo, làm cơ sở đề xuất hệ thống giải pháp hoàn thiện đến năm 2035.', bold_prefix='Thứ ba: ')

# 7. KẾT CẤU LUẬN ÁN
add_h2('7. Dự kiến kết cấu của luận án')
add_p('Ngoài phần mở đầu, kết luận, danh mục từ viết tắt, tài liệu tham khảo và phụ lục, nội dung chính của luận án được kết cấu thành 04 chương:')

# DỰ KIẾN KHUNG ĐỀ CƯƠNG CHI TIẾT
add_h1('DỰ KIẾN KHUNG NỘI DUNG CHI TIẾT CỦA LUẬN ÁN')

# CHƯƠNG 1
add_h2('CHƯƠNG 1. TỔNG QUAN TÌNH HÌNH NGHIÊN CỨU VÀ PHƯƠNG PHÁP NGHIÊN CỨU')
add_p('1.1. Tổng quan tình hình nghiên cứu liên quan đến đề tài luận án')
add_p('1.1.1. Các nghiên cứu về quản lý nhà nước đối với quyền sở hữu trí tuệ trong nền kinh tế số')
add_p('1.1.2. Các nghiên cứu về tiếp cận và sử dụng hệ thống sở hữu trí tuệ của doanh nghiệp')
add_p('1.1.3. Các nghiên cứu về mối quan hệ giữa thể chế sở hữu trí tuệ và đổi mới sáng tạo của doanh nghiệp')
add_p('1.2. Đánh giá chung tình hình nghiên cứu và nhận diện khoảng trống nghiên cứu')
add_p('1.2.1. Đánh giá tổng quát những kết quả đã đạt được trong các nghiên cứu đi trước')
add_p('1.2.2. Những vấn đề còn tranh luận hoặc chưa được giải thích đầy đủ')
add_p('1.2.3. Khoảng trống nghiên cứu của đề tài luận án')
add_p('1.3. Khung nghiên cứu tổng quát của luận án')
add_p('1.4. Thiết kế nghiên cứu và phương pháp nghiên cứu')
add_p('1.4.1. Quy trình nghiên cứu tổng thể')
add_p('1.4.2. Phương pháp nghiên cứu định tính và phỏng vấn sâu chuyên gia')
add_p('1.4.3. Phương pháp nghiên cứu định lượng: Mẫu khảo sát doanh nghiệp, đo lường thang đo và mô hình phân tích')

# CHƯƠNG 2
add_h2('CHƯƠNG 2. CƠ SỞ LÝ LUẬN VÀ KINH NGHIỆM THỰC TIỄN VỀ QUẢN LÝ NHÀ NƯỚC VỀ QUYỀN SỞ HỮU TRÍ TUỆ TRONG NỀN KINH TẾ SỐ')
add_p('2.1. Các khái niệm cơ bản liên quan đến đề tài')
add_p('2.1.1. Khái niệm quyền sở hữu trí tuệ và các đối tượng quyền sở hữu trí tuệ trong nền kinh tế số')
add_p('2.1.2. Khái niệm kinh tế số và kết quả đổi mới sáng tạo của doanh nghiệp')
add_p('2.1.3. Khái niệm quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số')
add_p('2.1.4. Khái niệm mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ của doanh nghiệp')
add_p('2.2. Các lý thuyết nền tảng của quản lý nhà nước về quyền sở hữu trí tuệ')
add_p('2.2.1. Lý thuyết thất bại của thị trường')
add_p('2.2.2. Lý thuyết quyền tài sản')
add_p('2.2.3. Lý thuyết thể chế')
add_p('2.2.4. Lý thuyết hệ thống đổi mới quốc gia')
add_p('2.3. Đặc điểm và vai trò của quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số')
add_p('2.3.1. Đặc điểm của quản lý nhà nước về quyền sở hữu trí tuệ trong môi trường số')
add_p('2.3.2. Vai trò của quản lý nhà nước về quyền sở hữu trí tuệ đối với nền kinh tế và doanh nghiệp')
add_p('2.4. Nội dung và tiêu chí đánh giá quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số')
add_p('2.4.1. Xây dựng chiến lược, chính sách và pháp luật về quyền sở hữu trí tuệ')
add_p('2.4.2. Tổ chức bộ máy và cơ chế phối hợp liên ngành trong quản lý nhà nước về quyền sở hữu trí tuệ')
add_p('2.4.3. Tổ chức thực hiện và hỗ trợ xác lập, khai thác, chuyển giao, thương mại hóa quyền sở hữu trí tuệ')
add_p('2.4.4. Kiểm tra, giám sát, thực thi và xử lý vi phạm quyền sở hữu trí tuệ trên môi trường số')
add_p('2.4.5. Tiêu chí đánh giá quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số (tính hiệu lực, hiệu quả, phù hợp và khả năng thích ứng, minh bạch)')
add_p('2.5. Các yếu tố ảnh hưởng đến quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số')
add_p('2.5.1. Nhóm các yếu tố thuộc môi trường vĩ mô (trình độ công nghệ, hội nhập quốc tế, nhận thức xã hội)')
add_p('2.5.2. Nhóm các yếu tố thuộc điều kiện bảo đảm (chuyển đổi số, hạ tầng dữ liệu, tài chính, năng lực nhân lực)')
add_p('2.6. Mối quan hệ giữa quản lý nhà nước về quyền sở hữu trí tuệ, mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo của doanh nghiệp')
add_p('2.6.1. Mối quan hệ giữa chất lượng quản lý nhà nước và mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ')
add_p('2.6.2. Mối quan hệ giữa chất lượng quản lý nhà nước và kết quả đổi mới sáng tạo của doanh nghiệp')
add_p('2.6.3. Mối quan hệ giữa mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ và kết quả đổi mới sáng tạo')
add_p('2.6.4. Vai trò trung gian của mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ')
add_p('2.7. Kinh nghiệm quản lý nhà nước về quyền sở hữu trí tuệ trong kinh tế số của một số quốc gia và bài học cho Việt Nam')
add_p('2.7.1. Kinh nghiệm của một số quốc gia tiêu biểu (Singapore, Hàn Quốc, Liên minh châu Âu)')
add_p('2.7.2. Bài học kinh nghiệm rút ra cho Việt Nam')
add_p('2.8. Khung phân tích, mô hình và các giả thuyết nghiên cứu')

# CHƯƠNG 3
add_h2('CHƯƠNG 3. THỰC TRẠNG QUẢN LÝ NHÀ NƯỚC VỀ QUYỀN SỞ HỮU TRÍ TUỆ VÀ KẾT QUẢ NGHIÊN CỨU THỰC NGHIỆM Ở VIỆT NAM')
add_p('3.1. Bối cảnh nghiên cứu và yêu cầu đặt ra đối với quản lý nhà nước về quyền sở hữu trí tuệ ở Việt Nam')
add_p('3.1.1. Tổng quan bối cảnh phát triển kinh tế số và chuyển đổi số quốc gia ở Việt Nam')
add_p('3.1.2. Khái quát hoạt động đổi mới sáng tạo của doanh nghiệp trong nền kinh tế số')
add_p('3.1.3. Yêu cầu đặt ra đối với công tác quản lý nhà nước về quyền sở hữu trí tuệ')
add_p('3.2. Phân tích thực trạng quản lý nhà nước về quyền sở hữu trí tuệ ở Việt Nam giai đoạn 2019–2025 theo bốn nội dung quản lý')
add_p('3.2.1. Thực trạng xây dựng chiến lược, chính sách và pháp luật về quyền sở hữu trí tuệ')
add_p('3.2.2. Thực trạng tổ chức bộ máy và phối hợp quản lý nhà nước về quyền sở hữu trí tuệ')
add_p('3.2.3. Thực trạng tổ chức thực hiện và hỗ trợ xác lập, khai thác, chuyển giao, thương mại hóa quyền')
add_p('3.2.4. Thực trạng kiểm tra, giám sát, thực thi và xử lý vi phạm quyền sở hữu trí tuệ trên môi trường số')
add_p('3.3. Đánh giá quản lý nhà nước về quyền sở hữu trí tuệ theo các tiêu chí và các điều kiện bảo đảm')
add_p('3.3.1. Đánh giá quản lý nhà nước về quyền sở hữu trí tuệ theo các tiêu chí')
add_p('3.3.2. Thực trạng các điều kiện bảo đảm (chuyển đổi số dịch vụ công, hạ tầng dữ liệu, nhân lực và tài chính)')
add_p('3.4. Kết quả nghiên cứu thực nghiệm từ khảo sát doanh nghiệp')
add_p('3.4.1. Đặc điểm mẫu khảo sát, kiểm định độ tin cậy và phân tích nhân tố khám phá')
add_p('3.4.2. Kết quả phân tích tương quan và hồi quy tuyến tính đa biến')
add_p('3.4.3. Kết quả kiểm định vai trò trung gian của mức độ tiếp cận và sử dụng hệ thống sở hữu trí tuệ')
add_p('3.4.4. Thảo luận kết quả nghiên cứu thực nghiệm và diễn giải định tính từ phỏng vấn chuyên gia')
add_p('3.5. Đánh giá chung thực trạng quản lý nhà nước về quyền sở hữu trí tuệ giai đoạn 2019–2025')
add_p('3.5.1. Những kết quả đạt được')
add_p('3.5.2. Những hạn chế và bất cập')
add_p('3.5.3. Nguyên nhân của những hạn chế (phân định nguyên nhân chủ quan và khách quan)')

# CHƯƠNG 4
add_h2('CHƯƠNG 4. QUAN ĐIỂM, ĐỊNH HƯỚNG VÀ GIẢI PHÁP HOÀN THIỆN QUẢN LÝ NHÀ NƯỚC VỀ QUYỀN SỞ HỮU TRÍ TUỆ TRONG NỀN KINH TẾ SỐ Ở VIỆT NAM ĐẾN NĂM 2035')
add_p('4.1. Bối cảnh ảnh hưởng và dự báo xu hướng phát triển kinh tế số ở Việt Nam')
add_p('4.1.1. Xu hướng phát triển công nghệ và các vấn đề mới phát sinh đối với quản lý sở hữu trí tuệ')
add_p('4.1.2. Chủ trương, chiến lược của Đảng và Nhà nước về phát triển kinh tế số và đổi mới sáng tạo')
add_p('4.2. Quan điểm và định hướng hoàn thiện quản lý nhà nước về quyền sở hữu trí tuệ đến năm 2035')
add_p('4.2.1. Quan điểm chỉ đạo trong hoàn thiện quản lý nhà nước về quyền sở hữu trí tuệ')
add_p('4.2.2. Định hướng hoàn thiện quản lý nhà nước về quyền sở hữu trí tuệ đến năm 2035')
add_p('4.3. Giải pháp hoàn thiện quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số')
add_p('4.3.1. Hoàn thiện xây dựng chiến lược, chính sách và pháp luật về quyền sở hữu trí tuệ trong nền kinh tế số')
add_p('4.3.2. Kiện toàn tổ chức bộ máy và tăng cường phối hợp quản lý nhà nước về quyền sở hữu trí tuệ')
add_p('4.3.3. Đổi mới tổ chức thực hiện và nâng cao hiệu quả hỗ trợ xác lập, khai thác, chuyển giao, thương mại hóa quyền sở hữu trí tuệ')
add_p('4.3.4. Nâng cao hiệu quả kiểm tra, giám sát, thực thi và xử lý vi phạm quyền sở hữu trí tuệ trên môi trường số')
add_p('4.4. Giải pháp hoàn thiện các điều kiện bảo đảm thực hiện quản lý nhà nước')
add_p('4.4.1. Đẩy mạnh chuyển đổi số dịch vụ công và hoàn thiện hạ tầng dữ liệu sở hữu trí tuệ quốc gia')
add_p('4.4.2. Nâng cao năng lực và chất lượng nguồn nhân lực quản lý nhà nước về sở hữu trí tuệ')
add_p('4.5. Điều kiện thực hiện các giải pháp')

output_path = r'G:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\Đề cương NCS Uyen.docx'
doc.save(output_path)
print(f'Done! Successfully saved finalized official outline docx (97% completion rate): {output_path}')
