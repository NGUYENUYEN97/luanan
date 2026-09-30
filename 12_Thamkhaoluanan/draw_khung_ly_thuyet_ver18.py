from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

prs = Presentation()
prs.slide_width = Inches(11.69)
prs.slide_height = Inches(8.27)
slide = prs.slides.add_slide(prs.slide_layouts[6])

def add_box(slide, text, left, top, width, height, font_size=10, fill_color=(255,255,255)):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*fill_color)
    shape.line.color.rgb = RGBColor(0, 0, 0)
    shape.line.width = Pt(1.5)
    
    tf = shape.text_frame
    tf.text = text
    tf.word_wrap = True
    # Disable auto size to prevent text from jumping around, let it fit the height
    tf.auto_size = MSO_AUTO_SIZE.NONE
    
    for i, p in enumerate(tf.paragraphs):
        # Header formatting
        if i < 2:
            p.alignment = PP_ALIGN.CENTER
        elif "---" in p.text:
            p.alignment = PP_ALIGN.CENTER
        else:
            p.alignment = PP_ALIGN.LEFT
            # Add a slight left indent for bullet points
            p.level = 0
            
        for run in p.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(0, 0, 0)
            
            # Bold the headers
            if i < 2:
                run.font.bold = True
    return shape

def add_line_arrow(slide, begin_x, begin_y, end_x, end_y, dashed=False, arrow_end=True, arrow_start=False, thickness=1.5):
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, begin_x, begin_y, end_x, end_y)
    connector.line.color.rgb = RGBColor(0,0,0)
    connector.line.width = Pt(thickness)
    if dashed:
        connector.line.dash_style = 4
        
    ln = connector.element.spPr.ln
    if ln is not None:
        if arrow_end:
            tailEnd = parse_xml(r'<a:tailEnd type="triangle" w="med" len="med" %s/>' % nsdecls('a'))
            ln.append(tailEnd)
        if arrow_start:
            headEnd = parse_xml(r'<a:headEnd type="triangle" w="med" len="med" %s/>' % nsdecls('a'))
            ln.append(headEnd)
    return connector


# --- TITLE ---
t_title = slide.shapes.add_textbox(Inches(0.5), Inches(0.05), Inches(10.69), Inches(0.4))
t_title.text_frame.text = "KHUNG LÝ THUYẾT NGHIÊN CỨU (MÔ HÌNH NHÂN QUẢ 5 BIẾN)"
t_title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
t_title.text_frame.paragraphs[0].runs[0].font.name = 'Times New Roman'
t_title.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
t_title.text_frame.paragraphs[0].runs[0].font.bold = True

# --- CONTEXT BOX (Bối cảnh) ---
ctx_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.3), Inches(0.9), Inches(11.1), Inches(6.6))
ctx_shape.fill.solid()
ctx_shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
ctx_shape.line.color.rgb = RGBColor(0, 0, 0)
ctx_shape.line.width = Pt(2.0)
ctx_shape.line.dash_style = 4

lbl = slide.shapes.add_textbox(Inches(0.4), Inches(0.95), Inches(5.0), Inches(0.4))
lbl.text_frame.text = "BỐI CẢNH: NỀN KINH TẾ SỐ (Kỷ nguyên dữ liệu & AI)"
lbl.text_frame.paragraphs[0].runs[0].font.name = 'Times New Roman'
lbl.text_frame.paragraphs[0].runs[0].font.size = Pt(11)
lbl.text_frame.paragraphs[0].runs[0].font.bold = True
lbl.text_frame.paragraphs[0].runs[0].font.italic = True


# --- LÝ THUYẾT NỀN TẢNG (Top) ---
t1 = "LÝ THUYẾT NỀN TẢNG\nLý thuyết quản lý công  •  Lý thuyết thể chế  •  Lý thuyết tăng trưởng nội sinh  •  Lý thuyết về quyền sở hữu trí tuệ"
add_box(slide, t1, Inches(1.5), Inches(0.35), Inches(8.69), Inches(0.5), font_size=10, fill_color=(245, 245, 245))


# --- BIẾN ĐỘC LẬP (B1) ---
t2 = "BIẾN ĐỘC LẬP (NGUYÊN NHÂN)\nQUẢN LÝ NHÀ NƯỚC VỀ SHTT\n--------------------------------------------\n1. Xây dựng chiến lược & pháp luật\n2. Tổ chức thực hiện chính sách\n3. Quản lý xác lập & bảo vệ TSTT\n4. Thanh tra, giám sát & xử lý vi phạm"
add_box(slide, t2, Inches(0.6), Inches(3.2), Inches(3.2), Inches(2.2), font_size=10, fill_color=(245, 245, 245))

# --- BIẾN TRUNG GIAN (B2) ---
t3 = "BIẾN TRUNG GIAN (TRUYỀN DẪN)\nSỰ GIA TĂNG TÀI SẢN TRÍ TUỆ\n--------------------------------------------\n• Tăng trưởng văn bằng bảo hộ\n• Tỷ lệ thương mại hóa & chuyển giao\n• An toàn môi trường sở hữu trí tuệ"
add_box(slide, t3, Inches(4.25), Inches(3.2), Inches(3.2), Inches(2.2), font_size=10, fill_color=(245, 245, 245))

# --- BIẾN PHỤ THUỘC (B3) ---
t4 = "BIẾN PHỤ THUỘC (KẾT QUẢ)\nPHÁT TRIỂN KINH TẾ SỐ\n--------------------------------------------\n• Thúc đẩy Đổi mới sáng tạo toàn diện\n• Tăng trưởng kinh tế tri thức\n• Nâng cao năng lực cạnh tranh"
add_box(slide, t4, Inches(7.9), Inches(3.2), Inches(3.2), Inches(2.2), font_size=10, fill_color=(245, 245, 245))


# --- ARROWS BETWEEN CORE BOXES ---
add_line_arrow(slide, Inches(3.8), Inches(4.3), Inches(4.25), Inches(4.3), thickness=2.0) # B1 -> B2
add_line_arrow(slide, Inches(7.45), Inches(4.3), Inches(7.9), Inches(4.3), thickness=2.0) # B2 -> B3


# --- BIẾN ĐIỀU TIẾT (Top Center) ---
tdt = "BIẾN ĐIỀU TIẾT (MỨC ĐỘ)\nMÔI TRƯỜNG & HỘI NHẬP\n--------------------------------------------\n• Trình độ công nghệ số (AI, Data)\n• Mức độ hội nhập quốc tế\n• Nhận thức pháp luật của xã hội"
add_box(slide, tdt, Inches(4.25), Inches(1.3), Inches(3.2), Inches(1.5), font_size=10, fill_color=(245, 245, 245))
# Arrow from Điều tiết down to Trung gian
add_line_arrow(slide, Inches(5.85), Inches(2.8), Inches(5.85), Inches(3.2), dashed=True, thickness=1.5)


# --- BIẾN KIỂM SOÁT (Bottom Right) ---
tks = "BIẾN KIỂM SOÁT (CỐ ĐỊNH)\nĐẶC ĐIỂM KINH TẾ - XÃ HỘI\n--------------------------------------------\n• Quy mô & cơ cấu nền kinh tế\n• Đặc thù ngành nghề (Công nghệ...)\n• Đặc điểm nhân khẩu học"
add_box(slide, tks, Inches(7.9), Inches(5.8), Inches(3.2), Inches(1.5), font_size=10, fill_color=(245, 245, 245))
# Arrow from Kiểm soát up to Phụ thuộc
add_line_arrow(slide, Inches(9.5), Inches(5.8), Inches(9.5), Inches(5.4), dashed=True, thickness=1.5)


# --- PHẢN HỒI CHÍNH SÁCH ---
tfb = "PHẢN HỒI CHÍNH SÁCH\nThực tiễn kinh tế số ➔ Đánh giá ➔ Điều chỉnh ➔ Hoàn thiện Quản lý nhà nước"
add_box(slide, tfb, Inches(1.5), Inches(7.6), Inches(8.69), Inches(0.5), font_size=10, fill_color=(255, 255, 255))

# Từ Phụ thuộc xuống Phản hồi (X=8.5)
add_line_arrow(slide, Inches(8.5), Inches(5.4), Inches(8.5), Inches(7.6), dashed=True, arrow_end=True, thickness=1.5)

# Từ Phản hồi vòng về Độc lập (Outside Context box)
add_line_arrow(slide, Inches(1.5), Inches(7.85), Inches(0.15), Inches(7.85), dashed=True, arrow_end=False)
add_line_arrow(slide, Inches(0.15), Inches(7.85), Inches(0.15), Inches(4.3), dashed=True, arrow_end=False)
add_line_arrow(slide, Inches(0.15), Inches(4.3), Inches(0.6), Inches(4.3), dashed=True)


prs.save("KhungLyThuyet_KinhTe_5Bien_TuyetDep.pptx")
print("Saved PPTX with highly aesthetic symmetrical layout.")
