from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

prs = Presentation()
prs.slide_width = Inches(8.27)
prs.slide_height = Inches(11.69)
slide = prs.slides.add_slide(prs.slide_layouts[6])

def add_box(slide, text, left, top, width, height, font_size=10, bold=False, fill_color=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        
    shape.line.color.rgb = RGBColor(0, 0, 0)
    shape.line.width = Pt(1.5)
    
    text_frame = shape.text_frame
    text_frame.text = text
    text_frame.word_wrap = True
    
    for paragraph in text_frame.paragraphs:
        if "\n-" in text or "\nX" in text or "\nM" in text or "\nY" in text or "\nO" in text or "NHÓM" in text:
            pass # Keep default left align for lists
        else:
            paragraph.alignment = PP_ALIGN.CENTER
            
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.font.bold = bold
    return shape

def add_arrow(slide, left, top, width, height, direction='DOWN'):
    shape_type = MSO_SHAPE.DOWN_ARROW
    if direction == 'RIGHT': shape_type = MSO_SHAPE.RIGHT_ARROW
    elif direction == 'UP': shape_type = MSO_SHAPE.UP_ARROW
    elif direction == 'LEFT': shape_type = MSO_SHAPE.LEFT_ARROW
    
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(89, 89, 89) # Dark gray arrows
    shape.line.color.rgb = RGBColor(0, 0, 0)
    return shape

# Title
t_title = slide.shapes.add_textbox(Inches(0.5), Inches(0.1), Inches(7.27), Inches(0.4))
t_title.text_frame.text = "KHUNG LÝ THUYẾT NGHIÊN CỨU"
t_title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
t_title.text_frame.paragraphs[0].runs[0].font.name = 'Times New Roman'
t_title.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
t_title.text_frame.paragraphs[0].runs[0].font.bold = True

# --- LÝ THUYẾT NỀN TẢNG ---
t1_text = "LÝ THUYẾT NỀN TẢNG\n\nLý thuyết quản lý công        |        Lý thuyết quản lý công mới        |        Lý thuyết quản trị số\nLý thuyết quản trị thích ứng        |        Lý thuyết thể chế        |        Lý thuyết về quyền sở hữu trí tuệ"
add_box(slide, t1_text, Inches(0.5), Inches(0.7), Inches(7.27), Inches(0.8), font_size=11, bold=True, fill_color=RGBColor(217, 225, 242))

add_arrow(slide, Inches(4.0), Inches(1.6), Inches(0.27), Inches(0.3))

# --- KHÁI NIỆM CỐT LÕI ---
t2_text = "KHÁI NIỆM CỐT LÕI\n\nQuyền sở hữu trí tuệ  -  Quản lý nhà nước  -  Kinh tế số  -  Quản trị số\nQuản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số"
add_box(slide, t2_text, Inches(0.5), Inches(2.0), Inches(7.27), Inches(0.8), font_size=11, bold=True, fill_color=RGBColor(255, 242, 204))

add_arrow(slide, Inches(4.0), Inches(2.9), Inches(0.27), Inches(0.3))

# --- MÔ HÌNH NHÂN QUẢ ---
# Dashed background container for main causal model
model_container = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.2), Inches(3.3), Inches(7.87), Inches(5.4))
model_container.fill.background()
model_container.line.color.rgb = RGBColor(0, 0, 0)
model_container.line.width = Pt(1.5)
model_container.line.dash_style = 4
slide.shapes._spTree.remove(model_container._element)
slide.shapes._spTree.insert(2, model_container._element)

tb_t3 = slide.shapes.add_textbox(Inches(0.3), Inches(3.4), Inches(4.0), Inches(0.3))
tb_t3.text_frame.text = "MÔ HÌNH NHÂN QUẢ GIỮA CÁC BIẾN"
tb_t3.text_frame.paragraphs[0].runs[0].font.name = 'Times New Roman'
tb_t3.text_frame.paragraphs[0].runs[0].font.bold = True
tb_t3.text_frame.paragraphs[0].runs[0].font.italic = True

# Biến độc lập
b_dl = "BIẾN ĐỘC LẬP: NỘI DUNG QUẢN LÝ NHÀ NƯỚC\n\nX1. Xây dựng chiến lược, chính sách, pháp luật\nX2. Tổ chức thực hiện\nX3. Hỗ trợ xác lập, khai thác, thương mại hóa\nX4. Thanh tra, kiểm tra, giám sát và thực thi"
add_box(slide, b_dl, Inches(0.5), Inches(3.8), Inches(4.2), Inches(1.3), font_size=11, bold=True, fill_color=RGBColor(226, 240, 217))

add_arrow(slide, Inches(2.45), Inches(5.2), Inches(0.3), Inches(0.3))

# Biến trung gian
b_tg = "BIẾN TRUNG GIAN: NĂNG LỰC QUẢN LÝ\n\nM1. Năng lực hoạch định chính sách\nM2. Năng lực tổ chức thực hiện\nM3. Năng lực phối hợp liên ngành\nM4. Năng lực ứng dụng công nghệ số\nM5. Năng lực thanh tra, giám sát, thực thi"
add_box(slide, b_tg, Inches(0.5), Inches(5.6), Inches(4.2), Inches(1.3), font_size=11, bold=True, fill_color=RGBColor(252, 228, 214))

add_arrow(slide, Inches(2.45), Inches(7.0), Inches(0.3), Inches(0.3))

# Biến phụ thuộc
b_pt = "BIẾN PHỤ THUỘC: CHẤT LƯỢNG QUẢN LÝ\n\nY1. Hiệu lực\nY2. Hiệu quả\nY3. Tính phù hợp\nY4. Tính đồng bộ và thống nhất\nY5. Tính minh bạch"
add_box(slide, b_pt, Inches(0.5), Inches(7.4), Inches(4.2), Inches(1.1), font_size=11, bold=True, fill_color=RGBColor(226, 240, 217))

# Biến điều tiết
b_dt = "BIẾN ĐIỀU TIẾT\n\nNHÓM THỂ CHẾ:\n- Mức độ hoàn thiện pháp luật\n- Chính sách phát triển KT số\n\nNHÓM CÔNG NGHỆ:\n- Hạ tầng số, AI, Dữ liệu, Nền tảng\n\nNHÓM MÔI TRƯỜNG:\n- Trình độ phát triển kinh tế số\n- Năng lực đổi mới sáng tạo\n- Hội nhập, Nhận thức xã hội"
add_box(slide, b_dt, Inches(5.1), Inches(4.5), Inches(2.8), Inches(3.4), font_size=11, bold=True, fill_color=RGBColor(237, 226, 246))

# Arrow from Biến điều tiết to the center flow
add_arrow(slide, Inches(4.8), Inches(6.1), Inches(0.3), Inches(0.2), 'LEFT')

add_arrow(slide, Inches(4.0), Inches(8.8), Inches(0.27), Inches(0.3))

# --- KẾT QUẢ / TÁC ĐỘNG ---
b_kq = "KẾT QUẢ / TÁC ĐỘNG\n\nO1. Bảo vệ hiệu quả quyền SHTT      O2. Thúc đẩy đổi mới sáng tạo      O3. Phát triển tài sản trí tuệ\nO4. Nâng cao năng lực cạnh tranh quốc gia      O5. Thúc đẩy phát triển kinh tế số"
add_box(slide, b_kq, Inches(0.5), Inches(9.2), Inches(7.3), Inches(0.9), font_size=11, bold=True, fill_color=RGBColor(255, 230, 153))

# --- FEEDBACK LOOP ---
b_fb = "VÒNG PHẢN HỒI CHÍNH SÁCH (FEEDBACK)\n\nKết quả thực hiện   ➜   Đánh giá   ➜   Điều chỉnh (Chính sách, Thể chế, Công cụ)   ➜   Hoàn thiện"
add_box(slide, b_fb, Inches(0.5), Inches(10.5), Inches(7.3), Inches(0.7), font_size=11, bold=True, fill_color=RGBColor(242, 242, 242))

# Feedback dashed line
l_up = slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, Inches(0.3), Inches(10.85), Inches(0), Inches(6.85)) # from 10.85 up to 4.0
l_up.line.color.rgb = RGBColor(255,0,0)
l_up.line.width = Pt(1.5)
l_up.line.dash_style = 4 

l_horz_bot = slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, Inches(0.3), Inches(10.85), Inches(0.2), Inches(0))
l_horz_bot.line.color.rgb = RGBColor(255,0,0)
l_horz_bot.line.width = Pt(1.5)
l_horz_bot.line.dash_style = 4

l_horz_top = slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, Inches(0.3), Inches(4.0), Inches(0.2), Inches(0))
l_horz_top.line.color.rgb = RGBColor(255,0,0)
l_horz_top.line.width = Pt(1.5)
l_horz_top.line.dash_style = 4

arr_fb = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(0.4), Inches(3.9), Inches(0.1), Inches(0.2))
arr_fb.fill.solid()
arr_fb.fill.fore_color.rgb = RGBColor(255,0,0)
arr_fb.line.color.rgb = RGBColor(255,0,0)

prs.save("KhungLyThuyet_SieuCap_GS_Dep.pptx")
print("Saved beautiful super master theoretical framework.")
