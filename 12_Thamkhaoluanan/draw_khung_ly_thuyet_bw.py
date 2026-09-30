from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE

prs = Presentation()
prs.slide_width = Inches(8.27)
prs.slide_height = Inches(11.69)
slide = prs.slides.add_slide(prs.slide_layouts[6])

def add_box(slide, text, left, top, width, height, font_size=11, bold=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    
    # Black and White ONLY
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(0, 0, 0)
    shape.line.width = Pt(1.5)
    
    text_frame = shape.text_frame
    text_frame.text = text
    text_frame.word_wrap = True
    text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    
    for i, paragraph in enumerate(text_frame.paragraphs):
        # Center align headers, left align bullet points/lists
        if "•" in paragraph.text or paragraph.text.startswith("-") or paragraph.text.startswith("Nhóm"):
            paragraph.alignment = PP_ALIGN.LEFT
        else:
            paragraph.alignment = PP_ALIGN.CENTER
            
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(font_size)
            run.font.color.rgb = RGBColor(0, 0, 0)
            # Make the first line (title) bold
            if i == 0:
                run.font.bold = True
            else:
                run.font.bold = bold
                
    return shape

def add_arrow(slide, left, top, width, height, direction='DOWN'):
    shape_type = MSO_SHAPE.DOWN_ARROW
    if direction == 'RIGHT': shape_type = MSO_SHAPE.RIGHT_ARROW
    elif direction == 'UP': shape_type = MSO_SHAPE.UP_ARROW
    elif direction == 'LEFT': shape_type = MSO_SHAPE.LEFT_ARROW
    
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0, 0, 0) # Solid black
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
t1_text = "LÝ THUYẾT NỀN TẢNG\n\n• Lý thuyết quản lý công                              • Lý thuyết quản trị thích ứng\n• Lý thuyết quản lý công mới                        • Lý thuyết thể chế\n• Lý thuyết quản trị số                                  • Lý thuyết về quyền sở hữu trí tuệ"
add_box(slide, t1_text, Inches(0.5), Inches(0.7), Inches(7.27), Inches(1.1), font_size=11)

add_arrow(slide, Inches(4.0), Inches(1.9), Inches(0.27), Inches(0.3))

# --- KHÁI NIỆM CỐT LÕI ---
t2_text = "KHÁI NIỆM CỐT LÕI\n\n• Quyền sở hữu trí tuệ     • Quản lý nhà nước     • Kinh tế số     • Quản trị số\n• Quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số"
add_box(slide, t2_text, Inches(0.5), Inches(2.3), Inches(7.27), Inches(0.8), font_size=11)

add_arrow(slide, Inches(4.0), Inches(3.2), Inches(0.27), Inches(0.3))

# --- MÔ HÌNH NHÂN QUẢ ---
# Dashed background container for main causal model
model_container = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.2), Inches(3.6), Inches(7.87), Inches(5.6))
model_container.fill.background()
model_container.line.color.rgb = RGBColor(0, 0, 0)
model_container.line.width = Pt(1.5)
model_container.line.dash_style = 4
slide.shapes._spTree.remove(model_container._element)
slide.shapes._spTree.insert(2, model_container._element)

tb_t3 = slide.shapes.add_textbox(Inches(0.3), Inches(3.7), Inches(4.0), Inches(0.3))
tb_t3.text_frame.text = "MÔ HÌNH NHÂN QUẢ GIỮA CÁC BIẾN"
tb_t3.text_frame.paragraphs[0].runs[0].font.name = 'Times New Roman'
tb_t3.text_frame.paragraphs[0].runs[0].font.bold = True
tb_t3.text_frame.paragraphs[0].runs[0].font.italic = True

# Biến độc lập
b_dl = "BIẾN ĐỘC LẬP\n(Nội dung quản lý nhà nước)\n\n• Xây dựng chiến lược, chính sách, pháp luật\n• Tổ chức thực hiện\n• Hỗ trợ xác lập, khai thác, thương mại hóa\n• Thanh tra, kiểm tra, giám sát, thực thi"
add_box(slide, b_dl, Inches(0.5), Inches(4.1), Inches(4.2), Inches(1.3), font_size=11)

add_arrow(slide, Inches(2.45), Inches(5.5), Inches(0.3), Inches(0.3))

# Biến trung gian
b_tg = "BIẾN TRUNG GIAN\n(Năng lực quản lý nhà nước)\n\n• Năng lực hoạch định chính sách\n• Năng lực tổ chức thực hiện\n• Năng lực phối hợp liên ngành\n• Năng lực ứng dụng công nghệ số\n• Năng lực thanh tra, giám sát, thực thi"
add_box(slide, b_tg, Inches(0.5), Inches(5.9), Inches(4.2), Inches(1.3), font_size=11)

add_arrow(slide, Inches(2.45), Inches(7.3), Inches(0.3), Inches(0.3))

# Biến phụ thuộc
b_pt = "BIẾN PHỤ THUỘC\n(Chất lượng quản lý nhà nước)\n\n• Tính hiệu lực\n• Tính hiệu quả\n• Tính phù hợp\n• Tính đồng bộ và thống nhất\n• Tính minh bạch"
add_box(slide, b_pt, Inches(0.5), Inches(7.7), Inches(4.2), Inches(1.3), font_size=11)

# Biến điều tiết
b_dt = "BIẾN ĐIỀU TIẾT\n(Các yếu tố ảnh hưởng)\n\nNhóm thể chế:\n• Mức độ hoàn thiện pháp luật\n• Chính sách phát triển kinh tế số\n\nNhóm công nghệ:\n• Hạ tầng số, AI, dữ liệu, nền tảng\n\nNhóm môi trường:\n• Trình độ phát triển kinh tế số\n• Năng lực đổi mới sáng tạo\n• Hội nhập, nhận thức xã hội"
add_box(slide, b_dt, Inches(5.1), Inches(4.7), Inches(2.8), Inches(3.6), font_size=11)

# Arrow from Biến điều tiết to the center flow
add_arrow(slide, Inches(4.8), Inches(6.4), Inches(0.3), Inches(0.2), 'LEFT')

add_arrow(slide, Inches(4.0), Inches(9.3), Inches(0.27), Inches(0.3))

# --- KẾT QUẢ / TÁC ĐỘNG ---
b_kq = "KẾT QUẢ VÀ TÁC ĐỘNG\n\n• Bảo vệ hiệu quả quyền SHTT        • Thúc đẩy đổi mới sáng tạo        • Phát triển tài sản trí tuệ\n• Nâng cao năng lực cạnh tranh quốc gia        • Thúc đẩy phát triển kinh tế số"
add_box(slide, b_kq, Inches(0.5), Inches(9.7), Inches(7.3), Inches(0.9), font_size=11)

# --- FEEDBACK LOOP ---
b_fb = "VÒNG PHẢN HỒI CHÍNH SÁCH\n\nKết quả thực hiện  ➔  Đánh giá  ➔  Điều chỉnh (chính sách, thể chế, công cụ)  ➔  Hoàn thiện"
add_box(slide, b_fb, Inches(0.5), Inches(10.8), Inches(7.3), Inches(0.6), font_size=11)

# Feedback dashed line - All Black
l_up = slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, Inches(0.3), Inches(11.1), Inches(0), Inches(6.9)) # from 11.1 up to 4.2
l_up.line.color.rgb = RGBColor(0,0,0)
l_up.line.width = Pt(1.5)
l_up.line.dash_style = 4 

l_horz_bot = slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, Inches(0.3), Inches(11.1), Inches(0.2), Inches(0))
l_horz_bot.line.color.rgb = RGBColor(0,0,0)
l_horz_bot.line.width = Pt(1.5)
l_horz_bot.line.dash_style = 4

l_horz_top = slide.shapes.add_shape(MSO_SHAPE.LINE_INVERSE, Inches(0.3), Inches(4.2), Inches(0.2), Inches(0))
l_horz_top.line.color.rgb = RGBColor(0,0,0)
l_horz_top.line.width = Pt(1.5)
l_horz_top.line.dash_style = 4

arr_fb = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, Inches(0.4), Inches(4.1), Inches(0.1), Inches(0.2))
arr_fb.fill.solid()
arr_fb.fill.fore_color.rgb = RGBColor(0,0,0)
arr_fb.line.color.rgb = RGBColor(0,0,0)

prs.save("KhungLyThuyet_BW.pptx")
print("Saved black and white layout.")
