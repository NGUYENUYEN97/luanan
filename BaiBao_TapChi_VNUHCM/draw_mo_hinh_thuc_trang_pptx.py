from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

prs = Presentation()
prs.slide_width = Inches(10)
prs.slide_height = Inches(5.625)
slide = prs.slides.add_slide(prs.slide_layouts[6])

def add_box(slide, text, left, top, width, height):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(0, 0, 0)
    shape.line.width = Pt(2.0)
    
    # Shadow effect like the image (optional but nice)
    shape.shadow.inherit = False
    
    tf = shape.text_frame
    tf.text = text
    tf.word_wrap = True
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(13)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.font.bold = True
    return shape

# Boxes
box_a = add_box(slide, "1. Chủ thể\nquyền SHTT", Inches(0.5), Inches(1.0), Inches(2.0), Inches(1.0))
box_b = add_box(slide, "2. Nền tảng số /\nSàn TMĐT", Inches(3.8), Inches(1.0), Inches(2.2), Inches(1.0))
box_d = add_box(slide, "Tài khoản /\nCửa hàng\nvi phạm", Inches(7.5), Inches(1.0), Inches(2.0), Inches(1.0))
box_c = add_box(slide, "3. Cơ quan Quản lý Nhà nước\n(QLTT, Thanh tra, Công an)", Inches(3.5), Inches(3.8), Inches(3.0), Inches(1.0))

def add_arrow(connector, dashed=False):
    ln = connector.element.spPr.ln
    if ln is not None:
        tailEnd = parse_xml(r'<a:tailEnd type="triangle" w="med" len="med" %s/>' % nsdecls('a'))
        ln.append(tailEnd)
    if dashed:
        connector.line.dash_style = 4 # Dashed

# Line 1: A -> B (Straight)
conn1 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(2.5), Inches(1.5), Inches(3.8), Inches(1.5))
conn1.line.color.rgb = RGBColor(0, 0, 0)
conn1.line.width = Pt(1.5)
add_arrow(conn1)

# Line 2: B -> D (Straight)
conn2 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(6.0), Inches(1.5), Inches(7.5), Inches(1.5))
conn2.line.color.rgb = RGBColor(0, 0, 0)
conn2.line.width = Pt(1.5)
add_arrow(conn2)

# Line 3: A -> C (Elbow: Down then Right)
conn3 = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, Inches(1.5), Inches(2.0), Inches(3.5), Inches(4.3))
conn3.line.color.rgb = RGBColor(0, 0, 0)
conn3.line.width = Pt(1.5)
add_arrow(conn3)

# Line 4: C -> D (Elbow: Right then Up)
conn4 = slide.shapes.add_connector(MSO_CONNECTOR.ELBOW, Inches(6.5), Inches(4.3), Inches(8.5), Inches(2.0))
conn4.line.color.rgb = RGBColor(0, 0, 0)
conn4.line.width = Pt(1.5)
add_arrow(conn4)

# Line 5: C -> B (Dashed vertical arrow)
conn5 = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(4.9), Inches(3.8), Inches(4.9), Inches(2.0))
conn5.line.color.rgb = RGBColor(0, 0, 0)
conn5.line.width = Pt(1.0)
add_arrow(conn5, dashed=True)

def add_label(slide, text, left, top, width=1.5):
    txBox = slide.shapes.add_textbox(left, top, Inches(width), Inches(0.4))
    tf = txBox.text_frame
    tf.text = text
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run.font.italic = True

add_label(slide, "1. Phát hiện vi phạm/\nYêu cầu gỡ bỏ", Inches(2.3), Inches(1.0), 1.6)
add_label(slide, "2. Rà soát & gỡ bỏ", Inches(6.0), Inches(1.2), 1.6)
add_label(slide, "3. Khiếu nại / Tố cáo", Inches(1.5), Inches(3.5), 2.0)
add_label(slide, "4. Kiểm tra / Xử phạt", Inches(6.4), Inches(3.5), 2.0)
add_label(slide, "Yêu cầu cung\ncấp thông\ntin/chứng cứ", Inches(4.9), Inches(2.3), 1.2)

prs.save("MoHinhThucTrangPhapLuat.pptx")
