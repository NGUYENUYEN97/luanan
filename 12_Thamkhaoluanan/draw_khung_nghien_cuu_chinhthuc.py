"""
KHUNG NGHIÊN CỨU CHÍNH THỨC - LUẬN ÁN TIẾN SĨ
Đề tài: Quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số ở Việt Nam

Tác giả: NCS Uyên
Phiên bản: CHÍNH THỨC (v1.0)
Ngày: 01/07/2026

Mô hình: Biến số kinh lượng (5 biến độc lập, 2 biến trung gian, 1 biến phụ thuộc,
          biến điều tiết, biến kiểm soát, phản hồi chính sách)
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls
import os

# ===========================================================
# THIẾT LẬP SLIDE
# ===========================================================
prs = Presentation()
prs.slide_width = Inches(13.33)   # 16:9 widescreen
prs.slide_height = Inches(7.5)
slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

# --- Nền slide trắng ---
bg = slide.background
fill = bg.fill
fill.solid()
fill.fore_color.rgb = RGBColor(255, 255, 255)


# ===========================================================
# HÀM TIỆN ÍCH
# ===========================================================
def add_rounded_box(slide, text_lines, left, top, width, height,
                    font_size=10, fill_color=(255,255,255),
                    border_color=(50,50,50), border_width=1.5,
                    header_lines=1, dashed_border=False,
                    header_font_size=None, body_alignment=PP_ALIGN.LEFT):
    """Tạo hộp bo góc với tiêu đề in đậm căn giữa và nội dung căn trái."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(*fill_color)
    shape.line.color.rgb = RGBColor(*border_color)
    shape.line.width = Pt(border_width)
    if dashed_border:
        shape.line.dash_style = 4

    # Điều chỉnh bo góc
    shape.adjustments[0] = 0.05

    tf = shape.text_frame
    tf.text = text_lines
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = Pt(8)
    tf.margin_right = Pt(8)
    tf.margin_top = Pt(6)
    tf.margin_bottom = Pt(6)

    _header_fs = header_font_size if header_font_size else font_size

    for i, p in enumerate(tf.paragraphs):
        if i < header_lines:
            p.alignment = PP_ALIGN.CENTER
        else:
            p.alignment = body_alignment

        for run in p.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(_header_fs if i < header_lines else font_size)
            run.font.color.rgb = RGBColor(30, 30, 30)
            if i < header_lines:
                run.font.bold = True

    return shape


def add_arrow(slide, x1, y1, x2, y2, dashed=False, thickness=1.5,
              arrow_end=True, arrow_start=False, color=(60, 60, 60)):
    """Tạo mũi tên nối giữa hai điểm."""
    connector = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, x1, y1, x2, y2)
    connector.line.color.rgb = RGBColor(*color)
    connector.line.width = Pt(thickness)
    if dashed:
        connector.line.dash_style = 4

    ln = connector.element.spPr.ln
    if ln is not None:
        if arrow_end:
            tailEnd = parse_xml(
                r'<a:tailEnd type="triangle" w="med" len="med" %s/>' % nsdecls('a'))
            ln.append(tailEnd)
        if arrow_start:
            headEnd = parse_xml(
                r'<a:headEnd type="triangle" w="med" len="med" %s/>' % nsdecls('a'))
            ln.append(headEnd)
    return connector


def add_label(slide, text, left, top, width, height,
              font_size=9, bold=False, italic=False, color=(80,80,80),
              alignment=PP_ALIGN.CENTER):
    """Thêm nhãn văn bản tự do."""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.text = text
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = alignment
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(font_size)
        run.font.bold = bold
        run.font.italic = italic
        run.font.color.rgb = RGBColor(*color)
    return tb


# ===========================================================
# BỐ CỤC TỔNG THỂ (tọa độ tính bằng Inches)
# ===========================================================
# Margin
ML = 0.4   # Left margin
MT = 0.1   # Top margin

# Kích thước các khối chính
BOX_W = 3.2   # Chiều rộng hộp biến
BOX_H_IND = 3.2  # Chiều cao hộp biến độc lập (cao hơn)
BOX_H_MED = 1.8  # Chiều cao hộp biến trung gian
BOX_H_DEP = 1.8  # Chiều cao hộp biến phụ thuộc

# Vị trí Y trung tâm (hàng chính)
Y_CENTER = 3.0

# Vị trí X các cột
X_IND = ML + 0.3                    # Biến độc lập
X_MED = X_IND + BOX_W + 1.2         # Biến trung gian
X_DEP = X_MED + BOX_W + 1.2         # Biến phụ thuộc


# ===========================================================
# 1. TIÊU ĐỀ
# ===========================================================
title = add_label(slide, "KHUNG NGHIÊN CỨU ĐỀ TÀI", Inches(0.5), Inches(MT),
                  Inches(12.33), Inches(0.35), font_size=16, bold=True,
                  color=(20, 20, 20))

subtitle = add_label(slide, "Quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số ở Việt Nam",
                      Inches(0.5), Inches(MT + 0.32), Inches(12.33), Inches(0.3),
                      font_size=12, italic=True, color=(60, 60, 60))


# ===========================================================
# 2. LÝ THUYẾT NỀN TẢNG (Thanh ngang phía trên)
# ===========================================================
t_theory = "LÝ THUYẾT NỀN TẢNG\nLý thuyết thể chế   •   Lý thuyết quản lý công   •   Lý thuyết quản trị số   •   Lý thuyết quyền SHTT   •   Lý thuyết hành vi"
add_rounded_box(slide, t_theory,
                Inches(ML + 0.3), Inches(0.75), Inches(12.33), Inches(0.6),
                font_size=10, fill_color=(248, 249, 250), border_color=(100, 100, 100),
                header_lines=1, header_font_size=11, body_alignment=PP_ALIGN.CENTER)


# ===========================================================
# 3. BỐI CẢNH NỀN KINH TẾ SỐ (Viền bao quanh nét đứt)
# ===========================================================
ctx_left = Inches(ML)
ctx_top = Inches(1.55)
ctx_width = Inches(12.53)
ctx_height = Inches(5.4)

ctx_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, ctx_left, ctx_top, ctx_width, ctx_height)
ctx_shape.fill.solid()
ctx_shape.fill.fore_color.rgb = RGBColor(253, 253, 255)
ctx_shape.line.color.rgb = RGBColor(130, 130, 180)
ctx_shape.line.width = Pt(1.5)
ctx_shape.line.dash_style = 4
ctx_shape.adjustments[0] = 0.02

add_label(slide, "BỐI CẢNH: NỀN KINH TẾ SỐ",
          Inches(ML + 0.15), Inches(1.6), Inches(3.0), Inches(0.3),
          font_size=10, bold=True, italic=True, color=(80, 80, 140),
          alignment=PP_ALIGN.LEFT)


# ===========================================================
# 4. BIẾN ĐỘC LẬP (Trái)
# ===========================================================
Y_IND = Y_CENTER - BOX_H_IND / 2 + 0.3
t_ind = ("BIẾN ĐỘC LẬP\n(CÁC YẾU TỐ ẢNH HƯỞNG)\n\n"
         "X₁. Khung pháp lý và chính sách\n"
         "X₂. Năng lực tổ chức và cán bộ\n"
         "X₃. Hạ tầng công nghệ số\n"
         "X₄. Phối hợp liên ngành\n"
         "X₅. Nhận thức của các chủ thể")
add_rounded_box(slide, t_ind,
                Inches(X_IND), Inches(Y_IND), Inches(BOX_W), Inches(BOX_H_IND),
                font_size=11, fill_color=(255, 255, 255), border_color=(40, 40, 40),
                header_lines=2, header_font_size=11)


# ===========================================================
# 5. BIẾN TRUNG GIAN (Giữa)
# ===========================================================
Y_MED = Y_CENTER - BOX_H_MED / 2 + 0.6
t_med = ("BIẾN TRUNG GIAN\n(CƠ CHẾ TRUYỀN DẪN)\n\n"
         "M₁. Năng lực phát hiện và\n      xử lý vi phạm trên môi\n      trường số\n"
         "M₂. Mức độ tuân thủ pháp\n      luật sở hữu trí tuệ")
add_rounded_box(slide, t_med,
                Inches(X_MED), Inches(Y_MED), Inches(BOX_W), Inches(BOX_H_MED + 0.8),
                font_size=11, fill_color=(255, 255, 255), border_color=(40, 40, 40),
                header_lines=2, header_font_size=11)


# ===========================================================
# 6. BIẾN PHỤ THUỘC (Phải)
# ===========================================================
Y_DEP = Y_CENTER - BOX_H_DEP / 2 + 0.6
t_dep = ("BIẾN PHỤ THUỘC\n(KẾT QUẢ NGHIÊN CỨU)\n\n"
         "Y: Hiệu quả quản lý nhà\n   nước về quyền sở hữu\n   trí tuệ trong nền\n   kinh tế số")
add_rounded_box(slide, t_dep,
                Inches(X_DEP), Inches(Y_DEP), Inches(BOX_W), Inches(BOX_H_DEP + 0.5),
                font_size=11, fill_color=(255, 255, 255), border_color=(40, 40, 40),
                header_lines=2, header_font_size=11)


# ===========================================================
# 7. MŨI TÊN CHÍNH: Độc lập → Trung gian → Phụ thuộc
# ===========================================================
# Tính Y center cho mũi tên ngang
arrow_y = Inches(Y_CENTER + 0.6)

# Độc lập → Trung gian
add_arrow(slide,
          Inches(X_IND + BOX_W), arrow_y,
          Inches(X_MED), arrow_y,
          thickness=2.0)

# Trung gian → Phụ thuộc
add_arrow(slide,
          Inches(X_MED + BOX_W), arrow_y,
          Inches(X_DEP), arrow_y,
          thickness=2.0)


# ===========================================================
# 8. BIẾN ĐIỀU TIẾT (Phía trên giữa)
# ===========================================================
X_MOD = X_MED
Y_MOD = 1.85
t_mod = ("BIẾN ĐIỀU TIẾT\n(BỐI CẢNH TÁC ĐỘNG)\n\n"
         "W₁. Tốc độ phát triển công nghệ mới\n"
         "W₂. Mức độ hội nhập kinh tế quốc tế")
add_rounded_box(slide, t_mod,
                Inches(X_MOD), Inches(Y_MOD), Inches(BOX_W), Inches(1.4),
                font_size=10, fill_color=(255, 255, 255), border_color=(100, 100, 100),
                header_lines=2, dashed_border=False, header_font_size=10)

# Mũi tên từ Điều tiết xuống (chỉ vào đường nối Độc lập → Trung gian)
add_arrow(slide,
          Inches(X_MOD + BOX_W / 2), Inches(Y_MOD + 1.4),
          Inches(X_MOD + BOX_W / 2), arrow_y,
          dashed=True, thickness=1.2, color=(100, 100, 100))

add_label(slide, "Tăng/giảm\ncường độ",
          Inches(X_MOD + BOX_W / 2 + 0.05), Inches(Y_MOD + 1.4 + 0.05),
          Inches(1.0), Inches(0.35), font_size=8, italic=True, color=(120, 120, 120))


# ===========================================================
# 9. BIẾN KIỂM SOÁT (Phía dưới bên phải)
# ===========================================================
X_CTL = X_DEP
Y_CTL = Y_DEP + BOX_H_DEP + 0.5 + 0.5
t_ctl = ("BIẾN KIỂM SOÁT\n\n"
         "C₁. Quy mô và loại hình doanh nghiệp\n"
         "C₂. Lĩnh vực hoạt động\n"
         "C₃. Khu vực địa lý")
add_rounded_box(slide, t_ctl,
                Inches(X_CTL), Inches(Y_CTL), Inches(BOX_W), Inches(1.4),
                font_size=10, fill_color=(255, 255, 255), border_color=(100, 100, 100),
                header_lines=1, dashed_border=False, header_font_size=10)

# Mũi tên từ Kiểm soát lên Phụ thuộc
add_arrow(slide,
          Inches(X_CTL + BOX_W / 2), Inches(Y_CTL),
          Inches(X_CTL + BOX_W / 2), Inches(Y_DEP + BOX_H_DEP + 0.5),
          dashed=True, thickness=1.2, color=(100, 100, 100))

add_label(slide, "Kiểm soát\ntính khách quan",
          Inches(X_CTL + BOX_W / 2 + 0.05), Inches(Y_CTL - 0.38),
          Inches(1.2), Inches(0.35), font_size=8, italic=True, color=(120, 120, 120))


# ===========================================================
# 10. PHẢN HỒI CHÍNH SÁCH (Thanh ngang dưới cùng)
# ===========================================================
Y_FB = 6.45
t_fb = "PHẢN HỒI CHÍNH SÁCH:  Thực tiễn kinh tế số  ➔  Đánh giá hiệu quả  ➔  Điều chỉnh chính sách  ➔  Hoàn thiện quản lý nhà nước"
add_rounded_box(slide, t_fb,
                Inches(ML + 0.3), Inches(Y_FB), Inches(8.5), Inches(0.45),
                font_size=10, fill_color=(248, 249, 250), border_color=(100, 100, 100),
                header_lines=0, dashed_border=True, body_alignment=PP_ALIGN.CENTER)

# Mũi tên từ Phụ thuộc xuống Phản hồi (đi vòng bên phải)
dep_bottom_y = Inches(Y_DEP + BOX_H_DEP + 0.5)
fb_right_x = Inches(ML + 0.3 + 8.5)
fb_y = Inches(Y_FB + 0.225)

# Đường thẳng xuống từ phụ thuộc
add_arrow(slide,
          Inches(X_DEP + BOX_W + 0.1), Inches(Y_CENTER + 0.6),
          Inches(X_DEP + BOX_W + 0.6), Inches(Y_CENTER + 0.6),
          dashed=True, thickness=1.0, arrow_end=False, color=(120, 120, 120))
add_arrow(slide,
          Inches(X_DEP + BOX_W + 0.6), Inches(Y_CENTER + 0.6),
          Inches(X_DEP + BOX_W + 0.6), Inches(Y_FB + 0.225),
          dashed=True, thickness=1.0, arrow_end=False, color=(120, 120, 120))
add_arrow(slide,
          Inches(X_DEP + BOX_W + 0.6), Inches(Y_FB + 0.225),
          Inches(ML + 0.3 + 8.5), Inches(Y_FB + 0.225),
          dashed=True, thickness=1.0, color=(120, 120, 120))

# Mũi tên từ Phản hồi vòng lên Biến độc lập (đi bên trái)
add_arrow(slide,
          Inches(ML + 0.3), Inches(Y_FB + 0.225),
          Inches(ML + 0.05), Inches(Y_FB + 0.225),
          dashed=True, thickness=1.0, arrow_end=False, color=(120, 120, 120))
add_arrow(slide,
          Inches(ML + 0.05), Inches(Y_FB + 0.225),
          Inches(ML + 0.05), Inches(Y_CENTER + 0.6),
          dashed=True, thickness=1.0, arrow_end=False, color=(120, 120, 120))
add_arrow(slide,
          Inches(ML + 0.05), Inches(Y_CENTER + 0.6),
          Inches(X_IND), Inches(Y_CENTER + 0.6),
          dashed=True, thickness=1.0, color=(120, 120, 120))


# ===========================================================
# 11. NHÃN GIẢ THUYẾT TRÊN MŨI TÊN
# ===========================================================
# Nhãn H₁–H₆ trên mũi tên Độc lập → Trung gian
add_label(slide, "H₁ – H₆",
          Inches(X_IND + BOX_W + 0.15), Inches(Y_CENTER + 0.6 - 0.35),
          Inches(0.9), Inches(0.25), font_size=9, bold=True, color=(30, 100, 180))

# Nhãn H₇, H₈ trên mũi tên Trung gian → Phụ thuộc
add_label(slide, "H₇, H₈",
          Inches(X_MED + BOX_W + 0.15), Inches(Y_CENTER + 0.6 - 0.35),
          Inches(0.9), Inches(0.25), font_size=9, bold=True, color=(30, 100, 180))

# Nhãn H₉ trên mũi tên Điều tiết
add_label(slide, "H₉",
          Inches(X_MOD + BOX_W / 2 - 0.45), Inches(Y_MOD + 1.4 + 0.15),
          Inches(0.4), Inches(0.2), font_size=9, bold=True, color=(30, 100, 180))


# ===========================================================
# 12. MŨI TÊN TỪ LÝ THUYẾT NỀN TẢNG XUỐNG CÁC BIẾN
# ===========================================================
theory_bottom = Inches(0.75 + 0.6)

# → Biến độc lập
add_arrow(slide,
          Inches(X_IND + BOX_W / 2), theory_bottom,
          Inches(X_IND + BOX_W / 2), Inches(Y_IND),
          dashed=True, thickness=1.0, color=(130, 130, 130))

# → Biến điều tiết
add_arrow(slide,
          Inches(X_MOD + BOX_W / 2), theory_bottom,
          Inches(X_MOD + BOX_W / 2), Inches(Y_MOD),
          dashed=True, thickness=1.0, color=(130, 130, 130))

# → Biến phụ thuộc
add_arrow(slide,
          Inches(X_DEP + BOX_W / 2), theory_bottom,
          Inches(X_DEP + BOX_W / 2), Inches(Y_DEP),
          dashed=True, thickness=1.0, color=(130, 130, 130))


# ===========================================================
# LƯU FILE
# ===========================================================
output_dir = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\12_Thamkhaoluanan"
output_path = os.path.join(output_dir, "KhungNghienCuu_ChinhThuc_v1.pptx")
prs.save(output_path)
print(f"[OK] Da luu khung nghien cuu chinh thuc tai:\n   {output_path}")
