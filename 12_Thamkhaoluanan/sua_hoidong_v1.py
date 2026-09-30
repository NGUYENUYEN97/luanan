"""
Sửa khung theo góp ý hội đồng phản biện - 6 điểm:
3.1 Làm rõ khái niệm trung tâm
3.2 Đổi HIỆU QUẢ → CHẤT LƯỢNG QLNN
3.3 Bổ sung tiêu chí tính phù hợp và đồng bộ
3.4 Sửa 4 nội dung QLNN đúng phạm vi luận án
3.5 Phân 2 nhóm yếu tố bên trong / bên ngoài
3.6 Đưa Đặc trưng KTS vào nhóm bên ngoài
"""
import sys; sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

SRC = r"g:\My Drive\Luan an uyen 26\KhungPhanTich_Final.pptx"
DST = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\12_Thamkhaoluanan\KhungPhanTich_HoiDong_v1.pptx"

prs = Presentation(SRC)
slide = prs.slides[0]
shapes = slide.shapes

def rgb(h): return RGBColor.from_string(h)

def clear_text(shape):
    tf = shape.text_frame
    while len(tf.paragraphs) > 1:
        p = tf.paragraphs[-1]._p
        p.getparent().remove(p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        p0._p.remove(r._r)

def set_text(shape, lines, anchor='top'):
    clear_text(shape)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP if anchor == 'top' else MSO_ANCHOR.MIDDLE

    for i, ln in enumerate(lines):
        txt  = ln[0]
        sz   = ln[1]
        bold = ln[2] if len(ln) > 2 else False
        ital = ln[3] if len(ln) > 3 else False
        algn = ln[4] if len(ln) > 4 else PP_ALIGN.LEFT
        color= ln[5] if len(ln) > 5 else '000000'

        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = algn
        p.space_before = Pt(1)
        p.space_after = Pt(1)

        if txt == '':
            p.space_before = Pt(sz)
            continue

        run = p.add_run()
        run.text = txt
        run.font.name = 'Times New Roman'
        run.font.size = Pt(sz)
        run.font.bold = bold
        run.font.italic = ital
        run.font.color.rgb = rgb(color)

def hide_shape(shape):
    """Ẩn shape bằng cách đặt fill trắng, viền trắng, xóa text"""
    try:
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb('FFFFFF')
        shape.line.color.rgb = rgb('FFFFFF')
    except: pass
    clear_text(shape)

# ═══════════════════════════════════════════════════════════
# 3.1 TIÊU ĐỀ - Làm rõ khái niệm trung tâm
# ═══════════════════════════════════════════════════════════
set_text(shapes[0], [
    ('KHUNG PHÂN TÍCH', 14, True, False, PP_ALIGN.LEFT),
    ('Quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số', 11, False, True, PP_ALIGN.LEFT, '333333'),
])

# ═══════════════════════════════════════════════════════════
# LÝ THUYẾT NỀN TẢNG - giữ nguyên, bổ sung lý thuyết kinh tế số
# ═══════════════════════════════════════════════════════════
set_text(shapes[1], [
    ('LÝ THUYẾT NỀN TẢNG', 11, True, False, PP_ALIGN.CENTER),
    ('', 3),
    ('Lý thuyết quản lý công mới  •  Lý thuyết thể chế  •  Lý thuyết quản trị số  •  Lý thuyết quyền sở hữu trí tuệ', 9, False, False, PP_ALIGN.CENTER, '333333'),
])

# ═══════════════════════════════════════════════════════════
# 3.5 YẾU TỐ ẢNH HƯỞNG - Đổi nhãn outer box
# ═══════════════════════════════════════════════════════════
set_text(shapes[2], [
    ('CÁC YẾU TỐ ẢNH HƯỞNG', 10.5, True, False, PP_ALIGN.CENTER),
])

# ═══════════════════════════════════════════════════════════
# 3.5 PHÂN 2 NHÓM: Thay 6 box nhỏ → 2 box lớn
# Shape[3] → NHÓM BÊN TRONG (resize taller)
# Shape[4] → NHÓM BÊN NGOÀI (reposition + resize)
# Shape[5-8] → ẩn
# ═══════════════════════════════════════════════════════════

# Nhóm bên trong (shape[3]) - chiếm nửa trên
shapes[3].top = Cm(4.0)
shapes[3].height = Cm(4.8)
shapes[3].fill.solid()
shapes[3].fill.fore_color.rgb = rgb('F0F4FF')  # xanh nhạt
shapes[3].line.color.rgb = rgb('336699')
set_text(shapes[3], [
    ('NHÓM YẾU TỐ BÊN TRONG', 9.5, True, False, PP_ALIGN.CENTER, '003366'),
    ('', 4),
    ('(1) Khung pháp lý và chính sách', 9, False, False, PP_ALIGN.LEFT),
    ('(2) Năng lực tổ chức và cán bộ', 9, False, False, PP_ALIGN.LEFT),
    ('(3) Nguồn lực tài chính', 9, False, False, PP_ALIGN.LEFT),
    ('(4) Phối hợp thể chế', 9, False, False, PP_ALIGN.LEFT),
], anchor='top')

# Nhóm bên ngoài (shape[4]) - chiếm nửa dưới
shapes[4].top = Cm(9.3)
shapes[4].height = Cm(5.5)
shapes[4].fill.solid()
shapes[4].fill.fore_color.rgb = rgb('FFF5E6')  # cam nhạt
shapes[4].line.color.rgb = rgb('996633')
set_text(shapes[4], [
    ('NHÓM YẾU TỐ BÊN NGOÀI', 9.5, True, False, PP_ALIGN.CENTER, '663300'),
    ('', 4),
    ('(5) Hạ tầng công nghệ số', 9, False, False, PP_ALIGN.LEFT),
    ('(6) Đặc trưng nền kinh tế số', 9, False, False, PP_ALIGN.LEFT),
    ('(7) Nhận thức xã hội về quyền SHTT', 9, False, False, PP_ALIGN.LEFT),
    ('(8) Hội nhập quốc tế', 9, False, False, PP_ALIGN.LEFT),
], anchor='top')

# Ẩn 4 box cũ không còn dùng
for idx in [5, 6, 7, 8]:
    hide_shape(shapes[idx])

# ═══════════════════════════════════════════════════════════
# 3.1 + 3.4 NỘI DUNG QLNN - 4 nội dung theo phạm vi luận án
# ═══════════════════════════════════════════════════════════
set_text(shapes[9], [
    ('NỘI DUNG QLNN', 10.5, True, False, PP_ALIGN.CENTER),
    ('VỀ QUYỀN SHTT', 10.5, True, False, PP_ALIGN.CENTER),
    ('', 5),
    ('(1) Xây dựng chiến lược, chính', 9, False, False, PP_ALIGN.LEFT),
    ('     sách và pháp luật', 9, False, False, PP_ALIGN.LEFT),
    ('(2) Tổ chức thực hiện chiến', 9, False, False, PP_ALIGN.LEFT),
    ('     lược, chính sách', 9, False, False, PP_ALIGN.LEFT),
    ('(3) Hỗ trợ xác lập, khai thác,', 9, False, False, PP_ALIGN.LEFT),
    ('     thương mại hóa TSTT', 9, False, False, PP_ALIGN.LEFT),
    ('(4) Thanh tra, kiểm tra, giám', 9, False, False, PP_ALIGN.LEFT),
    ('     sát và thực thi quyền SHTT', 9, False, False, PP_ALIGN.LEFT),
], anchor='top')

# ═══════════════════════════════════════════════════════════
# 3.2 + 3.3 HIỆU QUẢ → CHẤT LƯỢNG QLNN + 5 tiêu chí
# ═══════════════════════════════════════════════════════════
set_text(shapes[10], [
    ('CHẤT LƯỢNG QLNN', 10.5, True, False, PP_ALIGN.CENTER),
    ('VỀ QUYỀN SHTT', 10.5, True, False, PP_ALIGN.CENTER),
    ('', 5),
    ('Tiêu chí đánh giá:', 9, True, False, PP_ALIGN.LEFT, '333333'),
    ('(1) Hiệu lực', 9, False, False, PP_ALIGN.LEFT),
    ('(2) Hiệu quả', 9, False, False, PP_ALIGN.LEFT),
    ('(3) Tính phù hợp', 9, False, False, PP_ALIGN.LEFT),
    ('(4) Tính đồng bộ và thống nhất', 9, False, False, PP_ALIGN.LEFT),
    ('(5) Tính minh bạch', 9, False, False, PP_ALIGN.LEFT),
], anchor='top')

# ═══════════════════════════════════════════════════════════
# KẾT QUẢ VÀ TÁC ĐỘNG - giữ, cập nhật text
# ═══════════════════════════════════════════════════════════
set_text(shapes[11], [
    ('KẾT QUẢ', 10.5, True, False, PP_ALIGN.CENTER),
    ('VÀ TÁC ĐỘNG', 10.5, True, False, PP_ALIGN.CENTER),
    ('', 5),
    ('• Thúc đẩy đổi mới sáng tạo', 9, False, False, PP_ALIGN.LEFT),
    ('• Phát triển KT số bền vững', 9, False, False, PP_ALIGN.LEFT),
    ('• Bảo vệ quyền lợi chủ thể', 9, False, False, PP_ALIGN.LEFT),
    ('  sáng tạo', 9, False, False, PP_ALIGN.LEFT),
    ('• Nâng cao NL cạnh tranh QG', 9, False, False, PP_ALIGN.LEFT),
    ('• Thu hút đầu tư & CGCN', 9, False, False, PP_ALIGN.LEFT),
], anchor='top')

# ═══════════════════════════════════════════════════════════
# 3.6 ĐẶC TRƯNG KTS - Ẩn box riêng (đã chuyển vào nhóm bên ngoài)
# ═══════════════════════════════════════════════════════════
hide_shape(shapes[20])

# ═══════════════════════════════════════════════════════════
# PHẢN HỒI CHÍNH SÁCH - cập nhật text
# ═══════════════════════════════════════════════════════════
# Tìm shape PHAN HOI (shape 24 trong file gốc, nhưng file này có 32 shapes)
for s in shapes:
    if hasattr(s, 'text') and 'PHẢN HỒI' in s.text.upper():
        set_text(s, [
            ('PHẢN HỒI CHÍNH SÁCH', 10.5, True, False, PP_ALIGN.CENTER),
            ('', 3),
            ('Thực tiễn thực thi  →  Đánh giá, tổng kết kết quả  →  Điều chỉnh chính sách  →  Hoàn thiện QLNN về quyền SHTT', 9, False, True, PP_ALIGN.CENTER, '444444'),
        ])
        break

# Ghi chú nguồn
for s in shapes:
    if hasattr(s, 'text') and 'Nguồn' in s.text:
        set_text(s, [
            ('Nguồn: Tác giả tổng hợp và đề xuất', 8, False, True, PP_ALIGN.CENTER, '777777'),
        ])
        break

# ═══════════════════════════════════════════════════════════
prs.save(DST)
print(f'[OK] {DST}')
print('Đã áp dụng 6 điểm góp ý:')
print('  3.1 Tiêu đề làm rõ khái niệm trung tâm QLNN về quyền SHTT')
print('  3.2 HIỆU QUẢ → CHẤT LƯỢNG QLNN VỀ QUYỀN SHTT')
print('  3.3 Bổ sung đủ 5 tiêu chí: Hiệu lực, Hiệu quả, Phù hợp, Đồng bộ, Minh bạch')
print('  3.4 NỘI DUNG sửa thành 4 nội dung đúng phạm vi luận án')
print('  3.5 YẾU TỐ phân 2 nhóm: Bên trong (4) + Bên ngoài (4 + Hội nhập QT)')
print('  3.6 ĐẶC TRƯNG KTS chuyển vào nhóm yếu tố bên ngoài, ẩn box riêng')
