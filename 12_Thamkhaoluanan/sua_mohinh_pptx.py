"""
Sửa file MoHinhNghienCuu_TichHop_HoanHao.pptx
Giữ nguyên toàn bộ vị trí, kích thước, màu sắc, viền
CHỈ thay đổi NỘI DUNG TEXT trong các hộp:
  - Bỏ tất cả từ "Biến" (biến độc lập, phụ thuộc, trung gian, điều tiết, kiểm soát)
  - Dùng ngôn ngữ mô tả phù hợp luận án Việt Nam
"""
import sys; sys.stdout.reconfigure(encoding='utf-8')
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import copy

SRC = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\12_Thamkhaoluanan\MoHinhNghienCuu_TichHop_HoanHao.pptx"
DST = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\12_Thamkhaoluanan\KhungPhanTich_LuanAn_Final.pptx"

prs = Presentation(SRC)
slide = prs.slides[0]

def rgb(h): return RGBColor.from_string(h)

def set_text(shape, lines, center_title=True):
    """
    lines = list of (text, size_pt, bold, italic, align)
    Xóa sạch text cũ, viết mới giữ nguyên định dạng hộp
    """
    tf = shape.text_frame
    tf.word_wrap = True

    # Xóa các paragraph cũ (trừ paragraph 0)
    while len(tf.paragraphs) > 1:
        p_elem = tf.paragraphs[-1]._p
        p_elem.getparent().remove(p_elem)

    for i, ln in enumerate(lines):
        txt  = ln[0]
        sz   = ln[1]
        bold = ln[2] if len(ln) > 2 else False
        ital = ln[3] if len(ln) > 3 else False
        algn = ln[4] if len(ln) > 4 else (PP_ALIGN.CENTER if center_title and i==0 else PP_ALIGN.LEFT)
        co   = ln[5] if len(ln) > 5 else '000000'

        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = algn
        p.space_before = Pt(1)
        p.space_after  = Pt(1)

        # Xóa run cũ trong paragraph 0
        if i == 0:
            for run in list(p.runs):
                p._p.remove(run._r)

        if txt == '':
            p.space_before = Pt(sz)
            continue

        run = p.add_run()
        run.text = txt
        run.font.name  = 'Times New Roman'
        run.font.size  = Pt(sz)
        run.font.bold  = bold
        run.font.italic = ital
        run.font.color.rgb = rgb(co)

# ═══════════════════════════════════════════════════════════
# SỬA TỪNG SHAPE
# ═══════════════════════════════════════════════════════════

shapes = slide.shapes

# [0] TIÊU ĐỀ - Text box
set_text(shapes[0], [
    ('KHUNG PHÂN TÍCH', 14, True, False, PP_ALIGN.LEFT),
    ('Quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số', 11, False, True, PP_ALIGN.LEFT, '444444'),
])

# [1] LÝ THUYẾT NỀN TẢNG - giữ title, cập nhật nội dung
set_text(shapes[1], [
    ('LÝ THUYẾT NỀN TẢNG', 11, True, False, PP_ALIGN.CENTER),
    ('', 3),
    ('Lý thuyết quản lý công mới  •  Lý thuyết thể chế  •  Lý thuyết quản trị số  •  Lý thuyết quyền sở hữu trí tuệ', 9, False, False, PP_ALIGN.CENTER, '333333'),
])

# [2] OUTER BOX - CÁC YẾU TỐ ẢNH HƯỞNG (thay CÁC BIẾN ĐỘC LẬP)
set_text(shapes[2], [
    ('CÁC YẾU TỐ ẢNH HƯỞNG', 10.5, True, False, PP_ALIGN.CENTER),
])

# [3]-[8] Các hộp yếu tố nhỏ (giữ nguyên, chỉ thêm số thứ tự)
factor_labels = [
    '(1) Khung pháp lý\nvà chính sách',
    '(2) Năng lực tổ chức\nvà cán bộ',
    '(3) Nguồn lực\ntài chính',
    '(4) Hạ tầng\ncông nghệ số',
    '(5) Phối hợp\nthể chế',
    '(6) Nhận thức xã hội\nvề quyền SHTT',
]
for i, label in enumerate(factor_labels):
    lines_parts = label.split('\n')
    lines = [(p, 9.5, False, False, PP_ALIGN.CENTER) for p in lines_parts]
    set_text(shapes[3 + i], lines, center_title=True)

# [9] BIẾN TRUNG GIAN → NỘI DUNG QLNN VỀ QUYỀN SHTT
set_text(shapes[9], [
    ('NỘI DUNG QLNN', 10.5, True),
    ('VỀ QUYỀN SHTT', 10.5, True),
    ('', 4),
    ('(1) Hoạch định chính sách,\n     chiến lược quyền SHTT', 9, False, False, PP_ALIGN.LEFT),
    ('(2) Xây dựng và hoàn thiện\n     hệ thống pháp luật', 9, False, False, PP_ALIGN.LEFT),
    ('(3) Tổ chức bộ máy quản lý', 9, False, False, PP_ALIGN.LEFT),
    ('(4) Thanh tra, kiểm tra,\n     xử lý vi phạm', 9, False, False, PP_ALIGN.LEFT),
    ('(5) Ứng dụng công nghệ số\n     trong quản lý quyền SHTT', 9, False, False, PP_ALIGN.LEFT),
])

# [10] BIẾN PHỤ THUỘC → HIỆU QUẢ QLNN VỀ QUYỀN SHTT
set_text(shapes[10], [
    ('HIỆU QUẢ QLNN', 10.5, True),
    ('VỀ QUYỀN SHTT', 10.5, True),
    ('', 4),
    ('(1) Hiệu lực thực thi\n     pháp luật về quyền SHTT', 9, False, False, PP_ALIGN.LEFT),
    ('(2) Bảo vệ quyền SHTT\n     trên không gian số', 9, False, False, PP_ALIGN.LEFT),
    ('(3) Tính minh bạch,\n     công khai', 9, False, False, PP_ALIGN.LEFT),
    ('(4) Mức độ hài lòng\n     của chủ thể quyền SHTT', 9, False, False, PP_ALIGN.LEFT),
    ('(5) Khả năng thích ứng\n     với công nghệ mới', 9, False, False, PP_ALIGN.LEFT),
])

# [11] TÁC ĐỘNG VĨ MÔ → KẾT QUẢ VÀ TÁC ĐỘNG
set_text(shapes[11], [
    ('KẾT QUẢ', 10.5, True),
    ('VÀ TÁC ĐỘNG', 10.5, True),
    ('', 4),
    ('• Thúc đẩy đổi mới sáng tạo', 9, False, False, PP_ALIGN.LEFT),
    ('• Phát triển KT số bền vững', 9, False, False, PP_ALIGN.LEFT),
    ('• Bảo vệ quyền lợi chủ thể\n  sáng tạo', 9, False, False, PP_ALIGN.LEFT),
    ('• Nâng cao NL cạnh tranh QG', 9, False, False, PP_ALIGN.LEFT),
    ('• Thu hút đầu tư & CGCN', 9, False, False, PP_ALIGN.LEFT),
])

# [20] BIẾN ĐIỀU TIẾT → ĐẶC TRƯNG NỀN KINH TẾ SỐ
set_text(shapes[20], [
    ('ĐẶC TRƯNG', 10, True),
    ('NỀN KINH TẾ SỐ', 10, True),
    ('', 4),
    ('• Hạ tầng số quốc gia', 9, False, False, PP_ALIGN.LEFT),
    ('• Mức độ chuyển đổi số', 9, False, False, PP_ALIGN.LEFT),
    ('• Thương mại điện tử', 9, False, False, PP_ALIGN.LEFT),
    ('• Tài sản số và nền tảng số', 9, False, False, PP_ALIGN.LEFT),
])

# [22] BIẾN KIỂM SOÁT → ẩn/xóa nội dung (giữ hộp trống)
# Thay bằng ghi chú
set_text(shapes[22], [
    ('Nguồn: Tác giả tổng hợp', 8, False, True, PP_ALIGN.CENTER, '777777'),
])
# Ẩn viền hộp này
try:
    shapes[22].line.color.rgb = rgb('FFFFFF')  # trắng = ẩn
    shapes[22].fill.solid()
    shapes[22].fill.fore_color.rgb = rgb('FFFFFF')  # nền trắng
except: pass

# [24] PHẢN HỒI CHÍNH SÁCH - giữ, bỏ (FEEDBACK)
set_text(shapes[24], [
    ('PHẢN HỒI CHÍNH SÁCH', 10.5, True, False, PP_ALIGN.CENTER),
    ('', 3),
    ('Thực tiễn thực thi  →  Đánh giá, tổng kết kết quả  →  Điều chỉnh chính sách  →  Hoàn thiện QLNN về quyền SHTT', 9, False, True, PP_ALIGN.CENTER, '444444'),
])

# ═══════════════════════════════════════════════════════════
prs.save(DST)
print(f'[OK] Lưu: {DST}')
print('Các thay đổi:')
print('  [0] Tiêu đề: Khung phân tích QLNN về quyền SHTT trong KTS')
print('  [2] CÁC BIẾN ĐỘC LẬP → CÁC YẾU TỐ ẢNH HƯỞNG')
print('  [3-8] Giữ các yếu tố, thêm số thứ tự')
print('  [9] BIẾN TRUNG GIAN → NỘI DUNG QLNN VỀ QUYỀN SHTT')
print('  [10] BIẾN PHỤ THUỘC → HIỆU QUẢ QLNN VỀ QUYỀN SHTT')
print('  [11] TÁC ĐỘNG VĨ MÔ → KẾT QUẢ VÀ TÁC ĐỘNG')
print('  [20] BIẾN ĐIỀU TIẾT → ĐẶC TRƯNG NỀN KINH TẾ SỐ')
print('  [22] BIẾN KIỂM SOÁT → ẩn (ghi chú nguồn)')
print('  [24] PHẢN HỒI CHÍNH SÁCH - cập nhật text')
