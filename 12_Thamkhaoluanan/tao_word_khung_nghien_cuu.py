"""
KHUNG NGHIÊN CỨU v8.0 - KẾT HỢP DỌC + NGANG
Tổng thể: Dọc (top-down)
4 hộp chính: Ngang (left-right) với mũi tên →
Phản hồi: Vòng U từ KQ về YT

Layout:
  [LÝ THUYẾT NỀN TẢNG      ← full width, ngang]
              ↓
  [YẾU TỐ] → [NỘI DUNG] → [HIỆU QUẢ] → [KẾT QUẢ]  ← ngang
                                                 ↓
  [PHẢN HỒI CHÍNH SÁCH     ← full width, ngang]
       ↑___________________________________________|
"""
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
import os

# ─────────────── TIỆN ÍCH ───────────────
def fmt(run, sz=10, bold=False, italic=False, color='000000'):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(sz)
    run.font.bold = bold; run.font.italic = italic
    run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._element.get_or_add_rPr()
    rPr.insert(0, parse_xml(
        r'<w:rFonts {} w:eastAsia="Times New Roman" w:cs="Times New Roman"/>'.format(nsdecls('w'))))

def write_cell(cell, lines, valign=WD_ALIGN_VERTICAL.CENTER):
    """lines = (text, sz, bold, italic, align, color)"""
    for i, ln in enumerate(lines):
        t, sz = ln[0], ln[1]
        b = ln[2] if len(ln)>2 else False
        it= ln[3] if len(ln)>3 else False
        al= ln[4] if len(ln)>4 else WD_ALIGN_PARAGRAPH.CENTER
        co= ln[5] if len(ln)>5 else '000000'
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        if i == 0: p.text = ''
        run = p.add_run(t); fmt(run, sz, b, it, co)
        p.alignment = al
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after  = Pt(1)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    cell.vertical_alignment = valign

def set_bg(cell, hex_color):
    cell._tc.get_or_add_tcPr().append(
        parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), hex_color)))

def set_border(cell, top='none', bot='none', left='none', right='none',
               val='single', sz='8', color='222222'):
    def edge(name, v):
        if v == 'on':
            return '<w:{n} w:val="{v}" w:sz="{s}" w:space="0" w:color="{c}"/>'.format(
                n=name, v=val, s=sz, c=color)
        return '<w:{n} w:val="none" w:sz="0" w:space="0" w:color="FFFFFF"/>'.format(n=name)
    xml = '<w:tcBorders {ns}>{t}{b}{l}{r}</w:tcBorders>'.format(
        ns=nsdecls('w'),
        t=edge('top', top), b=edge('bottom', bot),
        l=edge('left', left), r=edge('right', right))
    cell._tc.get_or_add_tcPr().append(parse_xml(xml))

def all_borders(cell, val='single', sz='10', color='222222'):
    set_border(cell, 'on','on','on','on', val, sz, color)

def no_border(cell):
    set_border(cell)

def top_bot_border(cell, val='dashed', sz='6', color='888888'):
    set_border(cell, top='on', bot='on', left='none', right='none', val=val, sz=sz, color=color)

def set_pad(cell, t=40, b=40, l=60, r=60):
    cell._tc.get_or_add_tcPr().append(parse_xml(
        '<w:tcMar {ns}>'
        '<w:top w:w="{t}" w:type="dxa"/>'
        '<w:bottom w:w="{b}" w:type="dxa"/>'
        '<w:start w:w="{l}" w:type="dxa"/>'
        '<w:end w:w="{r}" w:type="dxa"/>'
        '</w:tcMar>'.format(ns=nsdecls('w'), t=t, b=b, l=l, r=r)))

def set_row_h(row, cm):
    row._tr.get_or_add_trPr().append(parse_xml(
        '<w:trHeight {} w:val="{}" w:hRule="atLeast"/>'.format(
            nsdecls('w'), int(cm*567))))

def add_p(doc, text, sz=12, bold=False, italic=False,
          align=WD_ALIGN_PARAGRAPH.CENTER, sb=0, sa=2, color='000000'):
    p = doc.add_paragraph()
    run = p.add_run(text); fmt(run, sz, bold, italic, color)
    p.alignment = align
    p.paragraph_format.space_before = Pt(sb)
    p.paragraph_format.space_after  = Pt(sa)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    p.paragraph_format.line_spacing = 1.5

# ─────────────── TÀI LIỆU ───────────────
doc = Document()
for s in doc.sections:
    s.page_width = Cm(21); s.page_height = Cm(29.7)
    s.top_margin = Cm(2.5); s.bottom_margin = Cm(2.0)
    s.left_margin = Cm(3.0); s.right_margin = Cm(2.0)
doc.styles['Normal'].font.name = 'Times New Roman'
doc.styles['Normal'].font.size = Pt(12)

add_p(doc, 'Sơ đồ 2.x. Khung phân tích của luận án',
      13, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, sa=8)

# ═══════════════════════════════════════════════════════
# BẢNG 7 CỘT × 9 HÀNG
#
# Cột:  0=YT   1=→   2=ND   3=→   4=HQ   5=→   6=KQ
# Chiều rộng: 3.4 | 0.75 | 3.4 | 0.75 | 3.1 | 0.75 | 3.05 = 15.2cm
#
# Hàng 0: [─ BỐI CẢNH: NỀN KINH TẾ SỐ (merge all) ─]
# Hàng 1: [─ LÝ THUYẾT NỀN TẢNG (merge all) ─]
# Hàng 2: [─── ↓ ─── (merge all) ───]
# Hàng 3: [YT] [→] [ND] [→] [HQ] [→] [KQ]   ← 4 hộp ngang
# Hàng 4: [ ] [ ] [ ] [ ] [ ] [ ] [↓]
# Hàng 5: [↑] [── PHẢN HỒI CHÍNH SÁCH (merge 1-6) ──]
# Hàng 6: [─ BỐI CẢNH FOOTER (merge all) ─]
# ═══════════════════════════════════════════════════════

NCOLS = 7
NROWS = 7
COL_W = [Cm(3.4), Cm(0.75), Cm(3.4), Cm(0.75), Cm(3.1), Cm(0.75), Cm(3.05)]

tbl = doc.add_table(rows=NROWS, cols=NCOLS)
tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

for row in tbl.rows:
    for j, w in enumerate(COL_W):
        row.cells[j].width = w

# Xóa nội dung và viền mặc định
for i in range(NROWS):
    for j in range(NCOLS):
        no_border(tbl.cell(i, j))
        tbl.cell(i, j).paragraphs[0].text = ''

# ── HÀNG 0: TIÊU ĐỀ BỐI CẢNH ──
r0 = tbl.cell(0, 0).merge(tbl.cell(0, NCOLS-1))
set_border(r0, top='on', left='on', right='on', val='dashed', sz='6', color='888888')
set_bg(r0, 'F7F7F7')
set_pad(r0, 18, 8, 60, 60)
write_cell(r0, [('BỐI CẢNH: NỀN KINH TẾ SỐ', 9, True, True, WD_ALIGN_PARAGRAPH.LEFT, '666666')])
set_row_h(tbl.rows[0], 0.45)

# ── HÀNG 1: LÝ THUYẾT NỀN TẢNG ──
r1 = tbl.cell(1, 0).merge(tbl.cell(1, NCOLS-1))
all_borders(r1, sz='8', color='444444')
set_bg(r1, 'EBEBEB')
set_pad(r1, 35, 35, 80, 80)
write_cell(r1, [
    ('LÝ THUYẾT NỀN TẢNG', 11, True, False),
    ('', 3, False, False),
    ('Lý thuyết quản lý công mới  •  Lý thuyết thể chế', 9, False, False, WD_ALIGN_PARAGRAPH.CENTER, '333333'),
    ('Lý thuyết quản trị số  •  Lý thuyết quyền sở hữu trí tuệ', 9, False, False, WD_ALIGN_PARAGRAPH.CENTER, '333333'),
])
set_row_h(tbl.rows[1], 1.6)

# ── HÀNG 2: MŨI TÊN ↓ ──
r2 = tbl.cell(2, 0).merge(tbl.cell(2, NCOLS-1))
set_border(r2, left='on', right='on', val='dashed', sz='6', color='888888')
set_bg(r2, 'F7F7F7')
write_cell(r2, [('▼', 13, True, False, WD_ALIGN_PARAGRAPH.CENTER, '444444')])
set_row_h(tbl.rows[2], 0.5)

# ── HÀNG 3: BỐN HỘP NGANG ──
# Cột 0: YẾU TỐ ẢNH HƯỞNG
c_yt = tbl.cell(3, 0)
all_borders(c_yt, sz='12', color='1a1a1a')
set_pad(c_yt, 35, 35, 50, 50)
write_cell(c_yt, [
    ('CÁC YẾU TỐ', 10.5, True, False),
    ('ẢNH HƯỞNG', 10.5, True, False),
    ('', 4, False, False),
    ('(1) Khung pháp lý', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     và chính sách', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(2) Năng lực tổ chức', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     và cán bộ', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(3) Hạ tầng công nghệ số', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(4) Nguồn lực tài chính', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(5) Phối hợp thể chế', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(6) Nhận thức xã hội', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     về SHTT', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
])

# Cột 1: mũi tên →
c_a1 = tbl.cell(3, 1)
no_border(c_a1)
set_bg(c_a1, 'FFFFFF')
write_cell(c_a1, [('→', 16, True, False, WD_ALIGN_PARAGRAPH.CENTER, '333333')])

# Cột 2: NỘI DUNG QLNN
c_nd = tbl.cell(3, 2)
all_borders(c_nd, sz='12', color='1a1a1a')
set_pad(c_nd, 35, 35, 50, 50)
write_cell(c_nd, [
    ('NỘI DUNG QLNN', 10.5, True, False),
    ('VỀ QUYỀN SHTT', 10.5, True, False),
    ('', 4, False, False),
    ('(1) Hoạch định chính', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     sách, chiến lược', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(2) Hoàn thiện pháp luật', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(3) Tổ chức bộ máy', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     quản lý', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(4) Thanh tra, kiểm tra', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     xử lý vi phạm', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(5) Ứng dụng công nghệ', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     số trong quản lý', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
])

# Cột 3: mũi tên →
c_a2 = tbl.cell(3, 3)
no_border(c_a2)
write_cell(c_a2, [('→', 16, True, False, WD_ALIGN_PARAGRAPH.CENTER, '333333')])

# Cột 4: HIỆU QUẢ QLNN
c_hq = tbl.cell(3, 4)
all_borders(c_hq, sz='12', color='1a1a1a')
set_pad(c_hq, 35, 35, 50, 50)
write_cell(c_hq, [
    ('HIỆU QUẢ QLNN', 10.5, True, False),
    ('VỀ QUYỀN SHTT', 10.5, True, False),
    ('', 4, False, False),
    ('(1) Hiệu lực', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     thực thi PL', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(2) Bảo vệ quyền SHTT', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     trên không gian số', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(3) Minh bạch,', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     công khai', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(4) Hài lòng của', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     chủ thể quyền', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(5) Thích ứng với', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     công nghệ mới', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
])

# Cột 5: mũi tên →
c_a3 = tbl.cell(3, 5)
no_border(c_a3)
write_cell(c_a3, [('→', 16, True, False, WD_ALIGN_PARAGRAPH.CENTER, '333333')])

# Cột 6: KẾT QUẢ, TÁC ĐỘNG
c_kq = tbl.cell(3, 6)
all_borders(c_kq, sz='12', color='1a1a1a')
set_pad(c_kq, 35, 35, 50, 50)
write_cell(c_kq, [
    ('KẾT QUẢ,', 10.5, True, False),
    ('TÁC ĐỘNG', 10.5, True, False),
    ('', 4, False, False),
    ('(1) Thúc đẩy', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     đổi mới sáng tạo', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(2) Phát triển', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     kinh tế số BV', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(3) Bảo vệ quyền lợi', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     chủ thể sáng tạo', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(4) Nâng cao NL', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     cạnh tranh QG', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('(5) Thu hút đầu tư', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
    ('     và CGCN', 9, False, False, WD_ALIGN_PARAGRAPH.LEFT),
])
set_row_h(tbl.rows[3], 5.5)

# ── HÀNG 4: VÒNG PHẢN HỒI (↓ dưới KQ, ↑ dưới YT) ──
# Tạo hàng để tạo hình U cho phản hồi
for j in range(NCOLS):
    no_border(tbl.cell(4, j))
    set_bg(tbl.cell(4, j), 'FFFFFF')

# Cột 0: mũi tên ↑ (đầu mũi tên phản hồi lên YT)
write_cell(tbl.cell(4, 0), [
    ('▲', 12, True, False, WD_ALIGN_PARAGRAPH.CENTER, '555555'),
    ('phản hồi', 7, False, True, WD_ALIGN_PARAGRAPH.CENTER, '777777'),
])
# Cột 1-5: đường ngang của phản hồi
c_mid = tbl.cell(4, 1).merge(tbl.cell(4, 5))
write_cell(c_mid, [('◄ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─', 9, False, False, WD_ALIGN_PARAGRAPH.CENTER, '666666')])
# Cột 6: mũi tên ↓ (từ KQ xuống vòng phản hồi)
write_cell(tbl.cell(4, 6), [('▼', 12, True, False, WD_ALIGN_PARAGRAPH.CENTER, '555555')])
set_row_h(tbl.rows[4], 0.55)

# ── HÀNG 5: PHẢN HỒI CHÍNH SÁCH ──
r5 = tbl.cell(5, 0).merge(tbl.cell(5, NCOLS-1))
all_borders(r5, val='dashed', sz='8', color='555555')
set_bg(r5, 'F0F0F0')
set_pad(r5, 35, 35, 80, 80)
write_cell(r5, [
    ('PHẢN HỒI CHÍNH SÁCH', 11, True, False),
    ('', 3, False, False),
    ('Thực tiễn thực thi  →  Đánh giá, tổng kết kết quả  →  Điều chỉnh chính sách  →  Hoàn thiện QLNN về SHTT',
     9.5, False, True, WD_ALIGN_PARAGRAPH.CENTER, '333333'),
])
set_row_h(tbl.rows[5], 1.5)

# ── HÀNG 6: ĐÓNG KHUNG BỐI CẢNH ──
r6 = tbl.cell(6, 0).merge(tbl.cell(6, NCOLS-1))
set_border(r6, bot='on', left='on', right='on', val='dashed', sz='6', color='888888')
set_bg(r6, 'F7F7F7')
write_cell(r6, [('', 3, False, False)])
set_row_h(tbl.rows[6], 0.2)

# ─────────────── GHI CHÚ ───────────────
add_p(doc, 'Nguồn: Tác giả tự xây dựng trên cơ sở tổng quan lý thuyết',
      11, italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, sb=6)
add_p(doc, '→ Chiều tác động trực tiếp    ◄── Vòng phản hồi chính sách    ▼ Từ trên xuống dưới',
      8.5, italic=True, align=WD_ALIGN_PARAGRAPH.LEFT, color='555555', sa=0)

# ─────────────── LƯU ───────────────
out = r"g:\My Drive\Luan an uyen 26\LuanAn_TS_SHTT_KinhTeSo\12_Thamkhaoluanan\KhungNghienCuu_LuanAn_v8.docx"
doc.save(out)
print(f"[OK] {out}")
print("  Bo cuc: doc (top-down) + ngang (4 hop song song)")
print("  Vong phan hoi hinh U: KQ -> xuong -> trai -> len -> YT")
