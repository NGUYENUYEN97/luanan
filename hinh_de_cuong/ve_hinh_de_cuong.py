# Vẽ Hình 1 (Khung phân tích) và Hình 2 (Mô hình nghiên cứu) cho đề cương, bản đen trắng
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

plt.rcParams["font.family"] = "Times New Roman"
OUT = os.path.dirname(os.path.abspath(__file__))


def box(ax, x, y, w, h, text, fs=10.5, bold=False, lw=1.2, dashed=False):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.01,rounding_size=0.015",
                                fc="white", ec="black", lw=lw, ls="--" if dashed else "-"))
    ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs,
            fontweight="bold" if bold else "normal", wrap=True, linespacing=1.25)


def arrow(ax, x1, y1, x2, y2, label=None, lx=0, ly=0.012, dashed=False):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=13,
                                 lw=1.1, color="black", ls="--" if dashed else "-"))
    if label:
        ax.text((x1 + x2) / 2 + lx, (y1 + y2) / 2 + ly, label, ha="center", va="bottom",
                fontsize=10, fontstyle="italic")


# ---------------- Hình 1. Khung phân tích ----------------
fig, ax = plt.subplots(figsize=(10, 6.6))
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

box(ax, 0.02, 0.86, 0.96, 0.11,
    "CƠ SỞ LÝ THUYẾT\nLý thuyết quyền tài sản · Lý thuyết thất bại thị trường · "
    "Lý thuyết thể chế · Lý thuyết quản lý công mới và quản trị thời đại số", fs=10.5)

box(ax, 0.02, 0.30, 0.25, 0.46,
    "YẾU TỐ ẢNH HƯỞNG\n\nBên trong bộ máy nhà nước:\nthể chế, năng lực cán bộ,\nnguồn lực, hạ tầng dữ liệu\n\n"
    "Bên ngoài:\ncông nghệ số và mô hình\nkinh doanh số; năng lực\nchủ thể quyền; thị trường\nvà dịch vụ trung gian;\nhội nhập quốc tế", fs=9.8)

box(ax, 0.33, 0.30, 0.32, 0.46,
    "NỘI DUNG QLNN ĐỐI VỚI\nKHAI THÁC QUYỀN SHTT\n\n(1) Chiến lược, chính sách,\npháp luật về khai thác quyền\n\n"
    "(2) Tổ chức bộ máy và\nphối hợp liên ngành\n\n(3) Hỗ trợ khai thác và phát triển\nthị trường quyền SHTT\n\n"
    "(4) Kiểm tra, giám sát và bảo vệ\ngiá trị khai thác trên môi trường số", fs=9.8)

box(ax, 0.71, 0.30, 0.27, 0.46,
    "KẾT QUẢ KHAI THÁC\nQUYỀN SHTT\n\nSử dụng trong sản xuất,\nkinh doanh; chuyển nhượng;\nli-xăng; góp vốn, thế chấp;\n"
    "thương mại hóa kết quả\nnghiên cứu; khai thác qua\nnền tảng số\n\n(đo bằng số liệu thống kê\nvà đánh giá của các đối tượng\nkhảo sát)", fs=9.8)

box(ax, 0.33, 0.13, 0.32, 0.10,
    "TIÊU CHÍ ĐÁNH GIÁ QLNN\nHiệu lực · Hiệu quả · Phù hợp và thích ứng\nĐồng bộ · Minh bạch", fs=9.8)

box(ax, 0.02, 0.01, 0.96, 0.08,
    "QUAN ĐIỂM, ĐỊNH HƯỚNG VÀ GIẢI PHÁP HOÀN THIỆN QLNN VỀ QUYỀN SHTT TRONG NỀN KINH TẾ SỐ\n"
    "(trọng tâm: thúc đẩy khai thác quyền SHTT) đến năm 2030, tầm nhìn đến năm 2045", fs=10)

arrow(ax, 0.50, 0.86, 0.50, 0.765)
arrow(ax, 0.27, 0.53, 0.33, 0.53)
arrow(ax, 0.65, 0.53, 0.71, 0.53)
arrow(ax, 0.49, 0.30, 0.49, 0.23)
arrow(ax, 0.845, 0.30, 0.65, 0.20, dashed=True)
arrow(ax, 0.49, 0.13, 0.49, 0.09)
fig.savefig(os.path.join(OUT, "Hinh1_KhungPhanTich.png"), dpi=220, bbox_inches="tight")
plt.close(fig)

# ---------------- Hình 2. Mô hình nghiên cứu ----------------
fig, ax = plt.subplots(figsize=(10, 5.6))
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")

# Khối biến độc lập: 4 nội dung QLNN theo đánh giá của doanh nghiệp
box(ax, 0.005, 0.02, 0.35, 0.93, "", lw=1.0, dashed=True)
ax.text(0.18, 0.915, "NỘI DUNG QLNN\n(theo đánh giá của doanh nghiệp)", ha="center", va="top",
        fontsize=10, fontweight="bold")
ivs = [("H1: Chính sách, pháp luật về\nkhai thác quyền SHTT (CSPL)", 0.62),
       ("H2: Tổ chức bộ máy và\nphối hợp liên ngành (BMPH)", 0.44),
       ("H3: Hỗ trợ khai thác và hạ tầng\nthị trường quyền SHTT (HTTT)", 0.26),
       ("H4: Kiểm tra, giám sát, bảo vệ quyền\ntrên môi trường số (KTGS)", 0.08)]
for t, y in ivs:
    box(ax, 0.025, y, 0.31, 0.14, t, fs=9.8)

box(ax, 0.47, 0.70, 0.24, 0.17, "Năng lực quản trị\nSHTT của doanh nghiệp\n(NLQT)", fs=10, bold=True)
box(ax, 0.76, 0.36, 0.23, 0.17, "Kết quả khai thác\nquyền SHTT của\ndoanh nghiệp (KQKT)", fs=10, bold=True)
box(ax, 0.47, 0.04, 0.52, 0.15,
    "Biến kiểm soát: quy mô, ngành, tuổi doanh nghiệp, hình thức sở hữu,\nmức độ số hóa, hoạt động nghiên cứu và phát triển",
    fs=9.5, dashed=True)

arrow(ax, 0.355, 0.445, 0.76, 0.445, label="H1, H2, H3, H4 (tác động trực tiếp)", lx=0.0, ly=0.012)
arrow(ax, 0.355, 0.785, 0.47, 0.785, label="a", ly=0.01)
arrow(ax, 0.71, 0.785, 0.875, 0.53, label="b", lx=0.03, ly=0.0)
ax.text(0.59, 0.905, "H5: NLQT làm trung gian (a × b)", ha="center", fontsize=10, fontstyle="italic")
arrow(ax, 0.875, 0.19, 0.875, 0.36, dashed=True)
fig.savefig(os.path.join(OUT, "Hinh2_MoHinhNghienCuu.png"), dpi=220, bbox_inches="tight")
plt.close(fig)
print("OK")
