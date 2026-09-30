# Vẽ Hình: Quy trình thực hiện nghiên cứu (bố cục ma trận 4 giai đoạn, bản đen trắng)
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Polygon

plt.rcParams["font.family"] = "Times New Roman"
OUT = os.path.dirname(os.path.abspath(__file__))

fig, ax = plt.subplots(figsize=(13, 8.6))
ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")


def box(x, y, w, h, text, fs=9, bold=False, fc="white", lw=1.0, ls="-", italic=False, ha="center"):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.002,rounding_size=0.008",
                                fc=fc, ec="black", lw=lw, ls=ls))
    tx = x + w / 2 if ha == "center" else x + 0.006
    ax.text(tx, y + h / 2, text, ha=ha, va="center", fontsize=fs, linespacing=1.2,
            fontweight="bold" if bold else "normal", fontstyle="italic" if italic else "normal")


def down(x, y1, y2):
    ax.add_patch(FancyArrowPatch((x, y1), (x, y2), arrowstyle="-|>", mutation_scale=9, lw=0.9, color="black"))


# Cột nhãn hàng bên trái
LX, LW = 0.005, 0.092
rows = {  # tên hàng: (y, h)
    "GIAI ĐOẠN": (0.885, 0.10),
    "PHƯƠNG PHÁP": (0.815, 0.055),
    "CÁC BƯỚC\nTHỰC HIỆN": (0.300, 0.500),
    "KẾT QUẢ\nDỰ KIẾN": (0.105, 0.180),
    "CHƯƠNG\nLUẬN ÁN": (0.015, 0.075),
}
for name, (y, h) in rows.items():
    box(LX, y, LW, h, name, fs=8.7, bold=True, fc="#d9d9d9")

stages = [
    dict(title="Giai đoạn 1\nXÂY DỰNG CƠ SỞ LÝ LUẬN\nVÀ KHUNG PHÂN TÍCH",
         method="Tổng quan tài liệu;\nphân tích văn bản chính sách",
         steps=["1.1. Xác định phạm vi, từ khóa\nvà tiêu chí chọn tài liệu",
                "1.2. Tra cứu, sàng lọc công trình\n(Scopus, WoS, luận án, báo cáo\nWIPO, OECD) và văn bản pháp luật",
                "1.3. Hệ thống hóa theo 3 nhóm:\nSHTT; kinh tế số; khai thác quyền\nSHTT trong kinh tế số",
                "1.4. Xây dựng khung phân tích,\ntiêu chí đánh giá và\nphiếu khảo sát nháp"],
         result="• Khoảng trống nghiên cứu ở\n  khâu khai thác quyền SHTT\n• Khung phân tích: 4 nội dung,\n  5 tiêu chí, 2 nhóm yếu tố\n• Câu hỏi, giả thuyết và\n  phiếu khảo sát nháp",
         chap="Chương 1, Chương 2"),
    dict(title="Giai đoạn 2\nPHỎNG VẤN CHUYÊN GIA\nVÀ HOÀN THIỆN PHIẾU",
         method="Phỏng vấn sâu bán cấu trúc;\nkhảo sát thử",
         steps=["2.1. Lựa chọn 15-20 người phỏng vấn:\nchuyên gia, cán bộ quản lý, tổ chức\ntrung gian, doanh nghiệp",
                "2.2. Thiết kế hướng dẫn phỏng vấn\ntheo 4 nội dung QLNN",
                "2.3. Phỏng vấn, ghi chép;\ntổng hợp ý kiến theo chủ đề",
                "2.4. Khảo sát thử 15-20 người;\nchỉnh sửa và hoàn thiện\nphiếu khảo sát chính thức"],
         result="• Ý kiến chuyên gia về thực trạng,\n  nguyên nhân\n• Phiếu khảo sát chính thức\n  (thang Likert 5 mức)",
         chap="Mở đầu (phương pháp),\nChương 3"),
    dict(title="Giai đoạn 3\nKHẢO SÁT VÀ\nPHÂN TÍCH THỰC TRẠNG",
         method="Thống kê mô tả;\nphân tích số liệu thứ cấp",
         steps=["3.1. Khảo sát 150-180 phiếu\n(3 nhóm chủ thể) tại Hà Nội,\nTP. Hồ Chí Minh, Đà Nẵng",
                "3.2. Làm sạch, tổng hợp phiếu;\ntính tỷ lệ, điểm trung bình\nvà độ lệch chuẩn",
                "3.3. Đánh giá 4 nội dung QLNN\ntheo 5 tiêu chí; so sánh ý kiến\ngiữa các nhóm đối tượng",
                "3.4. Phân tích số liệu thứ cấp\n2019-2025; đối chiếu với kết quả\nkhảo sát và ý kiến chuyên gia"],
         result="• Đánh giá 4 nội dung QLNN\n  theo 5 tiêu chí\n• Các yếu tố ảnh hưởng\n• Kết quả, hạn chế và nguyên nhân",
         chap="Chương 3"),
    dict(title="Giai đoạn 4\nĐỀ XUẤT QUAN ĐIỂM,\nĐỊNH HƯỚNG VÀ GIẢI PHÁP",
         method="Phân tích tổng hợp, so sánh;\ntham vấn chuyên gia",
         steps=["4.1. Phân tích bối cảnh, xu hướng\nvà yêu cầu từ chủ trương,\npháp luật mới",
                "4.2. Tổng hợp kết quả Giai đoạn 3\nvới bài học kinh nghiệm quốc tế",
                "4.3. Xây dựng quan điểm,\nđịnh hướng và nhóm giải pháp\ntheo thứ tự ưu tiên",
                "4.4. Tham vấn chuyên gia về tính\nkhả thi; hoàn thiện giải pháp\nvà kiến nghị"],
         result="• Quan điểm, định hướng đến\n  năm 2030, tầm nhìn 2045\n• 5 nhóm giải pháp và điều kiện\n  thực hiện\n• Kiến nghị với cơ quan nhà nước",
         chap="Chương 4, Kết luận"),
]

X0, GAP = 0.105, 0.012
W = (1 - X0 - 0.005 - GAP * 3) / 4
for i, s in enumerate(stages):
    x = X0 + i * (W + GAP)
    # Tiêu đề giai đoạn (khung đậm, vát mép phải để thể hiện chiều tiến trình)
    y, h = rows["GIAI ĐOẠN"]
    notch = 0.012 if i < 3 else 0
    ax.add_patch(Polygon([(x, y), (x + W - notch, y), (x + W, y + h / 2), (x + W - notch, y + h), (x, y + h)],
                         closed=True, fc="white", ec="black", lw=1.6))
    ax.text(x + W / 2 - notch / 2, y + h / 2, s["title"], ha="center", va="center", fontsize=10.2,
            fontweight="bold", linespacing=1.2)
    # Phương pháp
    y, h = rows["PHƯƠNG PHÁP"]
    box(x, y, W, h, s["method"], fs=9.8, italic=True, fc="#efefef")
    # Các bước
    y0, h0 = rows["CÁC BƯỚC\nTHỰC HIỆN"]
    sh, sgap = 0.100, 0.033
    ys = [y0 + h0 - sh - k * (sh + sgap) for k in range(4)]
    for k, (yy, txt) in enumerate(zip(ys, s["steps"])):
        box(x + 0.006, yy, W - 0.012, sh, txt, fs=9.6)
        if k < 3:
            down(x + W / 2, yy, yy - sgap)
    down(x + W / 2, rows["PHƯƠNG PHÁP"][0], y0 + h0 - 0.0)
    # Kết quả
    y, h = rows["KẾT QUẢ\nDỰ KIẾN"]
    down(x + W / 2, ys[-1], y + h)
    box(x, y, W, h, s["result"], fs=9.6, fc="white", lw=1.4, ls="--", ha="left")
    # Chương
    y, h = rows["CHƯƠNG\nLUẬN ÁN"]
    box(x, y, W, h, s["chap"], fs=10, bold=True, fc="#d9d9d9")
    # Mũi tên nối kết quả giai đoạn trước sang giai đoạn sau
    if i < 3:
        yk = rows["KẾT QUẢ\nDỰ KIẾN"][0] + rows["KẾT QUẢ\nDỰ KIẾN"][1] / 2
        ax.add_patch(FancyArrowPatch((x + W, yk), (x + W + GAP, yk), arrowstyle="-|>",
                                     mutation_scale=9, lw=0.9, color="black"))

fig.savefig(os.path.join(OUT, "Hinh_QuyTrinhNghienCuu.png"), dpi=220, bbox_inches="tight")
print("OK")
