# Ghi chú hoàn thiện chuyên đề tổng quan và đề cương (30/9/2026)

## Sản phẩm cuối (tách thành 2 file theo yêu cầu NCS)

- `ChuyenDe_TongQuan_FINAL.docx` (+ .pdf): Chuyên đề tổng quan gồm bìa, mục lục, Mở đầu, tổng quan 1.1-1.3, Kết luận, 129 tài liệu tham khảo. Phần chính 30 trang.
- `DeCuong_ChiTiet_FINAL.docx` (+ .pdf): Đề cương chi tiết gồm Mở đầu viết theo `DeCuong_FINAL_2026-09-30.md` (bản NCS thích), kết cấu chi tiết 4 chương theo mẫu trường, 20 tài liệu viện dẫn. Dài 19 trang.
- Đoạn "Tính cấp thiết" dùng chung cho cả hai file: `ChuyenDe_TongQuan_FINAL/00_TinhCapThiet.md`.
- Không còn dùng: `ChuyenDe_TongQuan_va_DeCuong_FINAL.docx/.pdf` (bản gộp), `Đề cương NCS Uyen_FINAL_2026-09-30.docx`, `ChuyenDe_TongQuan_FINAL/06_Phan2.md`.
- Nguồn dựng: thư mục `ChuyenDe_TongQuan_FINAL/` (các file .md, `refs.py`, `build_chuyen_de.py`, `cap_nhat_muc_luc.ps1`).

Sửa nội dung: sửa file .md tương ứng, sau đó chạy:
```
python ChuyenDe_TongQuan_FINAL\build_chuyen_de.py chuyende
python ChuyenDe_TongQuan_FINAL\build_chuyen_de.py decuong
powershell -File ChuyenDe_TongQuan_FINAL\cap_nhat_muc_luc.ps1 -Docx <đường dẫn file .docx>
```

So với file `DeCuong_FINAL_2026-09-30.md`, Mở đầu của đề cương chỉ thay đổi các điểm sau:
- Cỡ mẫu 150-180 phiếu (4 nhóm); phỏng vấn 15-20 người.
- Kinh nghiệm quốc tế thêm Hoa Kỳ.
- Thêm số liệu năm 2024, tỷ trọng kinh tế số và WIPO 2026 vào tính cấp thiết.
- Thêm Bảng 1 (tiêu chí và chỉ báo).
- Nêu rõ là không dùng EFA và hồi quy.
- Kết cấu chi tiết theo khung mẫu trường; bỏ mục "Dự kiến sản phẩm khoa học" vì không có trong mẫu.

## Các lựa chọn đã chốt với NCS

- Tên đề tài giữ nguyên: Quản lý nhà nước về quyền sở hữu trí tuệ trong nền kinh tế số ở Việt Nam; phạm vi giới hạn ở QLNN đối với khai thác quyền SHTT. Hai dự thảo 20/8 và 15/9 đề xuất đổi tên thành "QLNN về khai thác quyền SHTT...", nhưng không áp dụng vì NCS đã chốt giữ tên.
- Phương pháp: hỗn hợp, định lượng mô tả kết hợp định tính giải thích; 150-180 phiếu (4 nhóm chủ thể), 15-20 phỏng vấn sâu; không dùng EFA và hồi quy.
- Mốc giải pháp: đến năm 2030, tầm nhìn đến năm 2045.
- Kết cấu: bám mẫu của cơ sở đào tạo.

## Những gì đã chắt lọc từ các dự thảo

| Nguồn | Nội dung được kế thừa |
|---|---|
| Dự thảo 20/8 (`de_cuong_luan_an_khai_thac_QSHTT.docx`) | 4 nội dung QLNN; 5 tiêu chí và chỉ báo (Bảng 2.1); kinh nghiệm Hoa Kỳ, Hàn Quốc, Trung Quốc, Singapore; nhóm giải pháp tương ứng nội dung |
| Dự thảo 15/9 (`15.9 ĐỀ CƯƠNG LUẬN ÁN TIẾN SĨ.docx`) | Bỏ hồi quy; 3 nhóm khảo sát, 150-180 phiếu; tách khai thác thành 3 nhóm hình thức; coi thay đổi thể chế 2026 là chuẩn tham chiếu, không đánh giá hiệu quả; khoảng trống dữ liệu là vấn đề quản lý; nội dung "theo dõi, giám sát, đánh giá" thay vì "chống xâm phạm" |
| Bản 29/9 (3 file .md) | Ba nhóm tổng quan (SHTT, kinh tế số, khai thác quyền SHTT trong kinh tế số) |
| Các mục tổng quan đã viết (1.1.1.1, 1.1.1.2, 1.1.2) | Giữ cấu trúc và phần lớn nhận định; sửa trích dẫn sai, bỏ nguồn không kiểm chứng được, chuyển nội dung thị trường công nghệ sang 1.1.3 |
| Bảng 26 công trình trong nước | Dùng danh sách, nhưng **đọc lại bản gốc** vì bảng có nhiều sai sót về tác giả, tên bài |

## Các sai sót trích dẫn đã sửa

- Neves và cộng sự (2021): đúng là "The link between intellectual property rights, innovation, and growth: A meta-analysis", *Economic Modelling*, 97, 196-209 (bản cũ ghi sai tên bài, tạp chí).
- Báo cáo 2021 về quyền SHTT và kết quả hoạt động của doanh nghiệp: của EPO và EUIPO (bản cũ ghi EUIPO và OECD).
- Bỏ Goold (2024) vì không xác minh được công trình.
- Bảng 26 bài: T03 thực tế là Hoàng Lan Phương (2019); T02 là Nguyễn Minh Huyền Trang (một tác giả); T07 "Những nút thắt..." là Nguyễn Hữu Xuyên (2018), không phải Nguyễn Hữu Cẩn; file CVv286 là bài của Nguyễn Quốc Thịnh và cộng sự (2020).
- Trần Văn Hải: luận án của tác giả này về chính sách pháp luật hình sự, không phải về thương mại hóa; bản chuyên đề dùng bài Tran Van Hai (2022) trên *VNU Journal of Science* về khai thác thông tin phục vụ thương mại hóa quyền SHTT.

## Số liệu mới đưa vào (đều có nguồn)

- Kinh tế số ước đạt 14,02% GDP năm 2025, khoảng 72,1 tỷ USD (Cục Thống kê, thông cáo 05/01/2026).
- Năm 2024: hơn 53.600 văn bằng bảo hộ được cấp; 153 đơn đăng ký hợp đồng chuyển quyền sử dụng đối tượng sở hữu công nghiệp; 1.758 đơn đăng ký hợp đồng chuyển nhượng, trong đó 1.643 về nhãn hiệu, 98 về sáng chế (Báo cáo thường niên hoạt động SHTT 2024).
- WIPO WIPR 2026: thời gian từ khi công nghệ ra đời đến khi được ứng dụng giảm mạnh.

## Việc NCS cần làm

1. Điền thông tin trên trang bìa: cơ sở đào tạo, họ tên NCS, người hướng dẫn.
2. Rà lại hai nguồn chưa đối chiếu được bản gốc: Lê Nết (2006) và Đoàn Đức Lương, Ngô Minh Tiến (2022) (tên tạp chí lấy theo bảng cũ).
3. Nếu thầy/cô yêu cầu cỡ mẫu khác (ví dụ 100 phiếu), sửa ở `06_Phan2.md` (mục 6 và Bảng 2.2) và trong `hinh_de_cuong/ve_hinh_quy_trinh.py`.

## Bổ sung ngày 30/9/2026 (sau rà soát thư mục NotebookLM)

- Mục 1.1.1.2: thêm Phạm Minh Huyền (2024) vào hướng 1; Trần Văn Hải (2021), Hứa Thị Hồng (2023) vào hướng 2 (năng lực và phối hợp của bộ máy thực thi).
- Mục 1.1.2.1: thêm OECD (2014), Barefoot và cộng sự (2018) vào hướng 1; WIPO và Luiss (2024) vào hướng 2; UNCTAD (2017), Dahlman và cộng sự (2016) vào hướng 3; thêm hướng 4 về quan hệ giữa bảo hộ quyền SHTT và kinh tế số (Chen và Wu, 2022; Zheng và cộng sự, 2023; Yuan và Li, 2025; Ding và Yang, 2025).
- Mục 1.1.2.2: thêm Trần Thị Lan (2025), Nguyễn Phương Thảo (2024) vào hướng 1; Truong Nam Trung (2020), Ngô Hoài Sơn và Nguyễn Lê Kim Kiều (2022), Nguyễn Thị Hải Hà (2023) vào hướng 2; Nguyễn Thị Tuyết Nga (2025) vào hướng 3.
- Mục 1.2: thêm Bảng 1.3 (ma trận công trình và yếu tố), Bảng 1.4 (phương pháp các công trình thực nghiệm trong nước); cập nhật Bảng 1.1.
- Danh mục tài liệu tham khảo: từ 95 lên 129 tài liệu (đợt một 113, đợt hai 129). Tệp `ChuyenDe_TongQuan_FINAL.pdf` chưa được xuất lại, cần xuất từ bản .docx sau khi cập nhật mục lục.
- Căn cứ chọn tài liệu: `03_Literature_Screening/BaoCao_PhanLoai_TaiLieu_TongQuan_2026-09-30.md`.
- Đợt hai: bổ sung mục 1.1.1.1 (hướng 1, 2, 3), đoạn làm rõ thuật ngữ khai thác đầu mục 1.1.3, mục 1.1.3.1 (hướng 2, 3, 4), mục 1.1.3.2 (Phạm Thị Thúy Hằng, 2021); cập nhật Bảng 1.3, Bảng 1.4 và luận điểm "Thứ nhất" ở mục 1.2.3.
- Thống nhất đối tượng khảo sát gồm 4 nhóm (doanh nghiệp và chủ sở hữu quyền, cơ quan QLNN, tổ chức trung gian, nhà khoa học) ở mục 1.2.3 cho khớp với đề cương.
- Zotero: chạy `python ChuyenDe_TongQuan_FINAL/xuat_zotero.py` để xuất `Zotero_ChuyenDe_TongQuan.json` (CSL JSON, khuyến nghị) và `Zotero_ChuyenDe_TongQuan.ris`; nhập vào Zotero bằng File > Import. Mỗi mục có ghi chú khóa refs.py và số thứ tự [n] trong danh mục.
