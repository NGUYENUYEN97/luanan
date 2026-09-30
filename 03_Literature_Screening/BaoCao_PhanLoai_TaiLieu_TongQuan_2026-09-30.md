# Báo cáo rà soát, phân loại tài liệu cho chuyên đề tổng quan (30/9/2026)

## 1. Phạm vi và cách làm

- Rà soát toàn bộ tệp .pdf, .docx, .txt, .md trong `01_NotebookLM_Inputs`, trừ `01_VanBanPhapQuy` (văn bản pháp luật) và `Niêm giám` (số liệu). Sau khi loại các tệp trùng nội dung (cùng mã băm), còn 569 tệp.
- Với mỗi tệp, đọc văn bản của 3 trang đầu; với các tài liệu đề xuất, đọc thêm phần tóm tắt, kết luận hoặc phần phương pháp, ghi lại số trang đã đọc.
- Nguyên tắc liêm chính áp dụng:
  - Tác giả, năm, tên, nơi xuất bản chỉ được ghi khi đọc được trên bản gốc. Tài liệu còn thiếu thông tin được tách riêng vào nhóm F, không đề xuất đưa vào bản thảo.
  - Số trang in nếu phải suy từ số trang của tệp đều được ghi rõ ở cột Ghi chú.
  - Không dùng làm nguồn: các bản "tổng quan" do công cụ AI (LeapSpace) tạo ra, bản dịch máy, luận văn thạc sĩ, tài liệu tập huấn và phổ biến kiến thức.
  - Cột "Nội dung dùng để phân tích trong luận án" chỉ là gợi ý của người rà soát. Khi viết, NCS cần đọc toàn văn phần sẽ trích.

## 2. Kết quả phân loại

| Nhóm | Ý nghĩa | Số tệp |
|---|---|---|
| A | Đã trích dẫn trong chuyên đề | 91 |
| B | Đề xuất bổ sung, ưu tiên (đã xác minh trên bản gốc) | 35 |
| B* | Bản trùng hoặc tệp đi kèm của tài liệu nhóm B (tóm tắt, trang thông tin) | 24 |
| C | Tham khảo thêm, đúng chủ đề nhưng ưu tiên thấp hơn | 109 |
| D | Dùng cho phần khác của luận án (kinh nghiệm quốc tế, số liệu, văn bản) | 125 |
| E | Không dùng làm nguồn học thuật chính | 152 |
| F | Có giá trị nhưng thiếu thông tin thư mục, cần xác minh | 8 |
| G | Tệp phụ trợ, tệp rỗng, tệp không đọc được, tệp hành chính | 25 |

Chi tiết từng tệp nằm ở sheet `4_ToanBo_TaiLieu` của `PhanLoai_TaiLieu_TongQuan_2026-09-30.xlsx`.

## 3. Đề xuất bổ sung theo từng mục (nhóm B)

Bảng đầy đủ 14 cột theo mẫu `Literature review table MẪU.xlsx` nằm tại `Literature_Review_BoSung_TongQuan.md` và sheet `2_LiteratureReview`. BibTeX để nhập Zotero: `references_bo_sung_tong_quan.bib`.

**1.1.1.1 (hiện có 21 nguồn, bổ sung có chọn lọc):**
- WIPO (2024), *Models of Intellectual Property Governance and Administration*, dùng cho hướng 3. Đây là công trình quốc tế hiếm hoi phân tích trực tiếp mô hình tổ chức cơ quan quản lý SHTT.
- Fink, Maskus và Qian (2016), dùng cho hướng 2 (tác động của vi phạm quyền và nhu cầu dữ liệu cho chính sách thực thi).
- Drahos (1996), *A Philosophy of Intellectual Property*, dùng cho hướng 1 như một luồng quan điểm phản biện.

**1.1.1.2 (hiện có 10 nguồn, hướng 1 chỉ có 1 nguồn):**
- Phạm Minh Huyền (2024), luận án về giới hạn quyền tác giả, quyền liên quan, làm dày hướng 1.
- Trần Văn Hải (2021), luận án về chính sách pháp luật hình sự đối với các tội xâm phạm quyền SHTT, dùng cho hướng 2.
- Hứa Thị Hồng (2023), luận án về bảo vệ quyền SHTT tại biên giới, dùng cho hướng 2.
- Nguyễn Hồ Bích Hằng (2025), Tạp chí Luật học, dùng cho hướng 2. Cần đối chiếu tên bài tiếng Việt.

**1.1.2.1 (hiện có 11 nguồn):**
- Hướng 1 (đo lường): OECD (2014), Barefoot và cộng sự (2018, BEA), UNCTAD (2017). Ba nguồn này lấp khoảng giữa Bukht và Heeks (2017) và OECD (2020).
- Hướng 2: WIPO và Luiss (2024), số liệu mới về đầu tư tài sản vô hình.
- Hướng 3: Dahlman, Mealy và Wermelinger (2016, OECD), về vai trò hoạch định của Nhà nước ở nước đang phát triển.
- Đề xuất thêm hướng 4 "quan hệ giữa bảo hộ quyền SHTT và phát triển kinh tế số", đối xứng với hướng 3 của mục 1.1.2.2. Nguồn gồm Chen và Wu (2022), Zheng, Li và Zhuang (2023), Yuan và Li (2025), Ding và Yang (2025). Nếu muốn giữ mục 1.1.2 thuần về kinh tế số, có thể chuyển nhóm này sang hướng 4 của mục 1.1.3.1.

**1.1.2.2 (hiện có 11 nguồn):**
- Trần Thị Lan (2025), luận án kinh tế về các yếu tố ảnh hưởng đến phát triển kinh tế số. Thể chế số là một yếu tố trong mô hình nhưng quyền SHTT không có, qua đó củng cố khoảng trống của mục.
- Nguyễn Thị Hải Hà (2023), về tiêu chí đánh giá hiệu quả QLNN trong nền kinh tế số. Đây là tài liệu gần nhất với phần tiêu chí đánh giá của luận án.
- Nguyễn Phương Thảo (2024), về đo lường; Ngô Hoài Sơn và Nguyễn Lê Kim Kiều (2022) và Trương Nam Trung (2020), về thể chế và phương pháp QLNN; Nguyễn Thị Tuyết Nga (2025), về chính sách SHTT trong kinh tế số.

**1.1.3 (đoạn mở đầu làm rõ thuật ngữ khai thác):**
- Kamiyama, Sheehan và Martinez (2006, OECD). Tài liệu dùng trực tiếp thuật ngữ *exploitation* và liệt kê các kênh: sử dụng nội bộ, li-xăng, công cụ đàm phán, huy động vốn.
- Liu Haibo và Liu Liang (2016). Tài liệu chia hình thức thương mại hóa thành thực hiện nội bộ, lưu thông ra bên ngoài, tài chính SHTT và tố tụng.

**1.1.3.1:**
- Hướng 2: Goldfarb và Henrekson (2003) về chính sách từ trên xuống và từ dưới lên, là bằng chứng trực tiếp cho tranh luận về mức độ can thiệp của Nhà nước ở mục 1.2.3. Shane (2004) đặt cạnh Mowery và cộng sự (2001) để thể hiện hai luồng đánh giá về Bayh-Dole. Ngoài ra có Siegel, Veugelers và Wright (2007), Thursby và Kemp (2002), Holgersson và Aaboen (2019), UNECE (2011).
- Hướng 3: APEC (2023), báo cáo về hệ thống tài chính SHTT ở châu Á - Thái Bình Dương.
- Hướng 4: Lee (2026) về điều phối chính sách trong hệ sinh thái nội dung số; Fu và cộng sự (2026) về định giá bằng AI; UNCTAD (2024) về vai trò Nhà nước trong công nghiệp nội dung Hàn Quốc.

**1.1.3.2:**
- Phạm Thị Thúy Hằng (2021), luận án quản lý giáo dục về quản lý hoạt động SHTT ở 3 đại học miền Trung. Tài liệu khảo sát 712 người, gồm cả cán bộ quản lý và giảng viên, chuyên viên, theo cùng một bộ tiêu chí. Có thể bổ sung vào hướng 1 và Bảng 1.4.

## 4. Tài liệu cần xác minh trước khi dùng (nhóm F)

Danh sách 8 tài liệu, kèm thông tin còn thiếu, nằm ở sheet `3_CanXacMinh`. Trong đó đáng chú ý:
- Bài của Phan Quốc Nguyên và Trịnh Linh Chi (2025) trên Tạp chí Tòa án nhân dân về bồi thường thiệt hại do xâm phạm quyền SHTT.
- Bài trên Tạp chí Phát triển Khoa học và Công nghệ (2017) về các yếu tố ảnh hưởng đến vi phạm bản quyền số, có khảo sát 264 người và sử dụng SEM.
- Cuốn *The Exploitation of Intellectual Property Rights* do Schovsbo chủ biên. Thư mục chỉ có trang giới thiệu, nên cần bản đầy đủ nếu muốn dùng cho đoạn thuật ngữ.

## 5. Những chỗ thư mục chưa có tài liệu phù hợp

Các chỗ dưới đây nên tìm thêm từ cơ sở dữ liệu. Báo cáo này không tự đề xuất tên công trình khi chưa có bản gốc.
- Mục 1.1.1.2, hướng 1: giáo trình và sách chuyên khảo lý luận về quyền SHTT của các cơ sở đào tạo luật trong nước, xuất bản sau năm 2006. Từ khóa gợi ý: "giáo trình luật sở hữu trí tuệ", "lý luận quyền sở hữu trí tuệ".
- Mục 1.1.2.2, hướng 3: công trình trong nước về khai thác (không chỉ bảo hộ) tài sản trí tuệ trên nền tảng số. Từ khóa gợi ý: "thương mại hóa tài sản trí tuệ nền tảng số", "sàn giao dịch công nghệ trực tuyến".
- Yếu tố phối hợp liên ngành và nguồn lực của cơ quan quản lý: ngoài WIPO (2024), thư mục chưa có công trình phân tích trực tiếp. Từ khóa gợi ý: "IP office coordination", "inter-agency coordination intellectual property enforcement".

## 6. Thay đổi trong chuyên đề tổng quan (đã commit)

- Đồng bộ các tệp nguồn `.md` với bản `.docx` NCS chỉnh sửa ngày 30/9. Nội dung đồng bộ gồm đoạn "Thứ năm", Bảng 1.2, các đoạn bổ sung về mẫu và phương pháp ở mục 1.1.3.2.
- Đánh số các yếu tố (1)-(9) trong Bảng 1.2.
- Thêm Bảng 1.3: ma trận đối chiếu 33 công trình đã tổng quan với 9 yếu tố, kèm đoạn nhận xét.
- Thêm Bảng 1.4: chủ thể khảo sát, mẫu và phương pháp của 9 công trình thực nghiệm trong nước. Số liệu mẫu đã được đối chiếu với bản gốc.
- Điểm cần NCS thống nhất: mục 1.2.3 ("Thứ ba") nêu 3 nhóm chủ thể khảo sát (doanh nghiệp, cơ quan quản lý, tổ chức trung gian), trong khi đề cương chi tiết (`10_DeCuong.md`) nêu 4 nhóm (thêm nhà khoa học).

## 7. Đã đưa vào chuyên đề (cập nhật cùng ngày)

17 tài liệu nhóm B đã được viết vào các mục 1.1.1.2, 1.1.2.1, 1.1.2.2: Phạm Minh Huyền (2024), Trần Văn Hải (2021), Hứa Thị Hồng (2023), OECD (2014), Barefoot và cộng sự (2018), WIPO và Luiss (2024), UNCTAD (2017), Dahlman và cộng sự (2016), Chen và Wu (2022), Zheng và cộng sự (2023), Yuan và Li (2025), Ding và Yang (2025), Trần Thị Lan (2025), Nguyễn Phương Thảo (2024), Truong Nam Trung (2020), Ngô Hoài Sơn và Nguyễn Lê Kim Kiều (2022), Nguyễn Thị Hải Hà (2023), Nguyễn Thị Tuyết Nga (2025).

Chưa đưa vào: nhóm tài liệu cho mục 1.1.1.1, đoạn thuật ngữ đầu mục 1.1.3, mục 1.1.3.1, 1.1.3.2 (18 tài liệu) và Nguyễn Hồ Bích Hằng (2025) vì chưa đối chiếu được tên bài tiếng Việt.
