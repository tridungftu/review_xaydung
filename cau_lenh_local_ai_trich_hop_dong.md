# Câu lệnh cho Local AI: đọc hợp đồng, trích thông tin ra bảng

Dùng với model chạy trên RunPod, LM Studio hoặc Ollama. Đặt temperature = 0.

Có hai cách dùng:
- **Cách A (chat)**: dán cả System + User vào giao diện chat, lấy bảng TSV dán vào Excel.
- **Cách B (gọi API)**: gửi System và User thành 2 message riêng, yêu cầu JSON để script ghi vào master.

---

## Bước 1. Đọc hợp đồng và trích tham số

### System prompt (cố định, dùng cho mọi hợp đồng)

```
Bạn là chuyên viên kiểm soát hợp đồng xây dựng tại Việt Nam. Bạn đọc hợp đồng và trích tham số chính xác tuyệt đối để phục vụ kiểm tra hồ sơ thanh toán.

QUY TẮC BẮT BUỘC
1. Chỉ dùng thông tin có trong văn bản hợp đồng được cung cấp. Không suy đoán, không dùng kiến thức bên ngoài, không dùng quy định pháp luật để điền thay.
2. Mỗi tham số phải có "trich_dan": chép NGUYÊN VĂN câu trong hợp đồng làm căn cứ, không sửa chữ, không tóm tắt.
3. Không tìm thấy: gia_tri = "KHÔNG TÌM THẤY", trich_dan = "".
4. Hợp đồng ghi rõ "Không áp dụng": gia_tri = "Không áp dụng", vẫn ghi trich_dan.
5. Giá trị phải suy ra (không ghi trực tiếp): ghi vào ghi_chu "Suy ra: <cách suy ra>".
6. Định dạng: tỷ lệ ghi 10%; số tiền ghi số nguyên không dấu phân cách (1234567890); ngày ghi dd/mm/yyyy; thời hạn ghi số kèm đơn vị ở cột don_vi.
7. dieu_khoan: số Điều và tên Điều chứa câu trích dẫn, ví dụ "Điều 5. Loại Hợp đồng, Giá trị Hợp đồng, Tạm ứng và thanh toán".
8. Ghi vào ghi_chu mọi điểm bất thường: ghi sai tên/ký hiệu bên, cùng một nội dung ghi hai con số khác nhau, tỷ lệ mâu thuẫn, điều khoản tham chiếu đến điều không tồn tại.
9. Bỏ qua phần mục lục. Nếu cùng thông tin xuất hiện ở trang bìa và phần thân, lấy phần thân.
```

### User prompt (thay `{HOP_DONG}` bằng nội dung hợp đồng)

```
Đọc hợp đồng dưới đây và trích đủ các tham số trong DANH SÁCH, giữ đúng mã và thứ tự.

DANH SÁCH THAM SỐ (mã | nhóm | tham số)
TC-01 | Thông tin chung | Số hợp đồng
TC-02 | Thông tin chung | Ngày ký
TC-03 | Thông tin chung | Gói thầu
TC-04 | Thông tin chung | Dự án
TC-05 | Thông tin chung | Địa điểm
TC-06 | Thông tin chung | Bên A (bên giao thầu / chủ đầu tư)
TC-07 | Thông tin chung | Bên B (bên nhận thầu / nhà thầu)
TC-08 | Thông tin chung | Hiệu lực hợp đồng
TT-01 | Giá hợp đồng | Loại hợp đồng (trọn gói / đơn giá cố định / đơn giá điều chỉnh / theo thời gian / kết hợp)
TT-02 | Giá hợp đồng | Giá trị hợp đồng (đồng)
TT-03 | Giá hợp đồng | Giá trị hợp đồng đã gồm VAT hay chưa (Có / Không)
TT-04 | Giá hợp đồng | Thuế suất GTGT (%)
TT-05 | Giá hợp đồng | Cơ sở thanh toán khối lượng (KL nghiệm thu thực tế / KL hợp đồng / % hoàn thành)
TT-10 | Tạm ứng | Tỷ lệ tạm ứng (%), tính trên giá trị trước hay sau VAT
TT-11 | Tạm ứng | Cách thu hồi tạm ứng (bắt đầu khi nào, mỗi đợt thu bao nhiêu, thu hết khi nào)
TT-12 | Tạm ứng | Bảo lãnh tạm ứng
TT-20 | Thanh toán đợt | Tỷ lệ thanh toán mỗi đợt (% giá trị hoàn thành / nghiệm thu)
TT-21 | Thanh toán đợt | Tỷ lệ giữ lại mỗi đợt (%); không ghi trực tiếp thì = 100% - TT-20, ghi "Suy ra"
TT-22 | Thanh toán đợt | Thời hạn thanh toán
TT-30 | Quyết toán - bảo hành | Tỷ lệ thanh toán khi quyết toán (%)
TT-31 | Quyết toán - bảo hành | Tỷ lệ bảo hành hoặc bảo lãnh bảo hành (%)
TT-32 | Quyết toán - bảo hành | Thời hạn bảo hành
TT-33 | Quyết toán - bảo hành | Bảo lãnh thực hiện hợp đồng
TT-40 | Điều chỉnh - phát sinh | Điều kiện điều chỉnh đơn giá / giá hợp đồng
TT-41 | Điều chỉnh - phát sinh | Ngưỡng biến động để điều chỉnh giá (%), giá gốc làm căn cứ nếu có
TT-42 | Điều chỉnh - phát sinh | Cách xử lý khối lượng phát sinh, vượt hợp đồng
TT-50 | Tiến độ - phạt | Tiến độ / thời gian thực hiện hợp đồng
TT-51 | Tiến độ - phạt | Mức phạt vi phạm chất lượng (% tối đa, tính trên giá trị nào)
TT-52 | Tiến độ - phạt | Tạm giữ hoặc phạt chậm tiến độ giai đoạn - mức 1 (%/ngày, khi nào)
TT-53 | Tiến độ - phạt | Tạm giữ hoặc phạt chậm tiến độ giai đoạn - mức 2 (%/ngày, khi nào)
TT-54 | Tiến độ - phạt | Phạt chậm tổng tiến độ (%/ngày, tính trên giá trị nào)
TT-55 | Tiến độ - phạt | Trần tổng mức phạt (%)
TT-56 | Tiến độ - phạt | Cách khấu trừ tiền phạt vào thanh toán
TT-60 | Hồ sơ theo HĐ | Hồ sơ thanh toán mỗi đợt (liệt kê, cách nhau bằng dấu ;)
TT-61 | Hồ sơ theo HĐ | Hồ sơ quyết toán (liệt kê, cách nhau bằng dấu ;)
TT-62 | Hồ sơ theo HĐ | Số bộ hồ sơ thanh toán phải nộp

ĐỊNH DẠNG ĐẦU RA
{DINH_DANG}

HỢP ĐỒNG
{HOP_DONG}
```

### Thay `{DINH_DANG}` bằng một trong hai đoạn sau

**Cách A: bảng dán Excel**

```
Chỉ in bảng, không viết gì thêm. Các cột cách nhau bằng ký tự TAB, mỗi tham số một dòng, dòng đầu là tiêu đề:
Mã	Nhóm	Tham số	Giá trị	Đơn vị	Điều khoản	Trích dẫn nguyên văn	Ghi chú
```

**Cách B: JSON cho script**

```
Chỉ in một đối tượng JSON hợp lệ, không có ``` và không viết gì thêm, theo dạng:
{"TC-01": {"gia_tri": "...", "don_vi": "", "dieu_khoan": "...", "trich_dan": "...", "ghi_chu": ""}, "TC-02": {...}, ...}
Đủ tất cả các mã trong danh sách.
```

---

## Lưu ý khi chạy

1. **Độ dài**: một hợp đồng khoảng 25–35 nghìn token. Nên dùng model có ngữ cảnh từ 32k trở lên và đặt max context tương ứng. Nếu ngữ cảnh ngắn hơn, chỉ gửi trang đầu (căn cứ, các bên) cùng các Điều về giá, thanh toán, bảo lãnh, bảo hành, điều chỉnh giá, tiến độ, phạt. Với khung hợp đồng mẫu hiện tại, đó là Điều 5, 6, 7, 8, 16.
2. **Chuẩn bị văn bản**:
   - Word: Ctrl+A, Ctrl+C.
   - PDF có lớp chữ: copy trực tiếp.
   - PDF scan: OCR trước.
   - Nên bỏ mục lục và phần chữ ký.
3. **Kiểm tra kết quả**:
   - Với mỗi trích dẫn, dùng Ctrl+F tìm lại trong file Word. Không tìm thấy nghĩa là model đã viết lại, cần sửa tay.
   - So với master do `master_hd.py` tạo; chỗ khác nhau thì đọc lại hợp đồng.
4. **Bảo mật khi chạy trên RunPod**: văn bản hợp đồng được gửi lên máy chủ thuê. Dùng pod riêng, có xác thực, tắt pod sau khi dùng.

---

## Bước 2. Bảng khối lượng và đơn giá (BOQ) nằm trong PDF / Word

```
Đọc bảng khối lượng và đơn giá bên dưới, chép lại thành bảng để dán vào Excel.

QUY TẮC
1. Chép đúng từng dòng, đúng thứ tự, không gộp, không bỏ dòng, không tự tính lại.
2. Dòng tiêu đề nhóm (không có đơn vị tính) vẫn giữ, để trống các cột số.
3. Số ghi dạng số thuần: dấu chấm thập phân, không dấu phân cách hàng nghìn (1405542; 55.506).
4. Dòng "Cộng", "Thuế", "Tổng cộng" chép ở cuối, ghi rõ ở cột Tên công việc.
5. Ô nào đọc không rõ thì ghi ?? và ghi lý do ở cột Ghi chú.

ĐỊNH DẠNG ĐẦU RA
Chỉ in bảng, các cột cách nhau bằng TAB, dòng đầu là tiêu đề:
STT	Nhóm / hạng mục	Tên công việc	ĐVT	Khối lượng	Đơn giá	Thành tiền	Ghi chú

BẢNG
{BANG}
```

---

## Chạy tự động với máy AI trên RunPod (không cần chép dán)

`master_hd.py ai-trich` gửi đúng câu lệnh trên (bản JSON) tới model, rồi so kết quả AI với kết quả trích theo quy tắc.

1. Bật pod, chạy `1_cai_dat.bat`, rồi mở `2_mo_giao_dien.bat` và **giữ cửa sổ này mở** (tunnel `localhost:11434` → Ollama trên pod).
2. Kéo thả file hợp đồng vào `5_trich_hop_dong_AI.bat`, hoặc chạy lệnh:
   `python master_hd.py ai-trich "HĐ 0204.docx" --an-danh --model qwen3.6:35b`
3. Kết quả nằm trong thư mục `ket_qua_AI`:
   - `so_sanh_AI_<tên HĐ>.xlsx`
   - file `.json` chứa câu trả lời gốc của AI

Script tự làm các việc sau:
- **Chỉ gửi các Điều liên quan**: phần mở đầu và các Điều về giá, thanh toán, tạm ứng, bảo lãnh, bảo hành, điều chỉnh, tiến độ, phạt, nghiệm thu, hiệu lực. Với hợp đồng mẫu, 68 nghìn ký tự còn khoảng 19,5 nghìn ký tự, vừa ngữ cảnh 32k. Muốn gửi cả hợp đồng thì thêm `--ca-hop-dong`.
- **`--an-danh`**: che tên hai bên, tên người, mã số thuế, số tài khoản, số điện thoại trước khi gửi lên pod, theo nguyên tắc chỉ dùng tài liệu đã ẩn danh trong giai đoạn thử.
- **Tắt chế độ suy luận** (`think: false`) để qwen3.6 không viết phần suy luận dài bằng tiếng Anh.
- **So từng tham số**, xếp vào một trong các loại: Khớp / Lệch / Chỉ AI thấy / Chỉ quy tắc thấy / Khác cách ghi.
- **Kiểm tra từng trích dẫn của AI có thật trong hợp đồng không**. Dòng "KHÔNG - AI viết lại hoặc bịa" là chỗ phải đọc lại.
- **Ghi số token vào, token ra và thời gian chạy** để đưa vào đánh giá POC.
