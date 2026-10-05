# Quy trình review hồ sơ thanh toán khối lượng xây dựng

| | |
|---|---|
| Phiên bản | 2.3, ngày 04/10/2026 (Module 2 – Lớp 1 tính toàn vẹn file: `toan_ven.py`, mục 11; vai trò Local AI, mục 5) |
| Áp dụng cho | Hợp đồng thi công xây dựng, cung cấp vật tư thiết bị, EPC; mọi mẫu hồ sơ thanh toán Excel |
| Thay đổi so với 1.x | Lấy **master file hợp đồng** làm gốc; bỏ bước tính lại bảng tổng hợp; các chi tiết riêng của từng hợp đồng để trong tài liệu nội bộ |

**Bộ công cụ đi kèm**

| File | Dùng để |
|---|---|
| `master_hd.py` | Trích hợp đồng + BOQ ra master file Excel (bước nhập ban đầu, để người review xác nhận) |
| `csdl_hop_dong.py` | **Cơ sở dữ liệu** (file `co_so_du_lieu/hop_dong.db`): nạp master, nạp phụ lục, BOQ hiệu lực, nạp đợt đã duyệt, kiểm tra đợt mới, xuất Excel |
| `toan_ven.py` | **Module 2 – Lớp 1: tính toàn vẹn file** hồ sơ thanh toán, mọi mẫu Excel (mục 11) |
| `0_…`, `6_…` – `9_…`, `13_….bat` | Chạy một nút: nạp master, phụ lục, kiểm tra đợt, xuất CSDL; 13: kiểm tra tính toàn vẹn file |
| `hardcode_audit.py` | Phân rã công thức và quét số nhập tay trong file hồ sơ (đào sâu khi cần, mục 4 bước B3) |
| `app_kiem_tra_ho_so/` | App chạy trên máy: kiểm tra đủ hồ sơ, hỏi đáp, gọi các công cụ trên |

---

## Mục lục

1. [Nguyên tắc](#1-nguyên-tắc)
2. [Bảy nhóm dữ liệu gốc và nguồn của chúng](#2-bảy-nhóm-dữ-liệu-gốc-và-nguồn-của-chúng)
3. [Cơ sở dữ liệu hợp đồng – phụ lục – BOQ](#3-cơ-sở-dữ-liệu-hợp-đồng--phụ-lục--boq)
4. [Quy trình review mỗi đợt](#4-quy-trình-review-mỗi-đợt)
5. [Phân vai: Python, local AI, người review](#5-phân-vai-python-local-ai-người-review)
6. [Chọn mẫu và trọng yếu](#6-chọn-mẫu-và-trọng-yếu)
7. [Danh mục tài liệu](#7-danh-mục-tài-liệu)
8. [Cách chạy công cụ](#8-cách-chạy-công-cụ)
9. [Thiết kế chương trình đối chiếu hợp đồng – phụ lục – hồ sơ](#9-thiết-kế-chương-trình-đối-chiếu-hợp-đồng--phụ-lục--hồ-sơ)
10. [Giới hạn hiện tại](#10-giới-hạn-hiện-tại)
11. [Khung phần mềm kiểm tra đối chiếu](#11-khung-phần-mềm-kiểm-tra-đối-chiếu-review_ttkl)

---

## 1. Nguyên tắc

1. **Số đề nghị thanh toán được ghép từ một số ít dữ liệu gốc.** Excel đã tính phép cộng trừ, nên không tính lại. Việc cần làm là biết số đó ghép từ dữ liệu nào, rồi chứng minh từng dữ liệu gốc.
2. **Mỗi dữ liệu gốc có một nguồn cố định** (mục 2). Nguồn cho hợp đồng, BOQ, phụ lục, tham số thanh toán, số liệu các đợt trước được gom vào **một cơ sở dữ liệu chung cho mọi hợp đồng** (mục 3). Hồ sơ đợt nào cũng đối chiếu vào cơ sở dữ liệu, không đối chiếu vào chính nó. Excel chỉ là phiếu nhập và bản xuất để theo dõi.
3. **Ba lớp kiểm tra** cho mỗi con số:

| Lớp | Câu hỏi | Công cụ |
|---|---|---|
| Khớp hợp đồng + phụ lục | GT HĐ, KL HĐ, đơn giá, KL kỳ trước trong hồ sơ có đúng như CSDL (hiệu lực tại ngày nghiệm thu) không? | `csdl_hop_dong.py kiem-dot` |
| Toàn vẹn file | Số nào trong file là nhập tay, số nào là công thức, liên kết có trỏ đúng không? | `hardcode_audit.py` (khi cần) |
| Chứng từ | KL kỳ này có biên bản nghiệm thu, bản vẽ hỗ trợ không? | Người review, chọn mẫu |

4. **Chi tiết riêng của từng hợp đồng nằm trong master, không nằm trong quy trình.** Ví dụ tỷ lệ giữ lại khác nhau theo phần xây dựng và thiết bị, có tạm ứng hay không, phụ lục thay thiết bị: đó là tham số của hợp đồng đó.

---

## 2. Bảy nhóm dữ liệu gốc và nguồn của chúng

| # | Dữ liệu gốc | Lấy từ | Lưu trong CSDL (bảng) | Trường hợp đợt đầu |
|---|---|---|---|---|
| 1 | Khối lượng hợp đồng | Hợp đồng + BOQ (Bảng chi tiết giá trị HĐ) + phụ lục | `boq_dong` (gốc + dòng phụ lục) | Như các đợt khác |
| 2 | Đơn giá | Hợp đồng + BOQ + phụ lục | `boq_dong` | Như các đợt khác |
| 3 | KL thanh toán kỳ trước | File thanh toán đợt trước đã duyệt (cột lũy kế đến hết kỳ này của đợt đó) | `kl_thanh_toan` | **Nhập tay** (thường bằng 0) |
| 4 | Lũy kế đợt trước (giá trị) | File thanh toán đợt trước đã duyệt | `dot_thanh_toan` | **Nhập tay** |
| 5 | KL thực hiện kỳ này | Biên bản nghiệm thu; khi cần phân tích bản vẽ | Không lưu; kiểm tra theo mẫu | |
| 6 | Tham số thanh toán (tạm ứng, thu hồi, tỷ lệ thanh toán, giữ lại, VAT, phạt, điều chỉnh giá) | Trích xuất từ hợp đồng (+ phụ lục sửa điều khoản) | `tham_so` (+ `thay_doi_tham_so`) | |
| 7 | Số đã thanh toán, đã tạm ứng, đã thu hồi | Đợt thanh toán trước (+ UNC) | `dot_thanh_toan` | Nhập tay (tạm ứng nếu có) |

Hệ quả: **mọi số nhập tay trong hồ sơ nhà thầu mà không thuộc 7 nhóm trên là một cờ cần giải thích**. Nếu thuộc 7 nhóm thì phải khớp master (nhóm 1, 2, 3, 4, 6, 7) hoặc biên bản nghiệm thu (nhóm 5).

---

## 3. Cơ sở dữ liệu hợp đồng – phụ lục – BOQ

**Nguyên tắc lưu trữ**

- **Một kho duy nhất** cho mọi hợp đồng: file SQLite `co_so_du_lieu/hop_dong.db`. SQLite có sẵn trong Python, không cần cài thêm. Excel chỉ là **phiếu nhập** (phiếu phụ lục) và **bản xuất** để theo dõi; sửa trên bản xuất không làm đổi kho.
- **Số liệu gốc không bao giờ bị sửa.** Phụ lục là bản ghi thay đổi riêng, nối về hợp đồng bằng mã HĐ. Trigger trong CSDL chặn sửa/xoá dữ liệu gốc. Sửa sai phải qua lệnh có `--ly-do`, và lệnh đó được ghi vào nhật ký.
- **"Hiệu lực tại ngày X" = gốc + các phụ lục ĐÃ KÝ có ngày hiệu lực ≤ X.** Khi kiểm tra một đợt thanh toán, **X là ngày nghiệm thu của đợt đó**. Các giá trị hiệu lực do chương trình tính, không ai nhập tay.
- **Mã HĐ nội bộ** (`HD0001`…) làm khoá; số HĐ theo văn bản là thuộc tính, vì số HĐ hay bị ghi lệch giữa các văn bản. Lệnh nào cũng nhận số HĐ hoặc một đoạn duy nhất của nó (ví dụ `0101`).

### 3.1. Các bảng

```
hop_dong (1 dòng / HĐ, số liệu gốc)
 ├── tham_so            mã TC-xx / TT-xx gốc, kèm trích dẫn (cùng bộ mã với MAP_THAM_SO của MAU_DANH_MUC_HOP_DONG)
 ├── phu_luc            1 dòng / phụ lục: số, ngày ký, ngày hiệu lực, có đổi BOQ, GT tăng/giảm, GT sau PL, trạng thái
 │     └── thay_doi_tham_so   tham số bị phụ lục sửa: giá trị cũ, giá trị mới, trích dẫn
 ├── boq_dong           1 dòng / hạng mục / nguồn: dòng GOC của HĐ + dòng THAY / THEM / BO của từng phụ lục
 ├── dot_thanh_toan     1 dòng / đợt: ngày nghiệm thu, GT kỳ, lũy kế, giữ lại, thu hồi TƯ, thanh toán
 │     └── kl_thanh_toan      KL kỳ, KL lũy kế từng Mã HM (nguồn KL kỳ trước cho đợt sau)
 ├── khop_hang_muc      dòng trong hồ sơ TT → Mã HM, nhớ cho các đợt sau
 ├── canh_bao           cờ ĐỎ / VÀNG phát sinh khi nạp BOQ, phụ lục
 └── nhat_ky            ai làm gì, lúc nào, từ file nào (không sửa/xoá được)
```

| Bảng | Khoá | Cột chính |
|---|---|---|
| `hop_dong` | Mã HĐ | Số HĐ, phân loại, loại giá (TT-01), ngày ký, gói thầu, GT HĐ gốc (TT-02), đã gồm VAT (TT-03), VAT (TT-04), file nguồn |
| `tham_so` | Mã HĐ + mã tham số | Đủ 36 tham số TC-01 … TT-62 của mỗi HĐ (và Y-xx bổ sung nếu có): giá trị gốc, đơn vị, điều khoản, trích dẫn nguyên văn, trạng thái xác nhận. Lưu dạng dòng để thêm tham số mới không phải sửa cấu trúc; khi xuất Excel thì trải thành 36 cột |
| `phu_luc` | id (Mã HĐ + Số PL là duy nhất trong các PL chưa huỷ) | Ngày ký, ngày hiệu lực, loại, đổi BOQ (Có/Không), GT tăng/giảm trước và sau VAT, GT HĐ sau PL, căn cứ, file PL, phiếu nhập, trạng thái (Đã ký / Dự thảo / Đã hủy), lý do huỷ |
| `thay_doi_tham_so` | PL + mã tham số | Giá trị cũ, giá trị mới, trích dẫn. TT-02 do máy tự ghi từ giá trị phụ lục |
| `boq_dong` | Mã HĐ + nguồn + Mã HM | Thao tác (GOC / THAY / THEM / BO), STT, nhóm, tên, ĐVT, KL, đơn giá chưa VAT, thành tiền ghi trên tài liệu, khoá so khớp, vị trí nguồn (file / sheet / dòng) |
| `dot_thanh_toan` | Mã HĐ + đợt | Ngày nghiệm thu, ngày nhận hồ sơ, các số tổng hợp, nguồn |
| `kl_thanh_toan` | Mã HĐ + đợt + Mã HM | KL kỳ, KL lũy kế, đơn giá thanh toán, vị trí nguồn |

### 3.2. Lưu BOQ của mỗi hợp đồng

- **Chung một bảng `boq_dong` cho mọi hợp đồng**, không tách mỗi hợp đồng một file.
- **Phụ lục chỉ ghi các dòng thay đổi.** Không chép lại cả BOQ sau mỗi phụ lục, nhờ vậy nhìn vào là biết phụ lục nào sửa dòng nào.
- **Mã HM cố định suốt đời hợp đồng.** THAY và BO trỏ vào mã đang có. THEM được cấp mã mới tiếp theo, ví dụ HM0098.
- **BOQ hiệu lực là kết quả tính**, gồm cột "Nguồn" (HĐ gốc / PL số) và "Lịch sử" (Gốc: KL × ĐG → PL01 thay: KL × ĐG …).
- **Hai cột thành tiền**: thành tiền ghi trên tài liệu và KL × ĐG. Hai cột này dùng để bắt lỗi số học ngay trong BOQ.

Ví dụ (số giả định):

| Nguồn | Ngày HL | Thao tác | Mã HM | Tên | ĐVT | KL | Đơn giá |
|---|---|---|---|---|---|---|---|
| HĐ gốc | 01/03 | GOC | HM0006 | Móng CPĐD lớp dưới | 100m3 | 25 | 100.000.000 |
| PL01 | 15/06 | THAY | HM0006 | | | 30 | |
| PL01 | 15/06 | THEM | HM0098 | Cọc tiêu bổ sung | cái | 20 | 500.000 |
| PL01 | 15/06 | BO | HM0007 | | | | |

Đợt nghiệm thu ngày 10/06 được đối chiếu với BOQ gốc. Đợt nghiệm thu ngày 10/07 được đối chiếu với HM0006 = 30, có HM0098, và không còn HM0007. Ở dòng THAY, cột nào để trống thì giữ giá trị trước đó.

### 3.3. Danh mục tham số trích từ hợp đồng

| Mã | Tham số | Bắt buộc trước khi đối chiếu |
|---|---|:---:|
| TC-01 … TC-08 | Số HĐ, ngày ký, gói thầu, dự án, địa điểm, bên A, bên B, hiệu lực | |
| TT-01 | Loại hợp đồng (trọn gói / đơn giá cố định / đơn giá điều chỉnh / …) | ✔ |
| TT-02 | Giá trị hợp đồng | ✔ |
| TT-03 | Giá trị HĐ đã gồm VAT hay chưa | ✔ |
| TT-04 | Thuế suất GTGT | ✔ |
| TT-05 | Cơ sở thanh toán khối lượng | |
| TT-10, TT-11, TT-12 | Tạm ứng: tỷ lệ, thu hồi, bảo lãnh | ✔ (TT-10) |
| TT-20, TT-21, TT-22 | Tỷ lệ thanh toán mỗi đợt, tỷ lệ giữ lại, thời hạn thanh toán | ✔ (TT-20 hoặc TT-21) |
| TT-30 … TT-33 | Quyết toán, bảo hành (tỷ lệ, thời hạn), bảo lãnh thực hiện | |
| TT-40 … TT-42 | Điều chỉnh đơn giá, ngưỡng điều chỉnh, xử lý phát sinh | |
| TT-50 … TT-55 | Tiến độ, phạt chất lượng, tạm giữ chậm tiến độ, phạt tổng tiến độ, trần phạt | |
| TT-60, TT-61 | Hồ sơ thanh toán mỗi đợt, hồ sơ quyết toán | |

Loại hợp đồng quyết định cách đọc khối lượng: **đơn giá** thì KL thực tế là căn cứ thanh toán, phần vượt KL HĐ là phát sinh cần phụ lục; **trọn gói** thì KL chỉ để đo % hoàn thành.

### 3.4. Phụ lục

**Nhập phụ lục bằng phiếu Excel**, mỗi phụ lục một sheet:

- Tạo phiếu: `6_tao_phieu_phu_luc.bat` → `csdl_hop_dong.py tao-phieu <HĐ>`. Phiếu tự điền số HĐ và số PL tiếp theo. Phiếu kèm sheet `BOQ hien hanh` và `Tham so hien hanh` để tra mã.
- **Thông tin chung** (ô vàng là bắt buộc): số PL, ngày ký, ngày hiệu lực (trống = ngày ký), loại, **Có đổi BOQ?**, GT tăng/giảm trước và sau VAT, GT HĐ sau PL, căn cứ, file PL, trạng thái.
- **Mục A – tham số bị sửa**: mã TT-xx, giá trị cũ, giá trị mới, trích dẫn nguyên văn. Ví dụ gia hạn tiến độ sửa TT-50. **Không ghi TT-02**: máy tự ghi thay đổi TT-02 từ ô giá trị tăng/giảm ở phần thông tin chung, để giá trị HĐ chỉ nhập một lần.
- **Mục B – BOQ**, chỉ khi "Có đổi BOQ?" = Có:
  - Gõ từng dòng THAY / THEM / BO; **hoặc**
  - Dán bảng BOQ của phụ lục vào một sheet khác, ghi tên sheet đó và chọn phạm vi. Máy tự suy ra THAY / THEM / BO bằng cách so với BOQ hiện hành:
    - "Toàn bộ BOQ sau PL": hạng mục không còn trong bảng bị coi là BO.
    - "Chỉ hạng mục thay đổi": không suy ra BO.
- Nạp: `7_nap_phu_luc.bat` → `csdl_hop_dong.py nap-phu-luc <phiếu>`. Máy kiểm tra trước khi ghi. Có lỗi CHẶN thì không ghi gì.
- Phụ lục **Dự thảo** được lưu để theo dõi nhưng không áp dụng khi đối chiếu.
- Nhập sai thì `huy-phu-luc <HĐ> <PL> --ly-do "…"` rồi nạp lại. Bản bị huỷ vẫn còn trong kho để truy vết.

**Kiểm tra khi nạp phụ lục**

| Mã | Mức | Kiểm tra |
|---|---|---|
| P01–P05 | CHẶN | Thiếu số PL / ngày ký / trạng thái / ô "Có đổi BOQ?"; số PL trùng; "Có" mà không có dòng BOQ (hoặc ngược lại); THAY/BO trỏ mã không có trong BOQ hiện hành; THEM thiếu tên, ĐVT, KL, đơn giá; mã tham số sai |
| P10 | ĐỎ | Σ thay đổi BOQ × (1+VAT) ≠ GT tăng/giảm ghi trên phụ lục |
| P11 | ĐỎ | GT HĐ trước PL + tăng/giảm ≠ GT HĐ sau PL ghi trên phụ lục (nối chuỗi PL01 → PL02 → …) |
| P12 | VÀNG | Số phụ lục không liên tục |
| P13 | ĐỎ / VÀNG | PL ký trước ngày ký HĐ / PL có hiệu lực hồi tố |
| P14 | ĐỎ | BO hạng mục đã thanh toán lũy kế > 0; THAY giảm KL xuống dưới lũy kế đã thanh toán |
| P15 | ĐỎ | Hạng mục đã nghiệm thu vượt KL **trước** ngày ký phụ lục tăng KL: hợp thức hoá sau |
| P16 | ĐỎ / VÀNG | Đổi đơn giá trong HĐ trọn gói / đơn giá cố định: không ghi căn cứ thì ĐỎ |
| P17 | VÀNG | GT HĐ đổi nhưng phụ lục không đổi BOQ (trừ HĐ trọn gói) |
| P18 | VÀNG | Phụ lục dự thảo |

### 3.5. Lập và cập nhật

1. **Hợp đồng mới**: `master_hd.py tao <hợp đồng> --boq <bảng chi tiết>` → người review xác nhận tham số trong master Excel → `0_nap_master_vao_CSDL.bat` nạp vào kho. Khi nạp, máy kiểm tra BOQ gốc × (1+VAT) = GT HĐ. Hợp đồng đã có trong `MAU_DANH_MUC_HOP_DONG.xlsx` thì dùng `nap-danh-muc`: thêm HĐ mới, còn HĐ đã có thì chỉ so sánh và báo chỗ lệch, không ghi đè.
2. **Đợt đầu**: nhập KL kỳ trước (thường 0) và tạm ứng đã chi nếu có.
3. **Ký phụ lục**: phiếu phụ lục (mục 3.4). Không sửa BOQ gốc.
4. **Sau mỗi đợt được duyệt**: `nap-dot <HĐ> <file đợt đã duyệt> --dot N --ngay-nghiem-thu dd/mm/yyyy`. Đợt nạp từ master chưa có ngày nghiệm thu thì bổ sung bằng `ngay-nghiem-thu <HĐ> <đợt> <ngày>`.
5. **Theo dõi**: `9_xuat_CSDL_ra_Excel.bat` xuất các sheet sau.
   - `HOP_DONG`: mỗi HĐ một dòng.
     - **Đủ 36 tham số** theo đúng thứ tự cột của `DANH_MUC_HD`, hiển thị giá trị **hiệu lực** tại ngày xuất.
     - Ô bị phụ lục sửa tô vàng. Rê chuột vào ô để xem điều khoản, trích dẫn nguyên văn và lịch sử gốc → PL.
     - Ô đỏ "THIẾU" là tham số chưa trích được.
     - Bên phải là các cột kiểm soát: PL áp dụng, GT gốc, GT hiệu lực, BOQ × (1+VAT), chênh lệch, cờ đỏ, số tham số chưa xác nhận.
   - `PHU_LUC`: mỗi phụ lục một dòng, gồm cả cột "Tham số thay đổi (cũ → mới)" và số dòng BOQ THAY / THEM / BO. Thay đổi tham số không tách thành sheet riêng.
   - `BOQ_HIEU_LUC`, `BOQ_DONG`, `DOT_THANH_TOAN`, `KL_THANH_TOAN`, `CANH_BAO`, `NHAT_KY`.

---

## 4. Quy trình review mỗi đợt

| Bước | Việc | Đầu ra |
|---|---|---|
| **B0** | Hợp đồng đã có trong CSDL, tham số bắt buộc đã xác nhận, mọi phụ lục đã ký đã nạp; đợt trước đã nạp kèm ngày nghiệm thu | CSDL cập nhật |
| **B1** | Kiểm tra đủ hồ sơ theo `Ho so TT theo HD` + danh mục mục 7 (app tab 1–2) | Danh sách request |
| **B2** | Đối chiếu hồ sơ với hợp đồng + phụ lục qua CSDL (`kiem-dot`): giá trị HĐ → BOQ hiệu lực → từng hạng mục | Kết luận ĐẠT / CẦN PHỤ LỤC / SAI, danh sách yêu cầu phụ lục |
| **B3** | Phân rã công thức, tìm số nhập tay nằm ngoài 7 nhóm gốc (khi B2 có lệch hoặc file có cấu trúc lạ) | Cây truy vết, danh sách số nhập tay |
| **B4** | KL kỳ này: đối chiếu biên bản nghiệm thu, chọn mẫu kiểm tra bản vẽ, hoàn công | Giấy tờ làm việc KL |
| **B5** | Khấu trừ theo HĐ: giữ lại, thu hồi tạm ứng, tạm giữ và phạt tiến độ (đối chiếu biên bản vi phạm), VAT | Bảng khấu trừ |
| **B6** | Tổng hợp phát hiện, số kiến nghị thanh toán; sau khi đợt được duyệt thì nạp vào CSDL (`nap-dot`) | Báo cáo + CSDL cập nhật |

### B2. Nguyên tắc đối chiếu với hợp đồng và phụ lục

**Hợp đồng và các phụ lục đã ký là căn cứ duy nhất. Mọi sai lệch giữa hồ sơ thanh toán và hợp đồng đều phải có phụ lục.** Sai lệch không có phụ lục thì yêu cầu bổ sung phụ lục, phần giá trị chênh lệch tạm chưa thanh toán.

Đối chiếu theo thứ tự sau; bước sau dùng kết quả bước trước.

**Bước 1. Giá trị hợp đồng**

| Kết quả | Điều kiện | Xử lý |
|---|---|---|
| ĐẠT | GT HĐ trong hồ sơ = GT HĐ gốc + Σ GT tăng/giảm của các phụ lục đã ký có hiệu lực ≤ ngày nghiệm thu | Sang bước 2 |
| CẦN PHỤ LỤC | Khác, và không có phụ lục nào giải thích phần chênh | Yêu cầu phụ lục; chưa thanh toán phần chênh |
| PHỤ LỤC CHƯA HIỆU LỰC | Hồ sơ đã dùng GT sau một phụ lục dự thảo, hoặc phụ lục có hiệu lực sau ngày nghiệm thu | Không áp phụ lục đó cho đợt này; yêu cầu sửa hồ sơ |
| HỒ SƠ CHƯA CẬP NHẬT PHỤ LỤC (SAI) | Phụ lục đã có hiệu lực trước ngày nghiệm thu nhưng hồ sơ vẫn dùng GT cũ | Yêu cầu sửa hồ sơ |

**Bước 2. Lập BOQ hiệu lực cho đợt này**

- Bắt đầu từ BOQ gốc của hợp đồng.
- Áp lần lượt các phụ lục **đã ký, có ngày hiệu lực ≤ ngày nghiệm thu của đợt**, theo thứ tự ngày hiệu lực. Mốc là ngày nghiệm thu vì đó là thời điểm phát sinh nghĩa vụ thanh toán. Phụ lục ký sau ngày nghiệm thu mà làm hợp lệ khối lượng đã nghiệm thu thì bị gắn cờ "hợp thức hoá sau" (P15).
  - **THAY**: hạng mục có trong phụ lục thì lấy KL, đơn giá theo phụ lục, bỏ giá trị cũ.
  - **THÊM**: hạng mục mới trong phụ lục được thêm vào BOQ.
  - **BỎ**: hạng mục bị bỏ hoặc bị thay thế thì KL HĐ = 0.
- Hạng mục không có trong phụ lục nào thì giữ nguyên theo BOQ gốc.
- Tự kiểm tra: Σ(BOQ hiệu lực) × (1 + VAT) phải bằng GT HĐ sau phụ lục ở bước 1 (sheet `HOP_DONG` của bản xuất, cột "Chênh BOQ - GT"). Lệch nghĩa là phụ lục nhập chưa đúng hoặc chính phụ lục có sai số học (cờ P10).

**Bước 3. Đối chiếu từng hạng mục với BOQ hiệu lực, theo loại hợp đồng (TT-01)**

| Kiểm tra | Trọn gói | Đơn giá cố định | Đơn giá điều chỉnh |
|---|---|---|---|
| Hạng mục có trong BOQ hiệu lực | Bắt buộc | Bắt buộc | Bắt buộc |
| KL HĐ trong hồ sơ = BOQ hiệu lực | Bắt buộc | Bắt buộc | Bắt buộc |
| Đơn giá trong hồ sơ = BOQ hiệu lực | Bắt buộc | Bắt buộc | = đơn giá BOQ đã điều chỉnh theo phụ lục điều chỉnh giá |
| KL lũy kế ≤ KL HĐ | Bắt buộc; thanh toán theo % hoàn thành, phần vượt không thanh toán | Vượt thì phải có phụ lục hoặc biên bản phát sinh trước khi thanh toán phần vượt | Như đơn giá cố định |

Kết luận cho từng hạng mục:

| Kết luận | Khi nào |
|---|---|
| ĐẠT | Khớp BOQ hiệu lực (sai số làm tròn: đơn giá ≤ 0,5 đ, KL ≤ 0,001) |
| CẦN PHỤ LỤC | Lệch BOQ hiệu lực và không có phụ lục nào nói đến hạng mục này; hoặc hạng mục không có trong HĐ và phụ lục; hoặc KL lũy kế vượt KL HĐ (HĐ đơn giá) |
| SAI | Hạng mục có trong phụ lục nhưng hồ sơ vẫn khác phụ lục |

Giá trị ảnh hưởng của mỗi dòng không ĐẠT = GT lũy kế theo hồ sơ − GT lũy kế tính theo BOQ hiệu lực. Tổng các dòng CẦN PHỤ LỤC là số **tạm chưa thanh toán** của đợt này. Đầu ra là **danh sách yêu cầu phụ lục**: hạng mục, loại lệch, giá trị ảnh hưởng.

### B2a. Các kiểm tra tự động khi đối chiếu

**Theo từng hạng mục** (khớp bằng nhóm + tên + ĐVT, không phụ thuộc số dòng):

| Kiểm tra | Nguồn so sánh | Ảnh hưởng tính ra |
|---|---|---|
| KL HĐ trong hồ sơ = KL HĐ master | Nhóm 1 | |
| Đơn giá trong hồ sơ = đơn giá master | Nhóm 2 | (ĐG hồ sơ − ĐG master) × KL lũy kế |
| KL kỳ trước trong hồ sơ = KL lũy kế đợt trước đã duyệt | Nhóm 3 | (KL kỳ trước hồ sơ − master) × đơn giá |
| KL lũy kế ≤ KL HĐ | Nhóm 1 | Phần vượt cần phụ lục (HĐ đơn giá) |
| Hạng mục có trong hồ sơ nhưng không có trong master | | Phát sinh chưa có phụ lục |
| Hạng mục có trong master nhưng không có trong hồ sơ | | Bảng bị xoá dòng |

**Theo số tổng hợp** (tìm theo nhãn, không phụ thuộc vị trí ô):

| Kiểm tra | So với |
|---|---|
| Giá trị HĐ trong hồ sơ | TT-02 |
| Giữ lại / GT thực hiện kỳ này | TT-21 |
| GT lũy kế đến hết kỳ trước | Lũy kế đợt trước trong `Lich su TT` |
| Σ(KL kỳ trước × đơn giá master) × (1+VAT) | Lũy kế đợt trước (kiểm tra được cả khi chưa có KL kỳ trước từng hạng mục) |
| Σ(KL kỳ này × đơn giá master) × (1+VAT) | GT thực hiện kỳ này trong hồ sơ |

### B3. Khi nào cần phân rã công thức

Chạy khi B2 có lệch mà không giải thích được, hoặc khi file có liên kết ra file ngoài, sheet ẩn, bảng diễn giải khối lượng phức tạp. Công cụ lần từ ô đề nghị thanh toán xuống các ô gốc, đánh dấu số nhập tay, tham số nhúng trong công thức, liên kết ngoài (ví dụ ở tài liệu nội bộ).

### B4. Khối lượng kỳ này

- Mỗi hạng mục có KL kỳ này phải nằm trong biên bản nghiệm thu đã ký của đợt, đúng KL.
- Chọn mẫu theo mục 6 để kiểm tra bản vẽ thi công, hoàn công, bảng tính khối lượng thiết kế. Bản vẽ CAD (DWG/DXF) có thể đọc bằng Python; bản vẽ PDF cần người đo.
- Hạng mục vật tư, thiết bị (nếu HĐ có): cần thêm biên bản giao nhận, nghiệm thu vật liệu đầu vào, CO/CQ.

---

## 5. Phân vai: Python, local AI, người review

**Nguyên tắc:** Local AI chỉ **đọc tài liệu không có cấu trúc và đề xuất**. Mọi phép tính, đối chiếu số, chọn mẫu, kết luận do Python và người review.

| Bước (mục 4 / file phương pháp) | Python | Local AI | Người |
|---|---|---|---|
| B0 Hồ sơ đủ chưa | Phân loại file theo từ khoá, tình trạng theo danh mục, danh sách request | Phân loại file khi từ khoá không chắc (kèm trích dẫn) | Sửa phân loại sai |
| B1 Tham số HĐ, logic bảng TH | 36 tham số hiệu lực; kiểm tỷ lệ giữ lại, tạm ứng, GT HĐ | Trích tham số HĐ, phụ lục (`master_hd.py ai-trich`) | Xác nhận tham số |
| B2 Toàn vẹn file | 100%: công thức, số nhập tay, liên kết ngoài, #REF!, sheet ẩn, cây truy vết | – | Hỏi nhà thầu |
| B3 Đơn giá, BOQ, phụ lục | Đối chiếu BOQ hiệu lực, ảnh hưởng | Chép bảng BOQ / phụ lục dạng PDF, scan; tìm căn cứ hệ số đơn giá trong biên bản thương thảo | Kiểm bảng AI chép |
| B4 Kỳ trước | Đối chiếu lũy kế đợt đã duyệt | Chép bảng hồ sơ đợt trước bản ký (scan); đọc UNC, hoá đơn | Kiểm mẫu |
| B5 Chứng từ KL kỳ này | Chọn mẫu XD; ma trận thiết bị × chứng từ; so số lượng | **Đọc BB giao nhận, packing list, CO/CQ, BB lắp đặt, chạy thử, BB nghiệm thu → bảng dòng hàng**; gợi ý ghép tên với hạng mục; đánh dấu thiếu ngày, chữ ký | Đối chiếu chứng từ gốc, xác nhận; đo bản vẽ |
| B6 Tổng hợp | Phân loại (1)(2)(3), trọng yếu, tổng ảnh hưởng | Nháp lời văn phát hiện, nháp thư request | Chốt phát hiện, số kiến nghị |

**Kiểm soát khi dùng AI** (đã lập trình trong `review_ttkl/ai.py`):
1. Ẩn danh tên các bên, MST, số tài khoản, điện thoại trước khi gửi; chỉ gửi đoạn cần đọc; qua tunnel tới pod riêng.
2. Mỗi dòng AI trích phải có trích dẫn nguyên văn; Python tìm lại trong văn bản, không thấy thì gắn "KHÔNG - đọc lại chứng từ".
3. Kết quả AI ở trạng thái "Chưa xác nhận"; chỉ dòng người review "Đã xác nhận" mới là căn cứ (dòng chưa xác nhận chỉ để gợi ý, tô vàng).
4. Lời văn AI soạn: có con số không nằm trong dữ liệu đưa vào thì bỏ, dùng mô tả của Python.
5. Mỗi lần gọi ghi nhật ký (việc, token, thời gian) để đánh giá POC. AI tắt thì mọi bước vẫn chạy bằng Python.
6. Đo trước khi tin: mỗi loại tài liệu cần bộ đáp án để chấm (như `master_hd.py cham` với hợp đồng).

---

## 6. Chọn mẫu và trọng yếu

- **Kiểm tra 100%** bằng công cụ: mọi kiểm tra ở B2, B3.
- **Kiểm tra chứng từ 100%**: hạng mục có GT kỳ này ≥ ngưỡng trọng yếu thực hiện; hạng mục bị cờ ở B2; hạng mục vật tư thiết bị ghi nhận 100% KL trong kỳ.
- **Chọn mẫu theo giá trị** cho phần còn lại, tối thiểu 25 hạng mục hoặc đủ phủ 80% GT kỳ này.

Ngưỡng đề xuất (chốt với người giao việc):

| Ngưỡng | Cơ sở |
|---|---|
| Trọng yếu tổng thể | 1% GT thực hiện kỳ này (có VAT) |
| Trọng yếu thực hiện | 75% trọng yếu tổng thể |
| Sai sót không đáng kể | 5% trọng yếu tổng thể |

Phân loại phát hiện: (1) sai sót xác định, trừ vào số kiến nghị; (2) chưa đủ căn cứ, tạm chưa thanh toán; (3) khuyến nghị kiểm soát.

---

## 7. Danh mục tài liệu

Danh mục chung dưới đây nằm trong `app_kiem_tra_ho_so/danh_muc_tai_lieu.csv`, sửa được bằng Excel. Ngoài ra, mỗi HĐ có danh sách hồ sơ thanh toán riêng trong sheet `Ho so TT theo HD` của master.

| Nhóm | Tài liệu | Dùng cho nhóm dữ liệu gốc |
|---|---|---|
| A. Hợp đồng | Hợp đồng + điều kiện; Bảng chi tiết giá trị HĐ (BOQ); phụ lục; văn bản chấp thuận thay đổi; bảo lãnh (tạm ứng, thực hiện, bảo hành); tiến độ được duyệt | 1, 2, 6 |
| B. Thiết kế | Bản vẽ thi công được duyệt; bảng tính KL thiết kế; bản vẽ hoàn công; hồ sơ thay đổi thiết kế, phát sinh | 1, 5 |
| C. Nghiệm thu | BB nghiệm thu KL hoàn thành đợt; BB xác nhận KL và GT hoàn thành; BB nghiệm thu công việc, vật liệu đầu vào; nhật ký thi công; kết quả thí nghiệm; ảnh hiện trường | 5 |
| D. Vật tư, thiết bị (nếu có) | BB giao nhận; CO/CQ; packing list; BB lắp đặt, chạy thử | 5 |
| E. Thanh toán | Hồ sơ các đợt trước đã duyệt (bản ký + Excel); UNC / sổ phụ; hóa đơn GTGT; đối chiếu công nợ; công văn đề nghị thanh toán đợt này | 3, 4, 7 |
| F. Khác | Biên bản vi phạm (chất lượng, tiến độ) làm căn cứ khấu trừ; căn cứ thuế suất | 6 |

---

## 8. Cách chạy công cụ

```bash
# 1. Tạo master (một lần cho mỗi hợp đồng)
python master_hd.py tao "HĐ 0101.doc" --boq "Bang chi tiet gia tri HD.xlsx" --out master_0101.xlsx
#    thêm --ai để nhờ Ollama trích các mục quy tắc không bắt được

# 2. Mở master, xác nhận tham số; đợt đầu nhập KL kỳ trước (thường 0)

# 3. Mỗi đợt: đối chiếu hồ sơ nhà thầu với master
python master_hd.py doi-chieu master_0101.xlsx "HSTT dot 3.xlsx" --out doi_chieu_dot3.xlsx

# 4. Khi cần đào sâu: phân rã công thức, quét số nhập tay (mục B3)
python hardcode_audit.py "HSTT dot 3.xlsx" ket_qua.xlsx

# 4b. Thử AI: trích hợp đồng bằng model trên RunPod (qua tunnel localhost:11434) và so với quy tắc
python master_hd.py ai-trich "HĐ 0101.doc" --an-danh --model qwen3.6:35b

# 4c. Chấm kết quả trích xuất với đáp án (.tsv): hợp đồng -> chấm quy tắc; file .tsv của AI -> chấm AI
python master_hd.py cham DAP_AN_tai-lieu-001.tsv tai-lieu-001.docx
python master_hd.py cham DAP_AN_tai-lieu-001.tsv ket_qua_AI.tsv

# 5. Sau khi đợt được duyệt: nạp vào master làm kỳ trước cho đợt sau
python master_hd.py nap-dot master_0101.xlsx "HSTT dot 3 (ban duyet).xlsx"
```

**Cơ sở dữ liệu** (từ v2.2 dùng thay cho bước 3 và 5 ở trên; mỗi lệnh có file .bat tương ứng):

```bash
python csdl_hop_dong.py nap-master master_HD_0101.xlsx            # 0_nap_master_vao_CSDL.bat
python csdl_hop_dong.py nap-danh-muc MAU_DANH_MUC_HOP_DONG.xlsx   # HĐ đã có thì chỉ so sánh
python csdl_hop_dong.py tao-phieu 0101                            # 6_tao_phieu_phu_luc.bat
python csdl_hop_dong.py nap-phu-luc phieu_phu_luc_0101.xlsx       # 7_nap_phu_luc.bat  (--chi-kiem-tra: thử, không ghi)
python csdl_hop_dong.py huy-phu-luc 0101 PL01 --ly-do "nhập sai"
python csdl_hop_dong.py boq 0101 --ngay 10/07/2026                # in BOQ hiệu lực tại ngày
python csdl_hop_dong.py ngay-nghiem-thu 0101 2 10/07/2026         # bổ sung ngày NT cho đợt nạp từ master
python csdl_hop_dong.py kiem-dot 0101 "HSTT dot 3.xlsx" --ngay-nghiem-thu 20/08/2026 [--dot 3]   # 8_kiem_tra_dot_thanh_toan.bat
python csdl_hop_dong.py nap-dot 0101 "HSTT dot 3 (ban duyet).xlsx" --dot 3 --ngay-nghiem-thu 20/08/2026
python csdl_hop_dong.py xuat-excel [0101] [--ngay 30/09/2026]     # 9_xuat_CSDL_ra_Excel.bat
```

**Module 2 – Lớp 1: tính toàn vẹn file** (mục 11; kéo thả file vào `13_kiem_toan_ven.bat`):

```bash
python toan_ven.py "HSTT dot 3.xlsx" [--o "2. TH.!D38"] [--out ket_qua.xlsx]
```

Kết quả `kiem-dot` gồm các sheet:
- `Tong hop`: bước 1 đến 3, Σ KL kỳ này × đơn giá, lũy kế kỳ trước, **giá trị tạm chưa thanh toán**.
- `Hang muc`: mọi dòng, có cột nguồn (HĐ gốc / PL số) và lịch sử BOQ.
- `Yeu cau phu luc`: danh sách gửi nhà thầu.
- `Thieu trong ho so`.
- `Phu luc`: phụ lục nào áp dụng, phụ lục nào không áp dụng cho đợt này.

Cài đặt: `pip install openpyxl python-docx pdfplumber`. File `.doc` cần LibreOffice hoặc MS Word trên máy để chuyển sang `.docx`.

App (`app_kiem_tra_ho_so`, chạy bằng `chay_app.bat`) gồm các tab: upload và phân loại, tình trạng hồ sơ, hỏi đáp, quét số nhập tay, và **master hợp đồng / đối chiếu**.

---

## 9. Thiết kế chương trình đối chiếu hợp đồng – phụ lục – hồ sơ

*Đã lập trình trong `csdl_hop_dong.py` (v2.2). Dữ liệu ở mục 3.1.*

### 9.1. Các hàm chính

| Hàm | Việc |
|---|---|
| `phu_luc_ap_dung(hd, ngay)` | Phụ lục đã ký có ngày hiệu lực ≤ ngày, theo thứ tự hiệu lực |
| `boq_hieu_luc(hd, ngay)` | BOQ gốc + lần lượt THAY / THEM / BO; trả về BOQ hiệu lực (kèm nguồn, lịch sử) và danh sách hạng mục đã bỏ |
| `gt_hd_hieu_luc(hd, ngay)` | GT HĐ gốc (sau VAT) + Σ tăng/giảm sau VAT của phụ lục áp dụng |
| `tham_so_hieu_luc(hd, ma, ngay)` | Giá trị tham số gốc hoặc theo phụ lục gần nhất |
| `kiem_tra_phu_luc(...)` | Các kiểm tra P01–P18 (mục 3.4) trước khi ghi |
| `_khop_hstt(...)` | Ghép dòng hồ sơ với Mã HM theo thứ tự: bảng đã xác nhận `khop_hang_muc` → khoá nhóm + tên + ĐVT → tên + ĐVT duy nhất → gợi ý gần đúng ≥ 85% (chỉ gợi ý, không tự ghép) |
| `kiem_dot(...)` | Bước 1–3 mục B2. Ảnh hưởng mỗi dòng = GT lũy kế theo hồ sơ − min(KL lũy kế, KL hiệu lực) × đơn giá hiệu lực |

### 9.2. Thử nghiệm

Đã chạy thử trên 3 hợp đồng thật và các phụ lục giả lập. Các ca thử gồm: phụ lục đúng; phụ lục sai cố ý (bỏ hạng mục đã thanh toán, đổi đơn giá không căn cứ, giá trị lệch); phụ lục trỏ mã không tồn tại (bị chặn); phụ lục dán BOQ ở sheet riêng; hợp thức hoá sau; huỷ phụ lục; kiểm tra một hồ sơ trước và sau ngày hiệu lực phụ lục. Mọi kiểm tra cho kết quả như mong đợi. Kết quả chi tiết theo từng hợp đồng là dữ liệu nội bộ, không đưa lên kho mã.

---

## 10. Giới hạn hiện tại

- Quy tắc trích hợp đồng được viết và thử trên khung hợp đồng của 3 HĐ mẫu (tài liệu nội bộ). Hợp đồng viết theo khung khác (ví dụ FIDIC, mẫu Bộ Xây dựng) có thể có mục không bắt được. Mục đó hiện ô đỏ để nhập tay, hoặc dùng `--ai`.
- Đọc bảng chi tiết dựa vào tiêu đề cột ("Tên công việc", "Đơn vị", "Đơn giá", "kỳ trước", "kỳ này", "lũy kế"). Bảng không có cột đơn giá, hoặc tách khối lượng và giá trị ra hai sheet, chưa đọc được.
- Khớp hạng mục giữa các đợt dựa vào nhóm + tên + ĐVT. Nếu nhà thầu đổi tên hạng mục, hạng mục đó hiện là "không có trong BOQ hiệu lực" kèm gợi ý mã gần đúng, để người review ghép tay.
- Mốc ngày nghiệm thu đang tính **theo đợt**. Một đợt gồm nhiều biên bản nghiệm thu khác ngày thì lấy ngày của biên bản cuối cùng, và xem riêng các hạng mục chịu tác động của phụ lục ký giữa các ngày đó.
- App Streamlit chưa nối với CSDL và khung review; hiện chạy bằng các file .bat.
- `hardcode_audit.py` (bản đầu, viết theo mẫu một hồ sơ) được thay bằng `toan_ven.py` cho mọi mẫu; giữ lại để tham khảo.
- `hardcode_audit.py` phần lượng hoá chi tiết được viết theo một mẫu hồ sơ cụ thể (tên sheet khai báo ở đầu file); với mẫu khác chỉ dùng phần cây truy vết.
- Phần AI (`ai-trich`, `--ai`) đã thử với máy chủ giả lập; chưa chạy với model thật trên RunPod.

---

## 11. Module 2 – Lớp 1: tính toàn vẹn file (`toan_ven.py`)

Module 1 (trích hợp đồng + cơ sở dữ liệu hợp đồng – phụ lục – BOQ, mục 3) là **dữ liệu cho Lớp 2**. Module 2 kiểm **Lớp 1** trên chính file Excel hồ sơ thanh toán, không cần hợp đồng, chạy với mọi mẫu file (không khai báo tên sheet, cột). Tương ứng bước 2 của quy trình 6 bước, đầu ra là WP-C.

| Mã | Kiểm tra | Mức khi có phát hiện |
|---|---|---|
| K1 | Tổng quan: sheet ẩn; liên kết file ngoài, tách loại **được công thức dùng** và loại **chỉ còn trong tên định nghĩa (rác)**; ô đang lỗi (#REF!, #DIV/0!…); công thức chưa có giá trị lưu (file chưa tính lại) | Cao nếu file ngoài nằm trên đường tính số đề nghị; Trung bình nếu chỉ ở cột phụ |
| K2 | **Bản đồ công thức theo cột**: mẫu công thức chính của từng cột (chuẩn hoá kiểu R1C1), ô lệch mẫu, **số gõ tay nằm giữa cột công thức** | Cao (gõ tay), Trung bình (lệch mẫu) |
| K3 | **Liên kết dòng**: công thức trỏ sang sheet khác có trỏ đúng hạng mục không. So tên hạng mục hai đầu; chỉ xét cặp sheet liên kết theo dòng có hệ thống (≥ 20 ô), bỏ liên kết vào dòng tổng | Cao |
| K4 | Sheet ẩn: bị sheet hiện tham chiếu bao nhiêu ô; so từng ô với sheet hiện cùng tên ("… (2)") | Trung bình |
| K5 | Số lưu dạng chữ (Excel không cộng vào tổng) | Trung bình |
| K6 | Hằng số gõ cứng trong công thức (×1,08; ×60%…), chỉ cảnh báo khi nằm trên cột thuộc cây truy vết | Trung bình (Lớp 2 đối chiếu với điều khoản HĐ) |
| K7 | **Cây truy vết**: số đề nghị thanh toán = dữ liệu nào × dữ liệu nào. Tự tìm ô "giá trị đề nghị thanh toán kỳ này"; dải nhiều ô phân rã một ô đại diện cho mỗi mẫu công thức; **mỗi lá nhập tay được xếp vào 7 nhóm dữ liệu gốc**, lá "Ngoài 7 nhóm" là cờ | Cao |
| K8 | **Nguồn KL lũy kế và KL kỳ trước của từng hạng mục**: tính từ diễn giải / nhập tay / bằng KL HĐ / tỷ lệ × KL HĐ / link lệch dòng / nguồn trống / qua sheet ẩn; phát hiện diễn giải vượt KL HĐ đang bị công thức chặn trần. Bảng phân tầng dùng cho chọn mẫu ở bước 5 | Cao / Trung bình |

Cách lần nguồn KL (K8): đi theo tham chiếu một ô qua các sheet; gặp công thức cộng/trừ nhiều ô thì đi theo thành phần "lũy kế", rồi "kỳ này"; gặp IF/MIN/MAX thì đi theo nhánh thực sự cho ra giá trị (so giá trị đã lưu). Mỗi bước qua sheet khác đều so tên hạng mục để bắt link lệch dòng.

**Kiểm chứng:** chạy trên hồ sơ mẫu đã dùng để xây dựng file phương pháp ban đầu, module bắt lại được các dấu hiệu đã biết từ trước. Gồm: file ngoài nằm trên đường tính số đề nghị, liên kết lệch dòng sang sheet diễn giải, sheet ẩn bị tham chiếu và khác bản hiện, #REF!, số dạng chữ, số đợt trước và tỷ lệ gõ cứng, phân tầng nguồn KL khớp tổng lũy kế. Thêm 2 hồ sơ của mẫu file khác, mỗi file chạy khoảng 25 giây. Chi tiết số liệu ở file nội bộ.

Giới hạn: số liệu lấy theo giá trị Excel đã lưu (file chưa tính lại thì K1 cảnh báo); nhận diện tiêu đề cột, cột tên hạng mục bằng quy tắc nên mẫu lạ có thể xếp nhóm sai, khi đó xem cột "Đường đi" ở K8 và cây truy vết để kiểm.

> Khung kiểm tra đối chiếu đầy đủ B0–B6 (đã dựng thử) để ở thư mục `_de_sau\`, chưa dùng, sẽ ghép lại khi làm các module sau.

---

*Ví dụ áp dụng trên các hợp đồng thật (số liệu, kết quả chạy thử, tình trạng dữ liệu) nằm trong file nội bộ `quy-trinh_NOI_BO_vi_du_hop_dong.md`, không đưa lên kho mã công khai.*
