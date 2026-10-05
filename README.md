# Review hồ sơ thanh toán khối lượng xây dựng

Bộ công cụ hỗ trợ review hồ sơ thanh toán khối lượng xây dựng theo cách của auditor: đi từ số đề nghị thanh toán ngược về dữ liệu gốc. Dữ liệu gốc gồm hợp đồng, BOQ, phụ lục, đợt trước và biên bản nghiệm thu. Nguồn nào thiếu thì yêu cầu bổ sung.

Phương pháp đầy đủ: [`quy-trinh-review-ttkl.md`](quy-trinh-review-ttkl.md).

## Thành phần

| File | Dùng để |
|---|---|
| `master_hd.py` | Trích 36 tham số hợp đồng (TC-01 … TT-62) bằng quy tắc, có trích dẫn nguyên văn; tuỳ chọn nhờ local AI (Ollama); chấm điểm với đáp án |
| `csdl_hop_dong.py` | Cơ sở dữ liệu SQLite hợp đồng – phụ lục – BOQ: nạp phụ lục bằng phiếu Excel, BOQ hiệu lực theo ngày nghiệm thu, kiểm tra hồ sơ đợt mới, xuất Excel |
| `toan_ven.py` | Module 2 – Lớp 1: kiểm tra tính toàn vẹn file hồ sơ thanh toán (mọi mẫu Excel): công thức lệch mẫu, số gõ tay giữa cột công thức, liên kết lệch dòng, sheet ẩn, file ngoài, cây truy vết số đề nghị theo 7 nhóm dữ liệu gốc, nguồn KL từng hạng mục |
| `trich_lo_hop_dong.py` | Trích tham số cho cả lô hợp đồng nhờ máy Local AI trong mạng LAN; gộp quy tắc + AI, ra master từng HĐ và file tổng hợp |
| `giao_dien_trich_hd.py`, `may_ai.py` | **Một cửa sổ** trích tham số hợp đồng bằng Local AI (mở bằng `5_trich_hop_dong_AI.bat`): tab *Trích hợp đồng* (chọn / kéo thả nhiều hợp đồng, tiến độ từng nhóm, dừng, mở kết quả) và tab *Máy AI* (nhập / tự dò IP, kiểm tra, chọn model) |
| `hardcode_audit.py` | Phân rã công thức, tìm số nhập tay trong file hồ sơ |
| `app_kiem_tra_ho_so/` | App Streamlit chạy trên máy: kiểm tra đủ hồ sơ, hỏi đáp, gọi các công cụ trên |
| `0_…` – `13_….bat` | Chạy một nút trên Windows (kéo thả file); `5_` mở cửa sổ trích hợp đồng bằng Local AI |
| `cau_lenh_local_ai_trich_hop_dong.md` | Câu lệnh cho local AI trích hợp đồng |

## Cài đặt

```bash
pip install openpyxl python-docx pdfplumber pywin32
```

File `.doc` cần LibreOffice hoặc MS Word để chuyển sang `.docx`. Cơ sở dữ liệu dùng `sqlite3` có sẵn trong Python.

## Dữ liệu

Kho này **chỉ chứa mã nguồn và tài liệu chung**. File `.gitignore` liệt kê đích danh các file được phép đẩy lên (whitelist), nên những thứ sau không bao giờ lên kho:

- hợp đồng, hồ sơ thanh toán;
- master file;
- cơ sở dữ liệu (`co_so_du_lieu/`);
- kết quả chạy;
- ví dụ áp dụng trên hợp đồng thật;
- thông tin máy chủ AI (`cau_hinh_may_ai.json`), thư mục `hop_dong_can_trich/`, `ket_qua_AI/`;
- thư mục `_de_sau/`, `vu_viec/` và cấu hình riêng (`cau_hinh_review.json`).
