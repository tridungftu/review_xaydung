# -*- coding: utf-8 -*-
"""
toan_ven.py - MODULE 2: kiểm tra LỚP 1 - TÍNH TOÀN VẸN của file Excel hồ sơ thanh toán (mọi mẫu).

Lớp 1 (file phương pháp review):
    - Công thức tính đúng và nhất quán theo cột
    - Liên kết giữa các sheet trỏ đúng dòng, đúng hạng mục
    - Không có số nhập tay ở vị trí phải là công thức
    - Không có liên kết ra file ngoài, không có #REF!, sheet ẩn đều được giải thích
Bước 2 của quy trình 6 bước -> WP-C, kèm "bảng phân loại nguồn của từng dòng KL" (dùng cho bước 5 chọn mẫu).

Các kiểm tra:
    K1  Tổng quan workbook: sheet ẩn, liên kết file ngoài (dùng trong công thức / chỉ nằm trong tên rác),
        tên định nghĩa lỗi, ô lỗi (#REF!, #DIV/0!...), công thức chưa có giá trị lưu
    K2  Bản đồ công thức theo cột: mẫu công thức chính của mỗi cột, ô lệch mẫu, số nhập tay giữa cột công thức
    K3  Liên kết dòng: công thức trỏ sang sheet khác có trỏ đúng hạng mục không (so tên dòng hai đầu)
    K4  Sheet ẩn: được sheet hiện tham chiếu bao nhiêu ô; khác gì so với sheet hiện cùng tên
    K5  Số lưu dạng chữ
    K6  Tham số gõ cứng trong công thức (x1,08; x60%...)
    K7  Cây truy vết: số đề nghị thanh toán = dữ liệu nào x dữ liệu nào; lá nhập tay xếp vào 7 nhóm dữ liệu gốc
    K8  Nguồn khối lượng từng hạng mục: tính từ diễn giải / nhập tay / bằng KL HĐ / tỷ lệ x KL HĐ / link lệch / trống
Không tính lại bảng tổng hợp. Chỉ đọc file, không sửa.

Cách dùng:
    python toan_ven.py "HSTT dot 3.xlsx" [--o "2. TH.!D38"] [--out ket_qua.xlsx]
    (--o: ô đích để dựng cây truy vết; bỏ trống thì máy tự tìm ô "giá trị đề nghị thanh toán kỳ này")
"""
import argparse
import difflib
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from master_hd import bo_dau, nfc, num, tim_bang_chi_tiet, tim_so_tong_hop  # noqa: E402

LOI = ("#REF!", "#DIV/0!", "#VALUE!", "#N/A", "#NAME?", "#NUM!", "#NULL!")
# tham chiếu có tên sheet (kể cả file ngoài [n]) và tham chiếu trong sheet
QREF = re.compile(r"(?:'((?:[^']|'')+)'|((?:\[\d+\])?[^\s'!(),;:+\-*/^&=<>{}\"]+))!"
                  r"(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?")
LREF = re.compile(r"(?<![A-Za-z0-9_.!'\]$])(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?(?![\w(!])")
CHUOI = re.compile(r'"(?:[^"]|"")*"')
SO_NHUNG = re.compile(r"(?<![\w.$\[])(\d+(?:\.\d+)?%?)(?![\w.\[])")

MUC = {"Cao": 1, "Trung bình": 2, "Kiểm soát": 3}


# ============================================================== mô hình workbook
class SoDo:
    def __init__(self, path):
        self.path = Path(path)
        self.wb = openpyxl.load_workbook(path, data_only=False)
        self.wv = openpyxl.load_workbook(path, data_only=True)
        self.ngoai = {}
        for i, lk in enumerate(self.wb._external_links, 1):
            self.ngoai[i] = (lk.file_link.Target if lk.file_link else "?")
        self._cot_ten, self._hdr, self._merged = {}, {}, {}

    # ---------------------------------------------------------- truy cập ô
    def f(self, sh, c, r):
        try:
            return self.wb[sh].cell(r, c).value
        except KeyError:
            return None

    def v(self, sh, c, r):
        try:
            return self.wv[sh].cell(r, c).value
        except KeyError:
            return None

    def an(self, sh):
        return sh in self.wb.sheetnames and self.wb[sh].sheet_state != "visible"

    # ---------------------------------------------------------- tên dòng, tên cột
    def cot_ten(self, sh):
        """Cột chứa tên hạng mục: cột (trong 10 cột đầu) có nhiều ô chữ dài nhất."""
        if sh not in self._cot_ten:
            ws = self.wv[sh]
            dem = Counter()
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 3000), max_col=min(ws.max_column, 10)):
                for c in row:
                    if isinstance(c.value, str) and len(c.value.strip()) >= 8 and re.search(r"[A-Za-zÀ-ỹ]{3}", c.value):
                        dem[c.column] += 1
            self._cot_ten[sh] = dem.most_common(1)[0][0] if dem else 2
        return self._cot_ten[sh]

    def nhan_dong(self, sh, r):
        if sh not in self.wv.sheetnames:
            return ""
        v = self.v(sh, self.cot_ten(sh), r)
        return re.sub(r"\s+", " ", nfc(v)).strip()[:120] if v not in (None, "") else ""

    def _gop(self, sh):
        if sh not in self._merged:
            m = {}
            for rg in self.wb[sh].merged_cells.ranges:
                if rg.min_row > 60:
                    continue
                for rr in range(rg.min_row, rg.max_row + 1):
                    for cc in range(rg.min_col, rg.max_col + 1):
                        m[(rr, cc)] = (rg.min_row, rg.min_col)
            self._merged[sh] = m
        return self._merged[sh]

    def tieu_de(self, sh, c):
        """Tên cột: dòng tiêu đề (dòng có nhiều ô chữ nhất trong 40 dòng đầu) + 2 dòng dưới, có xử lý ô gộp."""
        if sh not in self._hdr:
            ws = self.wv[sh]
            best, h = (0, 0), 1
            kw = re.compile(r"^(stt|tt\b|ten|noi dung|hang muc|dien giai|don vi|dvt|khoi luong|kl\b|don gia|thanh tien|"
                            r"gia tri|ghi chu|ma hieu|so luong|luy ke|k[iy] nay|k[iy] truoc)")
            for r in range(1, min(ws.max_row, 40) + 1):
                cells = [ws.cell(r, cc).value for cc in range(1, min(ws.max_column, 40) + 1)]
                txt = [bo_dau(v) for v in cells if isinstance(v, str) and len(v.strip()) >= 2]
                diem = (sum(1 for t in txt if kw.match(t)), len(txt))
                if diem > best:
                    best, h = diem, r
            self._hdr[sh] = h
        h, m = self._hdr[sh], self._gop(sh)
        ws = self.wv[sh]
        out = []
        for r in range(max(1, h - 3), h + 4):
            # bỏ dòng thông tin chung (1 ô chữ gộp ngang cả bảng)
            n_chu = sum(1 for cc_ in range(1, min(ws.max_column, 30) + 1)
                        if isinstance(ws.cell(r, cc_).value, str) and ws.cell(r, cc_).value.strip())
            if r != h and n_chu < 2:
                continue
            rr, cc = m.get((r, c), (r, c))
            v = self.v(sh, cc, rr)
            if isinstance(v, str) and v.strip() and not re.fullmatch(r"[\[\(]?\d+[\]\)]?", v.strip()):
                t = re.sub(r"\s+", " ", v).strip()
                if t not in out:
                    out.append(t)
        return " / ".join(out)[:100]


# ============================================================== công thức
def bo_chuoi(f):
    return CHUOI.sub('""', f)


def tach_tham_chieu(f, sh_hien):
    """-> list dict(sh, c1, r1, c2, r2, ngoai, abs flags)."""
    out = []
    body = bo_chuoi(f)
    for m in QREF.finditer(body):
        sh = (m.group(1) or "").replace("''", "'") or m.group(2)
        ngoai = None
        mm = re.match(r"\[(\d+)\](.*)", sh)
        if mm:
            ngoai, sh = int(mm.group(1)), mm.group(2)
        out.append(dict(sh=sh, c1=column_index_from_string(m.group(4)), r1=int(m.group(6)),
                        c2=column_index_from_string(m.group(8) or m.group(4)), r2=int(m.group(10) or m.group(6)), ngoai=ngoai))
    body = QREF.sub(" ", body)
    for m in LREF.finditer(body):
        out.append(dict(sh=sh_hien, c1=column_index_from_string(m.group(2)), r1=int(m.group(4)),
                        c2=column_index_from_string(m.group(6) or m.group(2)), r2=int(m.group(8) or m.group(4)), ngoai=None))
    return out


def mau_cong_thuc(f, r, c):
    """Chuẩn hoá công thức về dạng tương đối (kiểu R1C1) để so 'cùng mẫu' giữa các dòng."""
    body = bo_chuoi(f)

    def tok(col_abs, col, row_abs, row):
        cc = column_index_from_string(col)
        rr = int(row)
        return (f"C{cc}" if col_abs else f"C[{cc - c}]") + (f"R{rr}" if row_abs else f"R[{rr - r}]")

    def q(m):
        s = tok(m.group(3), m.group(4), m.group(5), m.group(6))
        if m.group(8):
            s += ":" + tok(m.group(7), m.group(8), m.group(9), m.group(10))
        return f"<{m.group(1) or m.group(2)}>!" + s

    body = QREF.sub(q, body)

    def l(m):
        s = tok(m.group(1), m.group(2), m.group(3), m.group(4))
        if m.group(6):
            s += ":" + tok(m.group(5), m.group(6), m.group(7), m.group(8))
        return s

    return LREF.sub(l, body).replace(" ", "")


def la_tong(f):
    return bool(re.match(r"^=\s*(SUM|SUBTOTAL|SUMIF|SUMPRODUCT)\(", f, re.I))


def hang_so_nhung(f):
    """Hằng số đứng cạnh phép nhân / chia trong công thức (bỏ tham chiếu, chuỗi, đối số làm tròn)."""
    body = bo_chuoi(f)
    body = QREF.sub("X", body)
    body = LREF.sub("X", body)
    body = re.sub(r",\s*-?\d+\s*\)", ")", body)            # ROUND(x,0), INDEX(..,2)
    out = []
    for m in SO_NHUNG.finditer(body):
        a, b = body[:m.start()].rstrip()[-1:], body[m.end():].lstrip()[:1]
        if a in ("*", "/") or b in ("*", "/"):
            s = m.group(1)
            if s not in ("1", "0", "100", "1000"):
                out.append(s)
    return out


def la_so(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def la_cong_thuc(v):
    return isinstance(v, str) and v.startswith("=")


def giong(a, b):
    a, b = bo_dau(a), bo_dau(b)
    if not a or not b:
        return None
    if a == b or a in b or b in a:
        return 1.0
    return difflib.SequenceMatcher(None, a, b).ratio()


# ============================================================== phân loại 7 nhóm dữ liệu gốc
def nhom_goc(tieu_de, nhan="", sh=""):
    t = bo_dau(f"{tieu_de} {sh}")
    n = bo_dau(nhan)
    if "don gia" in t:
        return "2 Đơn giá"
    if re.search(r"k[iy] truoc|dot truoc|luy ke truoc", t + " " + n):
        return "3-4 Kỳ trước / lũy kế trước"
    if re.search(r"hop dong|\bhd\b", t) and re.search(r"khoi luong|\bkl\b", t):
        return "1 KL hợp đồng"
    if re.search(r"tam ung|da thanh toan|thanh toan dot|dot \d|unc", t + " " + n):
        return "7 Đã thanh toán / tạm ứng"
    if re.search(r"k[iy] nay|thuc hien|nghiem thu|hoan thanh", t):
        return "5 KL kỳ này"
    if re.search(r"%|ty le|thue suat|he so|giu lai", t + " " + n) or (re.search(r"\bthue\b|vat|gtgt", t + " " + n)
                                                                    and "gia tri" not in t):
        return "6 Tham số HĐ"
    if re.search(r"gia tri hop dong|gia tri hd", t + " " + n):
        return "1-2 Giá trị HĐ"
    if re.search(r"khoi luong|\bkl\b|dien giai|kich thuoc|dai|rong|cao|so luong", t):
        return "5 KL / diễn giải"
    return "Ngoài 7 nhóm - cần giải thích"


TAI_LIEU = {"1 KL hợp đồng": "A02 BOQ / A03 phụ lục", "2 Đơn giá": "A02 BOQ / A03 phụ lục",
            "1-2 Giá trị HĐ": "A01 hợp đồng / A03 phụ lục", "3-4 Kỳ trước / lũy kế trước": "E01 hồ sơ đợt trước đã duyệt",
            "5 KL kỳ này": "C01/C02 biên bản nghiệm thu", "5 KL / diễn giải": "B01 bản vẽ / B03 hoàn công / diễn giải KL",
            "6 Tham số HĐ": "A01 điều khoản hợp đồng", "7 Đã thanh toán / tạm ứng": "E01 hồ sơ đã duyệt, E02 UNC",
            "Ngoài 7 nhóm - cần giải thích": "Yêu cầu nhà thầu giải trình nguồn"}


# ============================================================== các kiểm tra
class KiemTra:
    def __init__(self, path, o_dich=None):
        self.sd = SoDo(path)
        self.o_dich = o_dich
        self.ph = []          # phát hiện
        self.kq = {}          # bảng kết quả theo kiểm tra

    def them(self, muc, kt, o, mo_ta, so_o=1, gia_tri=None, tai_lieu=""):
        self.ph.append(dict(muc=muc, kt=kt, o=o, mo_ta=mo_ta, so_o=so_o, gia_tri=gia_tri, tai_lieu=tai_lieu))

    # ---------------------------------------------------------- K1
    def k1_tong_quan(self):
        sd = self.sd
        dong, ngoai_dung = [], Counter()
        self._ct, self._ngoai_o = {}, {}         # (sh, r, c) -> công thức
        for ws in sd.wb.worksheets:
            n_ct = n_so = n_chua_tinh = 0
            loi_gt, loi_ct = [], 0
            wsv = sd.wv[ws.title]
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if la_cong_thuc(v):
                        n_ct += 1
                        self._ct[(ws.title, c.row, c.column)] = v
                        if "#REF!" in v:
                            loi_ct += 1
                        for m in re.finditer(r"\[(\d+)\]", bo_chuoi(v)):
                            ngoai_dung[int(m.group(1))] += 1
                            self._ngoai_o.setdefault(int(m.group(1)), set()).add((ws.title, c.column_letter))
                        gv = wsv.cell(c.row, c.column).value
                        if gv is None:
                            n_chua_tinh += 1
                        elif isinstance(gv, str) and gv.strip() in LOI:
                            loi_gt.append(c.coordinate)
                    elif la_so(v):
                        n_so += 1
            dong.append([ws.title, ws.sheet_state, ws.max_row, ws.max_column, n_ct, n_so, loi_ct, len(loi_gt), n_chua_tinh])
            if loi_ct:
                self.them("Trung bình" if ws.sheet_state == "visible" else "Kiểm soát", "K1", ws.title,
                          f"{loi_ct} công thức chứa #REF! (mất tham chiếu)", loi_ct, tai_lieu="Giải trình / sửa file")
            if loi_gt:
                self.them("Cao" if ws.sheet_state == "visible" else "Kiểm soát", "K1", f"{ws.title}!{loi_gt[0]}",
                          f"{len(loi_gt)} ô công thức đang ra lỗi (#REF!, #DIV/0!, #N/A...): " + ", ".join(loi_gt[:8]), len(loi_gt))
            if n_ct and n_chua_tinh > 0.5 * n_ct:
                self.them("Kiểm soát", "K1", ws.title, f"{n_chua_tinh}/{n_ct} công thức không có giá trị lưu: file chưa được Excel tính lại, "
                                                       "giá trị đọc được có thể không đúng", n_chua_tinh)
        self.kq["sheet"] = dong
        # liên kết file ngoài
        lk = []
        for i, tgt in sd.ngoai.items():
            lk.append([i, tgt, ngoai_dung.get(i, 0)])
            if ngoai_dung.get(i):
                cot = ", ".join(f"{a}!{b}" for a, b in sorted(self._ngoai_o.get(i, ())))
                self.them("Cao", "K1", f"[{i}]", f"{ngoai_dung[i]} công thức (cột {cot}) lấy số từ file ngoài: {tgt}", ngoai_dung[i],
                          tai_lieu="File Excel đang được liên kết (bản gốc)")
                self.ph[-1]["_ngoai"] = i
        rac = sum(1 for x in lk if not x[2])
        if rac:
            self.them("Kiểm soát", "K1", "Workbook", f"{rac} liên kết file ngoài chỉ còn trong tên định nghĩa (rác từ file mẫu khác), "
                                                     "không công thức nào dùng", rac)
        self.kq["ngoai"] = lk
        # tên định nghĩa lỗi
        ten_loi = [n for n, d in sd.wb.defined_names.items() if "#REF!" in str(d.attr_text)]
        self.kq["ten_loi"] = len(ten_loi)
        if ten_loi:
            self.them("Kiểm soát", "K1", "Workbook", f"{len(ten_loi)} tên định nghĩa (Name) bị #REF!", len(ten_loi))

    # ---------------------------------------------------------- K2
    def k2_ban_do_cong_thuc(self):
        sd = self.sd
        theo_cot = defaultdict(list)                # (sh, c) -> [(r, mẫu, f)]
        for (sh, r, c), f in self._ct.items():
            if sd.an(sh):
                continue
            theo_cot[(sh, c)].append((r, mau_cong_thuc(f, r, c), f))
        tom, lech, tay = [], [], []
        for (sh, c), ds in sorted(theo_cot.items()):
            ds_ = [x for x in ds if not la_tong(x[2])]
            if len(ds_) < 10:
                continue
            dem = Counter(m for _, m, _ in ds_)
            mau, n = dem.most_common(1)[0]
            ty = n / len(ds_)
            hdr = sd.tieu_de(sh, c)
            tom.append([sh, get_column_letter(c), hdr, len(ds), len(ds_), n, round(ty, 3), len(dem), ds_[0][2] if False else
                        next(f for _, m_, f in ds_ if m_ == mau)[:80]])
            if n < 10 or ty < 0.7:
                continue                                # cột không có mẫu chính rõ ràng
            dong_mau = sorted(r for r, m, _ in ds_ if m == mau)
            lo, hi = dong_mau[0], dong_mau[-1]
            for r, m, f in ds_:
                if m != mau and dem[m] <= 2 and lo <= r <= hi:
                    lech.append([sh, f"{get_column_letter(c)}{r}", hdr, sd.nhan_dong(sh, r), f[:100],
                                 next(ff for rr, mm, ff in ds_ if mm == mau)[:100], sd.v(sh, c, r)])
            # số gõ tay nằm giữa cột công thức
            co_ct = {r for r, _, _ in ds}
            for r in range(lo, hi + 1):
                if r in co_ct:
                    continue
                v = sd.f(sh, c, r)
                if la_so(v) and v != 0:
                    du_lieu = any(la_cong_thuc(sd.f(sh, cc, r)) for cc in range(1, sd.wb[sh].max_column + 1) if cc != c) \
                              or sd.nhan_dong(sh, r)
                    if du_lieu:
                        tay.append([sh, f"{get_column_letter(c)}{r}", hdr, sd.nhan_dong(sh, r), v,
                                    next(ff for rr, mm, ff in ds_ if mm == mau)[:80], nhom_goc(hdr, sd.nhan_dong(sh, r), sh)])
        self.kq.update(cot=tom, lech=lech, tay=tay)
        by = Counter(x[0] for x in lech)
        for sh, n in by.items():
            vd = [x[1] for x in lech if x[0] == sh][:6]
            self.them("Trung bình", "K2", f"{sh}!{vd[0]}", f"{n} ô công thức lệch mẫu chung của cột ({', '.join(vd)}...)", n,
                      tai_lieu="Giải trình vì sao công thức khác")
        by = defaultdict(list)
        for x in tay:
            by[(x[0], x[1].rstrip("0123456789"))].append(x)
        for (sh, col), xs in by.items():
            self.them("Cao", "K2", f"{sh}!{xs[0][1]}",
                      f"{len(xs)} số gõ tay nằm giữa cột công thức '{xs[0][2][:40]}' ({', '.join(x[1] for x in xs[:6])}...)",
                      len(xs), sum(x[4] for x in xs), TAI_LIEU.get(xs[0][6], ""))

    # ---------------------------------------------------------- K3
    def k3_lien_ket_dong(self):
        sd = self.sd
        out = []
        cap = Counter()
        for (sh, r, c), f in self._ct.items():
            for x in tach_tham_chieu(f, sh):
                if x["sh"] != sh and x["ngoai"] is None and x["r1"] == x["r2"]:
                    cap[(sh, x["sh"])] += 1
        for (sh, r, c), f in self._ct.items():
            if sd.an(sh):
                continue
            refs = [x for x in tach_tham_chieu(f, sh) if x["sh"] != sh and x["ngoai"] is None
                    and x["c1"] == x["c2"] and x["r1"] == x["r2"] and x["sh"] in sd.wb.sheetnames]
            if not refs or len({(x["sh"], x["r1"]) for x in refs}) != 1:
                continue
            x = refs[0]
            if cap[(sh, x["sh"])] < 20:
                continue                                # bảng tổng hợp, giấy đề nghị: không phải liên kết theo từng hạng mục
            if la_cong_thuc(sd.f(x["sh"], x["c1"], x["r1"])) and la_tong(sd.f(x["sh"], x["c1"], x["r1"])):
                continue                                # trỏ vào dòng tổng: liên kết tổng hợp, không so tên
            a, b = sd.nhan_dong(sh, r), sd.nhan_dong(x["sh"], x["r1"])
            if not a or len(a) < 6 or re.match(r"(tong|cong|thue|vat|gia tri)", bo_dau(b)):
                continue
            g = giong(a, b)
            if g is None:
                kl = "Trỏ tới dòng không có tên hạng mục"
            elif g < 0.6:
                kl = "Lệch dòng: tên hai đầu khác nhau"
            else:
                continue
            out.append([sh, f"{get_column_letter(c)}{r}", sd.tieu_de(sh, c), a, f[:80],
                        f"{x['sh']}!{get_column_letter(x['c1'])}{x['r1']}", b, round(g or 0, 2), kl, sd.v(sh, c, r)])
        self.kq["lien_ket"] = out
        by = defaultdict(list)
        for x in out:
            by[(x[0], x[1].rstrip("0123456789"), x[8])].append(x)
        for (sh, col, kl), xs in by.items():
            self.them("Cao", "K3", f"{sh}!{xs[0][1]}", f"{len(xs)} ô: {kl} ({', '.join(x[1] for x in xs[:6])} -> {xs[0][5]}...)",
                      len(xs), sum(x[9] for x in xs if la_so(x[9])), "Sửa liên kết; xác định nguồn đúng của hạng mục")

    # ---------------------------------------------------------- K4
    def k4_sheet_an(self):
        sd = self.sd
        an = [ws.title for ws in sd.wb.worksheets if ws.sheet_state != "visible"]
        tro = Counter()
        for (sh, r, c), f in self._ct.items():
            if sd.an(sh):
                continue
            for x in tach_tham_chieu(f, sh):
                if x["sh"] in an:
                    tro[(sh, x["sh"])] += 1
        dong, so_sanh = [], []
        for h in an:
            ws = sd.wb[h]
            ten_goc = re.sub(r"\s*\(\d+\)\s*$", "", h).strip()
            twin = next((w.title for w in sd.wb.worksheets if w.sheet_state == "visible" and w.title != h and
                         bo_dau(w.title) == bo_dau(ten_goc)), None)
            n_tro = sum(n for (a, b), n in tro.items() if b == h)
            khac = []
            if twin:
                a, b = sd.wv[twin], sd.wv[h]
                for r in range(1, max(a.max_row, b.max_row) + 1):
                    for c in range(1, max(a.max_column, b.max_column) + 1):
                        va, vb = a.cell(r, c).value, b.cell(r, c).value
                        if va != vb and not (la_so(va) and la_so(vb) and abs(va - vb) < 1e-9):
                            khac.append((r, c, va, vb))
                for r, c, va, vb in khac[:300]:
                    so_sanh.append([twin, h, f"{get_column_letter(c)}{r}", sd.nhan_dong(twin, r), va, vb])
            dong.append([h, ws.sheet_state, ws.max_row, sum(1 for k in self._ct if k[0] == h), n_tro,
                         ", ".join(f"{a} ({n})" for (a, b), n in tro.items() if b == h), twin or "", len(khac) if twin else None])
            mo_ta = f"Sheet ẩn '{h}'"
            if n_tro:
                mo_ta += f" được {n_tro} công thức ở sheet hiện tham chiếu"
            if twin:
                mo_ta += f"; là bản khác của '{twin}', khác {len(khac)} ô"
            self.them("Trung bình" if (n_tro or (twin and khac)) else "Kiểm soát", "K4", h, mo_ta, n_tro or 1,
                      tai_lieu="Giải trình sheet ẩn; xác nhận bản nào là bản đã ký")
        self.kq.update(an=dong, an_khac=so_sanh)

    # ---------------------------------------------------------- K5
    def k5_so_dang_chu(self):
        sd = self.sd
        out = []
        for ws in sd.wb.worksheets:
            if ws.sheet_state != "visible":
                continue
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if isinstance(v, str) and not v.startswith("="):
                        t = v.strip().replace(" ", "")
                        if t and re.fullmatch(r"[-+]?\d+([.,]\d+)?%?", t) and not re.fullmatch(r"\d{1,2}", t):
                            out.append([ws.title, c.coordinate, sd.tieu_de(ws.title, c.column), sd.nhan_dong(ws.title, c.row), v])
        # bỏ cột STT, số hiệu (nhiều ô chữ số liên tiếp ở cột trái tên)
        out = [x for x in out if column_index_from_string(re.match(r"[A-Z]+", x[1]).group(0)) > sd.cot_ten(x[0])]
        self.kq["chu"] = out
        if out:
            self.them("Trung bình", "K5", f"{out[0][0]}!{out[0][1]}", f"{len(out)} ô số đang lưu dạng chữ, Excel không cộng vào tổng "
                                                                     f"({', '.join(x[0] + '!' + x[1] for x in out[:5])}...)", len(out))

    # ---------------------------------------------------------- K6
    def k6_tham_so_nhung(self):
        sd = self.sd
        cot_cay = set()
        for row in self.kq.get("cay", []):
            m = re.match(r"(.+)!([A-Z]+)\d+", str(row[2] or ""))
            if m:
                cot_cay.add((m.group(1), m.group(2)))
        dem = defaultdict(list)
        for (sh, r, c), f in self._ct.items():
            if sd.an(sh):
                continue
            for s in hang_so_nhung(f):
                dem[s].append(f"{sh}!{get_column_letter(c)}{r}")
        out = []
        for s, cells in sorted(dem.items(), key=lambda x: -len(x[1])):
            sh0 = cells[0].split("!")[0]
            r0 = int(re.search(r"(\d+)$", cells[0]).group(1))
            c0 = column_index_from_string(re.search(r"!([A-Z]+)", cells[0]).group(1))
            loai = "Thuế suất?" if s in ("1.08", "1.1", "1.05", "1.1", "0.08", "8%", "10%") else \
                   ("Tỷ lệ %" if s.endswith("%") or re.fullmatch(r"0\.\d{1,2}", s) else "Hằng số khác (kích thước, hệ số kỹ thuật...)")
            trong_cay = [x for x in cells if (x.split("!")[0], re.search(r"!([A-Z]+)", x).group(1)) in cot_cay]
            out.append([s, loai, len(cells), len(trong_cay), ", ".join((trong_cay or cells)[:8]),
                        self._ct.get((sh0, r0, c0), "")[:100]])
        self.kq["nhung"] = out
        for s, loai, n, n_cay, vd, f in out:
            if not loai.startswith("Hằng số khác") and n_cay:
                n = n_cay
                self.them("Trung bình", "K6", vd.split(",")[0], f"Hằng số {s} ({loai}) gõ thẳng trong {n} công thức: cần khớp điều khoản HĐ "
                                                               "(Lớp 2) và nên đặt thành ô tham số", n, tai_lieu="A01 điều khoản hợp đồng")

    # ---------------------------------------------------------- K7 cây truy vết
    def k7_cay_truy_vet(self, max_sau=14):
        sd = self.sd
        dich = self.o_dich
        if not dich:
            t = tim_so_tong_hop(sd.wv)
            dich = (t.get("tt_ky_nay") or {}).get("o")
        rows = []
        if not dich or "!" not in dich:
            self.kq["cay"], self.kq["dich"] = rows, None
            self.them("Kiểm soát", "K7", "-", "Không tự tìm được ô 'giá trị đề nghị thanh toán kỳ này'; chạy lại với --o 'Sheet!Ô'")
            return
        sh0, cd0 = dich.rsplit("!", 1)
        sh0 = sh0.strip("'")
        m = re.match(r"([A-Z]+)(\d+)", cd0.replace("$", ""))
        da = {}

        def them_dong(cap, ten, o, f, v, loai, nhom="", ghi=""):
            rows.append([cap, ("    " * cap) + ten, o, f, v, loai, nhom, TAI_LIEU.get(nhom, "") if loai in ("Nhập tay", "Nhóm nhập tay") else "", ghi])

        def tham(sh, c, r, cap):
            key = (sh, c, r)
            o = f"{sh}!{get_column_letter(c)}{r}"
            ten = (sd.nhan_dong(sh, r) or "") + (f"  [{sd.tieu_de(sh, c)}]" if sd.tieu_de(sh, c) else "")
            if key in da:
                them_dong(cap, ten, o, "", sd.v(sh, c, r), "Đã phân rã", ghi=f"xem dòng {da[key]}")
                return
            f = sd.f(sh, c, r)
            if not la_cong_thuc(f):
                loai = "Nhập tay" if la_so(f) else ("Trống" if f in (None, "") else "Chữ")
                da[key] = len(rows) + 2
                them_dong(cap, ten, o, "", f, loai, nhom_goc(sd.tieu_de(sh, c), sd.nhan_dong(sh, r), sh) if loai == "Nhập tay" else "")
                return
            da[key] = len(rows) + 2
            them_dong(cap, ten, o, f[:150], sd.v(sh, c, r), "Công thức", ghi="sheet ẩn" if sd.an(sh) else "")
            if cap >= max_sau:
                them_dong(cap + 1, "… dừng vì quá sâu", "", "", None, "")
                return
            for s in hang_so_nhung(f):
                them_dong(cap + 1, f"Hằng số trong công thức: {s}", o, "", s, "Tham số gõ cứng", "6 Tham số HĐ")
            for x in tach_tham_chieu(f, sh):
                if x["ngoai"] is not None:
                    them_dong(cap + 1, f"File ngoài: {sd.ngoai.get(x['ngoai'], '?')[:80]}",
                              f"[{x['ngoai']}]{x['sh']}!{get_column_letter(x['c1'])}{x['r1']}", "", None, "File ngoài",
                              ghi="File Excel đang được liên kết")
                    continue
                if x["sh"] not in sd.wb.sheetnames:
                    them_dong(cap + 1, f"Sheet không tồn tại: {x['sh']}", "", "", None, "Lỗi tham chiếu")
                    continue
                n_o = (x["c2"] - x["c1"] + 1) * (x["r2"] - x["r1"] + 1)
                if n_o == 1:
                    tham(x["sh"], x["c1"], x["r1"], cap + 1)
                    continue
                # dải: gom theo mẫu công thức, phân rã 1 ô đại diện cho mỗi mẫu; ô nhập tay gom thành 1 nút
                mau, tay, tong_tay = defaultdict(list), [], 0.0
                for rr in range(x["r1"], x["r2"] + 1):
                    for cc in range(x["c1"], x["c2"] + 1):
                        ff = sd.f(x["sh"], cc, rr)
                        if la_cong_thuc(ff):
                            mau[mau_cong_thuc(ff, rr, cc)].append((cc, rr))
                        elif la_so(ff):
                            tay.append((cc, rr))
                            tong_tay += ff
                rg = f"{x['sh']}!{get_column_letter(x['c1'])}{x['r1']}:{get_column_letter(x['c2'])}{x['r2']}"
                them_dong(cap + 1, f"Dải {n_o} ô: {sum(len(v) for v in mau.values())} công thức ({len(mau)} mẫu), {len(tay)} ô nhập tay",
                          rg, "", None, "Dải")
                if tay:
                    hd = sd.tieu_de(x["sh"], tay[0][0])
                    them_dong(cap + 2, f"{len(tay)} ô nhập tay trong dải  [{hd}]", rg, "", tong_tay, "Nhóm nhập tay",
                              nhom_goc(hd, "", x["sh"]))
                for mm, cells in sorted(mau.items(), key=lambda kv: -len(kv[1]))[:12]:
                    # đại diện: ô có giá trị lớn nhất (tránh dòng tiêu đề nhóm bằng 0)
                    cc, rr = max(cells, key=lambda o: abs(sd.v(x["sh"], o[0], o[1])) if la_so(sd.v(x["sh"], o[0], o[1])) else -1)
                    them_dong(cap + 2, f"Mẫu công thức áp dụng cho {len(cells)} ô, đại diện:", "", "", None, "Mẫu")
                    tham(x["sh"], cc, rr, cap + 3)

        tham(sh0, column_index_from_string(m.group(1)), int(m.group(2)), 0)
        self.kq["cay"], self.kq["dich"] = rows, dich
        la = [r for r in rows if r[5] in ("Nhập tay", "Nhóm nhập tay")]
        ngoai7 = [r for r in la if r[6].startswith("Ngoài")]
        if ngoai7:
            self.them("Cao", "K7", ngoai7[0][2], f"Số đề nghị thanh toán phụ thuộc {len(ngoai7)} nút nhập tay KHÔNG thuộc 7 nhóm dữ liệu gốc "
                                                f"({', '.join(r[2] for r in ngoai7[:5])}...)", len(ngoai7),
                      tai_lieu="Yêu cầu nhà thầu giải trình nguồn")
        fn = [r for r in rows if r[5] == "File ngoài"]
        if fn:
            self.them("Cao", "K7", fn[0][2], f"Số đề nghị thanh toán phụ thuộc {len(fn)} tham chiếu file ngoài", len(fn),
                      tai_lieu="File Excel đang được liên kết")

    # ---------------------------------------------------------- K8 nguồn KL từng hạng mục
    def k8_nguon_kl(self):
        sd = self.sd
        try:
            bang = tim_bang_chi_tiet(sd.wv)
        except Exception as e:
            self.kq["nguon"] = []
            self.them("Kiểm soát", "K8", "-", f"Không nhận được bảng chi tiết KL x đơn giá: {e}")
            return
        sh, cols = bang["sheet"], bang["cols"]
        cot_kl = cols.get("kl_lk") or cols.get("kl_kn")
        out = []
        for it in bang["items"]:
            if not it["dvt"]:
                continue
            nguon, duong, co = self._nguon(sh, cot_kl, it["dong"], it.get("kl_hd"))
            if cols.get("kl_kt") and (it.get("kl_kt") or 0):
                n_kt, d_kt, co_kt = self._nguon(sh, cols["kl_kt"], it["dong"], it.get("kl_hd"))
                n_kt = n_kt + (" (qua sheet ẩn)" if "qua sheet ẩn" in co_kt else "")
            else:
                n_kt = "0 / không có"
            dg = it.get("don_gia") or 0
            gt_kn = it.get("gt_kn") if it.get("gt_kn") is not None else (it.get("kl_kn") or 0) * dg
            gt_lk = it.get("gt_lk") if it.get("gt_lk") is not None else (it.get("kl_lk") or 0) * dg
            out.append([it["dong"], it["ten"][:80], it["dvt"], it.get("kl_hd"), it.get("kl_lk"), round(gt_kn or 0), round(gt_lk or 0),
                        nguon, "; ".join(co), " -> ".join(duong), it.get("kl_kt"), n_kt])
        self.kq.update(nguon=out, bang_sh=sh, cot_kl=get_column_letter(cot_kl) if cot_kl else "")
        tang = defaultdict(lambda: [0, 0.0, 0.0])
        for x in out:
            t = tang[x[7]]
            t[0] += 1
            t[1] += x[5]
            t[2] += x[6]
        self.kq["tang"] = sorted(([k] + v for k, v in tang.items()), key=lambda r: -r[3])
        chan = [x for x in out if "đang bị chặn" in x[8]]
        if chan:
            self.them("Trung bình", "K8", f"{sh}!{self.kq['cot_kl']}{chan[0][0]}",
                      f"{len(chan)} hạng mục có KL diễn giải vượt KL HĐ, đang bị công thức chặn về KL HĐ: phát sinh chưa xử lý, "
                      "BOQ sai hay diễn giải thổi phồng?", len(chan), tai_lieu="A04 hồ sơ phát sinh / B04 thay đổi thiết kế")
        kt_dem = Counter(x[11] for x in out if not x[11].startswith("0 /"))
        self.kq["tang_kt"] = [[k, n] for k, n in kt_dem.most_common()]
        for k, n in kt_dem.items():
            if k.startswith(("Nhập tay", "Tỷ lệ")) or "sheet ẩn" in k:
                self.them("Cao" if "Nhập tay" in k or "sheet ẩn" in k else "Trung bình", "K8", f"{sh}!{get_column_letter(cols['kl_kt'])}",
                          f"{n} hạng mục có KL kỳ trước nguồn '{k}'", n, tai_lieu="E01 hồ sơ đợt trước đã duyệt (đối chiếu từng hạng mục)")
        for k, n, gkn, glk in self.kq["tang"]:
            if k in ("Nhập tay", "Bằng KL HĐ", "Nguồn trống", "Link lệch dòng", "File ngoài") and n:
                self.them("Cao", "K8", f"{sh}!{self.kq['cot_kl']}", f"{n} hạng mục có KL lũy kế nguồn '{k}': GT lũy kế {glk:,.0f}, kỳ này {gkn:,.0f}",
                          n, glk, {"Nhập tay": "Diễn giải KL, bản vẽ / hoàn công", "Bằng KL HĐ": "Hoàn công + BB nghiệm thu (không chấp nhận 'bằng KL HĐ')",
                                   "Nguồn trống": "Sửa liên kết, xác định nguồn", "Link lệch dòng": "Sửa liên kết, xác định nguồn",
                                   "File ngoài": "File Excel đang liên kết"}[k])
            elif k == "Tỷ lệ × KL HĐ" and n:
                self.them("Trung bình", "K8", f"{sh}!{self.kq['cot_kl']}", f"{n} hạng mục KL lũy kế = tỷ lệ % × KL HĐ (ước tính tiến độ): GT lũy kế {glk:,.0f}",
                          n, glk, "Xác nhận % tiến độ của tư vấn giám sát")

    def _lech(self, sh, r, x):
        a, b = self.sd.nhan_dong(sh, r), self.sd.nhan_dong(x["sh"], x["r1"])
        g = giong(a, b)
        return bool(a) and len(a) >= 6 and (g is None or g < 0.6)

    def _nguon(self, sh, c, r, kl_hd, sau=0, duong=None, co=None):
        """Lần theo ô KL qua các tham chiếu 1 ô cho tới khi gặp công thức tính, số gõ tay, hoặc ô trống."""
        sd = self.sd
        duong = (duong or []) + [f"{sh}!{get_column_letter(c)}{r}"]
        co = co or []
        if sd.an(sh) and "qua sheet ẩn" not in co:
            co.append("qua sheet ẩn")
        f = sd.f(sh, c, r)
        v = sd.v(sh, c, r)
        bang_hd = lambda x: la_so(x) and la_so(kl_hd) and kl_hd and abs(x - kl_hd) <= 1e-6 * max(1, abs(kl_hd))
        if f in (None, ""):
            return "Nguồn trống", duong, co
        if la_so(f):
            if bang_hd(f) and sau > 0:
                co.append("số gõ tay đúng bằng KL HĐ")
            return "Nhập tay", duong, co
        if not la_cong_thuc(f) or sau > 8:
            return "Khác", duong, co
        refs = tach_tham_chieu(f, sh)
        if any(x["ngoai"] is not None for x in refs):
            return "File ngoài", duong, co
        don = [x for x in refs if x["c1"] == x["c2"] and x["r1"] == x["r2"]]
        body = bo_chuoi(f)
        chi_1_ref = len(refs) == 1 and len(don) == 1 and re.fullmatch(r"=\s*\+?\s*(?:'[^']+'|[^!=]+)?!?\$?[A-Z]{1,3}\$?\d+\s*", body)
        if chi_1_ref:
            x = don[0]
            if bang_hd(v) and x["sh"] in sd.wb.sheetnames and re.search(r"hop dong|\bhd\b", bo_dau(sd.tieu_de(x["sh"], x["c1"]))):
                return "Bằng KL HĐ", duong + [f"{x['sh']}!{get_column_letter(x['c1'])}{x['r1']}"], co
            if x["sh"] != sh and self._lech(sh, r, x):
                return "Link lệch dòng", duong + [f"{x['sh']}!{get_column_letter(x['c1'])}{x['r1']}"], co
            if x["sh"] in sd.wb.sheetnames:
                return self._nguon(x["sh"], x["c1"], x["r1"], kl_hd, sau + 1, duong, co)
            return "Nguồn trống", duong, co
        # tham chiếu 1 ô nhân hằng số: tỷ lệ x KL HĐ ?
        if len(don) == 1 and len(refs) == 1 and hang_so_nhung(f):
            if la_so(v) and la_so(kl_hd) and kl_hd and 0 < v / kl_hd < 1:
                return "Tỷ lệ × KL HĐ", duong, co
            return "Hệ số × ô khác", duong, co
        cong_tru = re.fullmatch(r"=\s*[+\-]?\s*X(\s*[+\-]\s*X)+\s*", LREF.sub("X", QREF.sub("X", body)))
        if cong_tru and len(don) == len(refs) >= 2:
            # lũy kế = kỳ trước + kỳ này, kỳ này = lũy kế - kỳ trước ...: đi theo thành phần 'lũy kế', rồi 'kỳ này'
            def uu_tien(x):
                t = bo_dau(sd.tieu_de(x["sh"], x["c1"]))
                return (0 if "luy ke" in t and "truoc" not in t else 1 if re.search(r"k[iy] nay", t) else 2)
            x = sorted(don, key=uu_tien)[0]
            co.append("cộng/trừ nhiều ô")
            if x["sh"] != sh and self._lech(sh, r, x):
                return "Link lệch dòng", duong + [f"{x['sh']}!{get_column_letter(x['c1'])}{x['r1']}"], co
            return self._nguon(x["sh"], x["c1"], x["r1"], kl_hd, sau + 1, duong, co)
        if re.match(r"^=\s*(IF|MIN|MAX)\(", f, re.I):
            co.append("có IF/MIN/MAX")
            # đi theo nhánh thực sự cho ra giá trị (so giá trị đã lưu); ưu tiên tham chiếu sang sheet khác (nguồn diễn giải)
            ung = [x for x in don if x["sh"] in sd.wb.sheetnames]
            gia = lambda x: sd.v(x["sh"], x["c1"], x["r1"])
            khop = [x for x in ung if la_so(gia(x)) and la_so(v) and abs(gia(x) - v) <= 1e-9 * max(1, abs(v))]
            khop.sort(key=lambda x: x["sh"] == sh)
            dg = next((x for x in ung if x["sh"] != sh), None)
            if dg and self._lech(sh, r, dg):
                return "Link lệch dòng", duong + [f"{dg['sh']}!{get_column_letter(dg['c1'])}{dg['r1']}"], co
            if dg and la_so(gia(dg)) and la_so(kl_hd) and gia(dg) > kl_hd + 1e-9 and bang_hd(v):
                co.append(f"diễn giải {gia(dg):g} > KL HĐ {kl_hd:g}, đang bị chặn về KL HĐ")
                x = dg
            else:
                x = (khop or ([dg] if dg else []) or ung or [None])[0]
            if x:
                if x["sh"] != sh and self._lech(sh, r, x):
                    return "Link lệch dòng", duong + [f"{x['sh']}!{get_column_letter(x['c1'])}{x['r1']}"], co
                return self._nguon(x["sh"], x["c1"], x["r1"], kl_hd, sau + 1, duong, co)
        if bang_hd(v) and len(refs) == 1:
            return "Bằng KL HĐ", duong, co
        return "Tính từ diễn giải / công thức", duong, co

    # ---------------------------------------------------------- chạy
    def chay(self):
        for k in (self.k1_tong_quan, self.k2_ban_do_cong_thuc, self.k3_lien_ket_dong, self.k4_sheet_an,
                  self.k5_so_dang_chu, self.k7_cay_truy_vet, self.k6_tham_so_nhung, self.k8_nguon_kl):
            try:
                k()
            except Exception as e:  # một kiểm tra lỗi không làm dừng các kiểm tra khác
                self.them("Kiểm soát", k.__name__[:2].upper(), "-", f"Kiểm tra lỗi: {type(e).__name__}: {e}")
        self._xep_lai()
        self.ph.sort(key=lambda x: (MUC[x["muc"]], x["kt"]))
        return self

    def _xep_lai(self):
        """File ngoài chỉ dùng ở cột phụ (không nằm trên cây truy vết số đề nghị) -> hạ mức."""
        cot_cay = set()
        for row in self.kq.get("cay", []):
            m = re.match(r"(.+)!([A-Z]+)\d+", str(row[2] or ""))
            if m:
                cot_cay.add((m.group(1), m.group(2)))
        if not cot_cay:
            return
        for x in self.ph:
            if x.get("_ngoai") and not (self._ngoai_o.get(x["_ngoai"], set()) & cot_cay):
                x["muc"] = "Trung bình"
                x["mo_ta"] += " | không nằm trên đường tính ra số đề nghị thanh toán (cột phụ / kiểm tra)"


# ============================================================== ghi kết quả
HF = PatternFill("solid", fgColor="1F4A7A")
HFONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BASE = Font(name="Arial", size=10)
NOTE = Font(name="Arial", size=9, italic=True, color="6B6B6B")
TO = {"Cao": PatternFill("solid", fgColor="F8D7D3"), "Trung bình": PatternFill("solid", fgColor="FFF2CC"),
      "Kiểm soát": PatternFill("solid", fgColor="E3EBF5")}


def _sheet(wb, ten, tieu_de, cot, rong, dong, ghi=(), fmt=None, to_cot=None):
    ws = wb.create_sheet(ten)
    ws.cell(1, 1, tieu_de).font = Font(name="Arial", size=12, bold=True)
    r = 2
    for g in ghi:
        ws.cell(r, 1, g).font = NOTE
        r += 1
    r += 1
    for j, h in enumerate(cot, 1):
        c = ws.cell(r, j, h)
        c.font, c.fill = HFONT, HF
        c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(j)].width = rong[j - 1]
    ws.freeze_panes = ws.cell(r + 1, 1)
    for d in dong:
        r += 1
        for j, v in enumerate(d, 1):
            c = ws.cell(r, j, v)
            if isinstance(v, str) and v.startswith("="):
                c.data_type = "s"
            c.font = BASE
            if fmt and j in fmt:
                c.number_format = fmt[j]
        if to_cot is not None and d[to_cot] in TO:
            for j in range(1, len(d) + 1):
                ws.cell(r, j).fill = TO[d[to_cot]]
    if not dong:
        ws.cell(r + 1, 1, "(không có)").font = NOTE
    return ws


def ghi(kt, out):
    k, sd = kt.kq, kt.sd
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    dem = Counter(x["muc"] for x in kt.ph)
    _sheet(wb, "Phat hien", f"LỚP 1 - TÍNH TOÀN VẸN FILE: {sd.path.name}",
           ["#", "Mức", "Kiểm tra", "Vị trí", "Mô tả", "Số ô", "Giá trị liên quan", "Tài liệu / việc cần làm"],
           [4, 11, 8, 30, 90, 8, 18, 40],
           [[i + 1, x["muc"], x["kt"], x["o"], x["mo_ta"], x["so_o"], x["gia_tri"], x["tai_lieu"]] for i, x in enumerate(kt.ph)],
           ghi=[f"Cao {dem['Cao']} · Trung bình {dem['Trung bình']} · Kiểm soát {dem['Kiểm soát']}.  "
                "Cao = có thể làm sai số tiền; Trung bình = sai logic / thiếu căn cứ; Kiểm soát = lỗi cấu trúc file.",
                "Chỉ đọc file, không tính lại bảng tổng hợp. Lớp 2 (đối chiếu hợp đồng, BOQ, phụ lục) và Lớp 3 (chứng từ) làm ở module khác."],
           fmt={7: "#,##0"}, to_cot=1)
    _sheet(wb, "K1 Tong quan", "K1 - Tổng quan workbook",
           ["Sheet", "Trạng thái", "Số dòng", "Số cột", "Công thức", "Số gõ tay", "Công thức #REF!", "Ô đang lỗi", "Công thức chưa có giá trị lưu"],
           [30, 11, 9, 8, 11, 11, 12, 10, 14], k.get("sheet", []),
           ghi=[f"Liên kết file ngoài: {len(k.get('ngoai', []))} (được công thức dùng: {sum(1 for x in k.get('ngoai', []) if x[2])}); "
                f"tên định nghĩa lỗi #REF!: {k.get('ten_loi', 0)}"])
    _sheet(wb, "K1 File ngoai", "K1 - Liên kết file ngoài", ["#", "File", "Số công thức đang dùng"], [5, 120, 14], k.get("ngoai", []))
    _sheet(wb, "K2 Mau cot", "K2 - Mẫu công thức chính của từng cột",
           ["Sheet", "Cột", "Tiêu đề", "Công thức", "Không tính tổng", "Ô theo mẫu chính", "Tỷ lệ", "Số mẫu khác nhau", "Mẫu chính (ví dụ)"],
           [26, 6, 36, 10, 10, 10, 8, 10, 60], k.get("cot", []), fmt={7: "0%"},
           ghi=["Cột có mẫu chính (≥ 10 ô, ≥ 70%) mới xét lệch mẫu. Dòng tổng (SUM/SUBTOTAL) không xét."])
    _sheet(wb, "K2 Lech mau", "K2 - Ô công thức lệch mẫu chính của cột",
           ["Sheet", "Ô", "Tiêu đề cột", "Hạng mục", "Công thức ô này", "Công thức mẫu chính", "Giá trị"],
           [24, 8, 30, 44, 50, 50, 14], k.get("lech", []), fmt={7: "#,##0.###"})
    _sheet(wb, "K2 Go tay", "K2 - Số gõ tay nằm giữa cột công thức",
           ["Sheet", "Ô", "Tiêu đề cột", "Hạng mục", "Giá trị gõ tay", "Công thức mẫu của cột", "Nhóm dữ liệu gốc"],
           [24, 8, 30, 44, 14, 50, 26], k.get("tay", []), fmt={5: "#,##0.###"})
    _sheet(wb, "K3 Lien ket dong", "K3 - Liên kết sang sheet khác trỏ sai hạng mục",
           ["Sheet", "Ô", "Tiêu đề cột", "Hạng mục (dòng này)", "Công thức", "Trỏ tới", "Hạng mục (dòng đích)", "Độ giống", "Kết luận", "Giá trị"],
           [22, 8, 26, 40, 30, 22, 40, 8, 28, 14], k.get("lien_ket", []), fmt={10: "#,##0.###"},
           ghi=["So tên hạng mục ở dòng có công thức với tên ở dòng được trỏ tới (bỏ dấu). Độ giống < 0,6 hoặc dòng đích không có tên thì liệt kê."])
    _sheet(wb, "K4 Sheet an", "K4 - Sheet ẩn",
           ["Sheet ẩn", "Trạng thái", "Số dòng", "Công thức", "Bị sheet hiện tham chiếu (ô)", "Từ sheet", "Bản hiện cùng tên", "Số ô khác bản hiện"],
           [26, 11, 9, 10, 14, 40, 22, 12], k.get("an", []))
    _sheet(wb, "K4 Khac ban hien", "K4 - Khác biệt giữa sheet ẩn và sheet hiện cùng tên (tối đa 300 ô / cặp)",
           ["Sheet hiện", "Sheet ẩn", "Ô", "Hạng mục", "Giá trị bản hiện", "Giá trị bản ẩn"], [22, 22, 8, 50, 18, 18], k.get("an_khac", []))
    _sheet(wb, "K5 So dang chu", "K5 - Số lưu dạng chữ", ["Sheet", "Ô", "Tiêu đề cột", "Hạng mục", "Nội dung"], [24, 8, 30, 50, 14],
           k.get("chu", []))
    _sheet(wb, "K6 Hang so", "K6 - Hằng số gõ cứng trong công thức (nhân / chia)",
           ["Hằng số", "Loại", "Số công thức", "Trong cột thuộc cây truy vết", "Ví dụ ô", "Công thức ví dụ"], [10, 22, 12, 14, 70, 60],
           k.get("nhung", []),
           ghi=["Hằng số kiểu thuế suất, tỷ lệ nằm trên đường tính ra số đề nghị thanh toán phải khớp điều khoản hợp đồng (Lớp 2) "
                "và nên nằm ở ô tham số riêng. Hằng số kích thước trong diễn giải KL chỉ liệt kê."])
    _sheet(wb, "K7 Cay truy vet", f"K7 - Cây truy vết từ {k.get('dich') or '(không có ô đích)'}",
           ["Cấp", "Thành phần", "Ô", "Công thức", "Giá trị", "Loại nút", "Nhóm dữ liệu gốc", "Tài liệu chứng minh", "Ghi chú"],
           [5, 70, 26, 50, 16, 14, 26, 34, 18], k.get("cay", []), fmt={5: "#,##0.###"},
           ghi=["Số đề nghị thanh toán = dữ liệu nào x dữ liệu nào. Dải nhiều ô: phân rã một ô đại diện cho mỗi mẫu công thức.",
                "Mọi nút nhập tay phải thuộc 7 nhóm dữ liệu gốc; 'Ngoài 7 nhóm' là cờ cần nhà thầu giải trình."])
    ws = _sheet(wb, "K8 Nguon KL", f"K8 - Nguồn KL lũy kế từng hạng mục (bảng {k.get('bang_sh', '?')}, cột {k.get('cot_kl', '?')})",
                ["Nhóm nguồn", "Số hạng mục", "GT kỳ này", "GT lũy kế"], [34, 12, 18, 18], k.get("tang", []), fmt={3: "#,##0", 4: "#,##0"},
                ghi=["Dùng cho chọn mẫu (bước 5): nhóm 'Nhập tay', 'Bằng KL HĐ', 'Nguồn trống', 'Link lệch dòng' rủi ro cao."])
    _sheet(wb, "K8 Tung dong", "K8 - Nguồn KL từng hạng mục",
           ["Dòng", "Hạng mục", "ĐVT", "KL HĐ", "KL lũy kế", "GT kỳ này", "GT lũy kế", "Nguồn KL lũy kế", "Lưu ý", "Đường đi",
            "KL kỳ trước", "Nguồn KL kỳ trước"],
           [7, 50, 8, 12, 12, 16, 16, 26, 26, 70, 12, 26], k.get("nguon", []),
           fmt={4: "#,##0.###", 5: "#,##0.###", 6: "#,##0", 7: "#,##0", 11: "#,##0.###"})
    _sheet(wb, "K8 Nguon ky truoc", "K8 - Nguồn KL kỳ trước (số hạng mục có KL kỳ trước)", ["Nguồn", "Số hạng mục"], [40, 12],
           k.get("tang_kt", []))
    wb.save(out)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description="Module 2 - Lớp 1: tính toàn vẹn file hồ sơ thanh toán")
    ap.add_argument("file")
    ap.add_argument("--o", help="ô đích cho cây truy vết, VD \"2. TH.!D38\"")
    ap.add_argument("--out")
    ar = ap.parse_args(argv)
    kt = KiemTra(ar.file, ar.o).chay()
    out = Path(ar.out or Path(ar.file).with_name(f"toan_ven_{Path(ar.file).stem[:60]}.xlsx"))
    ghi(kt, out)
    dem = Counter(x["muc"] for x in kt.ph)
    print(f"Phát hiện: Cao {dem['Cao']} · Trung bình {dem['Trung bình']} · Kiểm soát {dem['Kiểm soát']}")
    for x in kt.ph[:25]:
        print(f"  [{x['muc']:^10}] {x['kt']} {x['o'][:30]:30} {x['mo_ta'][:110]}")
    if "tang" in kt.kq:
        print("Nguồn KL lũy kế:", "; ".join(f"{a}: {b}" for a, b, *_ in kt.kq["tang"]))
    print("Kết quả:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
