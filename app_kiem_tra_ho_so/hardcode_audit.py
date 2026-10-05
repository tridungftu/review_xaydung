# -*- coding: utf-8 -*-
"""
hardcode_audit.py - Quét số nhập tay trong hồ sơ thanh toán khối lượng (TTKL)
và lượng hoá giá trị bị ảnh hưởng.

Cách dùng:
    python hardcode_audit.py <file_ho_so.xlsx> [file_ket_qua.xlsx]

Chuỗi truy vết mặc định (đổi trong CFG nếu hồ sơ khác tên sheet):
    2. TH.  <-  3. GTHT  <-  4. BTHKLTH  <-  5. BBNT  <-  6. DG KL  (<- 7. TK THEP)

Nguyên tắc:
- Đọc file 2 lần: một lần lấy công thức, một lần lấy giá trị đã lưu.
- Với mỗi hạng mục ở GTHT, lần theo công thức để biết khối lượng lũy kế
  và khối lượng kỳ trước đến từ đâu (diễn giải, nhập tay, = KL HĐ, % KL HĐ...).
- Giá trị bị ảnh hưởng = giá trị của hạng mục phụ thuộc vào ô nhập tay đó.
- Ảnh hưởng tới số đề nghị thanh toán F1 được tính bằng công thức Excel
  trong file kết quả (VAT, tỷ lệ giữ lại là ô tham số, có thể sửa).
Script chỉ đọc file hồ sơ, không ghi đè.
"""
import re
import sys
import collections
import difflib
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

CFG = dict(
    TH="2. TH.", GTHT="3. GTHT", BTH="4. BTHKLTH", BBNT="5. BBNT",
    DG="6. DG KL", TKT="7. TK THEP",
    VAT=0.08, RET_XD=0.05, RET_CN=0.30,
    # Hệ số suy ra từ đơn giá của mẫu hồ sơ (GTHT!I440 = 100.000 x k); đổi theo hồ sơ thực tế
    K_GIA=0.964809217320883,
)

# --------------------------------------------------------------------------
REF_RE = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_.]+))!\$?([A-Z]{1,3})\$?(\d+)")


def refs(formula):
    """Trả về danh sách (sheet, col, row) mà công thức tham chiếu."""
    if not isinstance(formula, str):
        return []
    out = []
    for m in REF_RE.finditer(formula):
        out.append(((m.group(1) or m.group(2)), m.group(3), int(m.group(4))))
    return out


def ref_row(formula, sheet):
    for s, c, r in refs(formula):
        if s == sheet:
            return r
    return None


def is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


# Mô tả và mức rủi ro cho từng loại nguồn khối lượng lũy kế
LK_SRC = {
    "DG_DIENGIAI":     ("Diễn giải kích thước (SUM/PRODUCT)", "Thấp"),
    "DG_CT_KHAC":      ("Công thức khác trong DG KL", "Trung bình"),
    "DG_NHAPTAY_THEP": ("Số dán tay, khớp bảng TK thép", "Trung bình"),
    "DG_NHAPTAY":      ("Số nhập tay trong DG KL", "Cao"),
    "DG_BANG_KLHD":    ("DG KL lấy = KL hợp đồng", "Cao"),
    "DG_TYLE":         ("DG KL = % x KL hợp đồng", "Cao"),
    "DG_TRONG":        ("Link lệch dòng / ô nguồn trống", "Cao"),
    "BBNT_G_BANG_KLHD": ("Kỳ này = KL HĐ (công thức G=E)", "Cao"),
    "BBNT_G_NHAPTAY":  ("Kỳ này nhập tay tại BBNT", "Cao"),
    "BBNT_G_KHAC":     ("Công thức khác tại BBNT", "Trung bình"),
}
LK_DOC = {
    "DG_DIENGIAI": "B01 bản vẽ thi công / B03 hoàn công",
    "DG_CT_KHAC": "B06 diễn giải KL",
    "DG_NHAPTAY_THEP": "B05 thống kê thép gốc, bản vẽ kết cấu",
    "DG_NHAPTAY": "B06 diễn giải KL / B03 hoàn công",
    "DG_BANG_KLHD": "B03 hoàn công / C01 BB nghiệm thu công việc",
    "DG_TYLE": "C05 xác nhận % tiến độ của TVGS",
    "DG_TRONG": "F02 giải trình + sửa liên kết",
    "BBNT_G_BANG_KLHD": "D01 BB giao nhận, D07 BB lắp đặt, D03/D04 CO-CQ",
    "BBNT_G_NHAPTAY": "D01 BB giao nhận, D07 BB lắp đặt / C02 BB nghiệm thu",
    "BBNT_G_KHAC": "C02 BB nghiệm thu",
}
KT_SRC = {
    "NHAPTAY":   "Nhập tay",
    "TYLE":      "% x KL HĐ",
    "BANG_KLHD": "Bằng KL HĐ",
    "SHEET_AN":  "Trỏ sang sheet ẩn",
    "KHAC":      "Công thức khác",
}


def classify_dg(formula, dg_row, m_note):
    if formula is None:
        return "DG_TRONG"
    if not is_formula(formula):
        return "DG_NHAPTAY_THEP" if "thống k" in norm(m_note) else "DG_NHAPTAY"
    f = formula.replace("+", "").replace(" ", "")
    if re.fullmatch(rf"=D{dg_row}", f):
        return "DG_BANG_KLHD"
    if re.search(r"\d+%\*D\d+|D\d+\*\d+%", f):
        return "DG_TYLE"
    if "SUM(K" in f.upper() or "PRODUCT(" in f.upper():
        return "DG_DIENGIAI"
    return "DG_CT_KHAC"


def classify_kt(formula, e_ref_row):
    if not is_formula(formula):
        return "NHAPTAY"
    f = formula.replace("+", "").replace(" ", "")
    if "DGPLH" in f.upper().replace(" ", "") or "PLHĐ" in formula:
        return "SHEET_AN"
    if re.fullmatch(rf"=E{e_ref_row}", f):
        return "BANG_KLHD"
    if re.search(rf"E{e_ref_row}\*", f):
        return "TYLE"
    return "KHAC"


def price_pattern(p, k):
    if p is None:
        return ""
    if float(p).is_integer():
        return "Số nguyên"
    base = p / k
    if abs(base - round(base, -2)) < 0.01:
        return "Giá tròn x k"
    if abs(base * 1.08 - round(base * 1.08, -2)) < 0.01:
        return "Giá tròn ÷ 1,08 x k"
    return "Lẻ, chưa rõ nguồn"


# --------------------------------------------------------------------------
def analyse(path):
    wb = openpyxl.load_workbook(path)                  # công thức
    wv = openpyxl.load_workbook(path, data_only=True)  # giá trị đã lưu
    c = CFG
    g, gv = wb[c["GTHT"]], wv[c["GTHT"]]
    t, tv = wb[c["BTH"]], wv[c["BTH"]]
    n, nv = wb[c["BBNT"]], wv[c["BBNT"]]
    d, dv = wb[c["DG"]], wv[c["DG"]]

    # Bảng tổng thép để nhận diện số dán tay
    steel_totals = {}
    if c["TKT"] in wv.sheetnames:
        ts = wv[c["TKT"]]
        for r in range(1, ts.max_row + 1):
            if "ng kh" in norm(ts[f"B{r}"].value):  # 'tæng khèi l­îng' (font TCVN3)
                v = num(ts[f"R{r}"].value)
                if v:
                    steel_totals[round(v, 9)] = r

    items, issues = [], []
    part = None
    for r in range(1, g.max_row + 1):
        b = norm(g[f"B{r}"].value)
        if b.startswith("phần xây dựng"):
            part = "XD"
        elif b.startswith("phần thiết bị"):
            part = "TB"
        elif b.startswith("theo phụ lục"):
            part = "PL"
        e_raw = g[f"E{r}"].value
        r2 = ref_row(g[f"F{r}"].value, c["BTH"])
        if part is None or e_raw in (None, "") or r2 is None:
            continue
        r3 = ref_row(t[f"F{r2}"].value, c["BBNT"])
        if r3 is None:
            issues.append(("Liên kết", f"{c['BTH']}!F{r2}", "Không trỏ về BBNT", ""))
            continue

        it = dict(
            gtht_row=r, part=part, name=str(g[f"B{r}"].value).strip(),
            unit=str(g[f"D{r}"].value or "").strip(),
            E=num(gv[f"E{r}"].value), I=num(gv[f"I{r}"].value),
            F=num(gv[f"F{r}"].value) or 0, G=num(gv[f"G{r}"].value) or 0,
            H=num(gv[f"H{r}"].value) or 0,
            K=num(gv[f"K{r}"].value) or 0, L=num(gv[f"L{r}"].value) or 0,
            M=num(gv[f"M{r}"].value) or 0,
            bth_row=r2, bbnt_row=r3, flags=[],
        )
        it["price_pattern"] = price_pattern(it["I"], c["K_GIA"]) if part == "XD" else (
            "Số nguyên" if it["I"] is not None and float(it["I"]).is_integer() else "Lẻ")

        # --- nguồn KL lũy kế
        hf = n[f"H{r3}"].value
        r4 = ref_row(hf, c["DG"])
        it["dg_row"] = r4
        it["dgL"] = None
        if r4:
            src = classify_dg(d[f"L{r4}"].value, r4, d[f"M{r4}"].value)
            it["lk_cell"] = f"{c['DG']}!L{r4}"
            it["lk_formula"] = str(d[f"L{r4}"].value)[:80] if d[f"L{r4}"].value is not None else "(trống)"
            it["dgL"] = num(dv[f"L{r4}"].value)
            dl_raw = it["dgL"] or 0  # giá trị Excel thực sự dùng, kể cả khi lệch dòng
            ratio = difflib.SequenceMatcher(None, norm(n[f"B{r3}"].value), norm(d[f"B{r4}"].value)).ratio()
            if ratio < 0.85:
                it["flags"].append(f"Link lệch: BBNT!H{r3} trỏ tới DG KL dòng {r4} "
                                   f"('{str(d[f'B{r4}'].value).strip()[:40]}')")
                src = "DG_TRONG"
                it["lk_formula"] += " [lệch dòng]"
                it["dgL"] = None
            if src == "DG_NHAPTAY_THEP" and it["dgL"] is not None:
                tr = steel_totals.get(round(it["dgL"], 9))
                if tr:
                    it["flags"].append(f"Khớp TK THEP!R{tr} (mất liên kết)")
            # nhánh của công thức chặn trần ở BBNT!H
            dl = dl_raw
            Eq, Fq = num(nv[f"E{r3}"].value) or 0, num(nv[f"F{r3}"].value) or 0
            if dl < Eq and dl < Fq:
                it["branch"] = "Giữ KL kỳ trước (DG < kỳ trước)"
                if it["dgL"] is not None and src != "DG_TRONG" and Fq - dl > 1e-6 * max(1, Fq):
                    it["flags"].append("KL diễn giải thấp hơn KL đã nghiệm thu kỳ trước")
            elif dl > Eq:
                it["branch"] = "Chặn trần = KL HĐ (DG > HĐ)"
                if Eq > 0 and dl / Eq > 3 and src != "DG_TRONG":
                    it["flags"].append(f"KL diễn giải gấp {dl / Eq:.1f} lần KL HĐ - kiểm tra đơn vị tính (m² / 100m²...)")
            else:
                it["branch"] = "Theo DG KL"
        else:
            gf = n[f"G{r3}"].value
            if is_formula(gf) and re.fullmatch(rf"=\+?E{r3}", gf.replace(" ", "")):
                src = "BBNT_G_BANG_KLHD"
            elif not is_formula(gf):
                src = "BBNT_G_NHAPTAY"
            else:
                src = "BBNT_G_KHAC"
            it["lk_cell"] = f"{c['BBNT']}!G{r3}"
            it["lk_formula"] = str(gf)[:80]
            it["branch"] = "H = F + G"
        it["lk_src"] = src

        # --- nguồn KL kỳ trước
        ff = n[f"F{r3}"].value
        it["kt_src"] = classify_kt(ff, r3)
        it["kt_cell"] = f"{c['BBNT']}!F{r3}"
        it["kt_formula"] = str(ff)[:80]

        # --- kiểm tra kiểu dữ liệu và tính nhất quán KL HĐ giữa các sheet
        for sh, rr in ((t, r2), (n, r3)):
            v = sh[f"E{rr}"].value
            if isinstance(v, str) and not v.startswith("="):
                it["flags"].append(f"KL HĐ lưu dạng text tại {sh.title}!E{rr}")
        e_vals = {num(gv[f"E{r}"].value), num(tv[f"E{r2}"].value), num(nv[f"E{r3}"].value)}
        if len({x for x in e_vals if x is not None}) > 1:
            it["flags"].append("KL HĐ khác nhau giữa GTHT/BTHKLTH/BBNT")
        if it["E"] is not None and it["E"] < 0:
            it["flags"].append("KL âm (hạng mục bị thay thế ở PLHĐ)")
        items.append(it)

    # --- PLHĐ: dòng gốc bị thay thế nhưng KL HĐ gốc vẫn giữ nguyên
    base = collections.defaultdict(list)
    for it in items:
        if it["part"] in ("TB", "XD"):
            base[norm(it["name"])[:35]].append(it)
    for it in items:
        if it["part"] == "PL" and (it["E"] or 0) < 0:
            for o in base.get(norm(it["name"])[:35], []):
                o["flags"].append(f"Đã bị thay thế ở PLHĐ (GTHT dòng {it['gtht_row']}) nhưng KL HĐ gốc chưa trừ")
    return wb, wv, items, issues, steel_totals


def th_inputs(wb, wv):
    """Số nhập tay và tham số nhúng trong công thức ở bảng TH."""
    c = CFG
    s, sv = wb[c["TH"]], wv[c["TH"]]
    rows = []
    for row in s.iter_rows():
        for cell in row:
            v = cell.value
            if v is None:
                continue
            if cell.column >= 8:
                colh = s.cell(16, cell.column).value or ("Tổng" if cell.column in (9, 12) else "")
                label = f"Lũy kế thanh toán - {s['H' + str(cell.row)].value or ''} - {colh}".strip(" -")
            else:
                label = f"{s['A' + str(cell.row)].value or ''} {s['B' + str(cell.row)].value or ''}".strip()
            if isinstance(v, (int, float)) and cell.column > 3:
                rows.append((f"{s.title}!{cell.coordinate}", label, "Số nhập tay", v,
                             sv[cell.coordinate].value))
            elif is_formula(v):
                if "[" in v:
                    rows.append((f"{s.title}!{cell.coordinate}", label, "Liên kết file ngoài", v,
                                 sv[cell.coordinate].value))
                body = REF_RE.sub("REF", v)
                body = re.sub(r"\$?[A-Z]{1,3}\$?\d+", "REF", body)
                consts = re.findall(r"\d+(?:\.\d+)?%|\d+\.\d+", body)
                if consts:
                    rows.append((f"{s.title}!{cell.coordinate}", label, "Tham số nhúng: " + ", ".join(sorted(set(consts))), v,
                                 sv[cell.coordinate].value))
    return rows


def integrity(wb, items):
    out = []
    for ws in wb.worksheets:
        nref = sum(1 for row in ws.iter_rows() for cl in row
                   if isinstance(cl.value, str) and "#REF!" in cl.value)
        if ws.sheet_state != "visible":
            out.append(("Sheet ẩn", ws.title, f"Trạng thái: {ws.sheet_state}", ""))
        if nref:
            out.append(("Lỗi tham chiếu (REF)", ws.title, f"{nref} công thức bị mất tham chiếu (lỗi REF)", ""))
    for lk in wb._external_links:
        tgt = lk.file_link.Target if lk.file_link else "?"
        out.append(("Liên kết file ngoài", "Workbook", tgt, ""))
    for it in items:
        for f in it["flags"]:
            out.append(("Cờ hạng mục", f"{CFG['GTHT']}!B{it['gtht_row']}", f, it["name"][:60]))
    return out



# --------------------------------------------------------------------------
# CÂY TRUY VẾT: dữ liệu nào x dữ liệu nào ra số đề nghị thanh toán
# --------------------------------------------------------------------------
QREF = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_.\[\]]+))!\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?")
LREF = re.compile(r"(?<![A-Za-z_!'\]])\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![\d(A-Za-z])")


def parse_refs(formula, cur_sheet):
    """Tách tham chiếu trong công thức -> list (sheet, c1, r1, c2, r2, external)."""
    out = []
    body = formula
    for m in QREF.finditer(formula):
        sh = m.group(1) or m.group(2)
        ext = sh.startswith("[")
        out.append((sh, m.group(3), int(m.group(4)), m.group(5) or m.group(3),
                    int(m.group(6) or m.group(4)), ext))
    body = QREF.sub(" ", formula)
    for m in LREF.finditer(body):
        out.append((cur_sheet, m.group(1), int(m.group(2)), m.group(3) or m.group(1),
                    int(m.group(4) or m.group(2)), False))
    return out


def embedded_params(formula):
    body = QREF.sub(" ", formula)
    body = LREF.sub(" ", body)
    return sorted(set(re.findall(r"\d+(?:\.\d+)?%|\d+\.\d+", body)))


def row_label(ws, row, col):
    if ws.title == CFG["TH"] and col >= 8:
        return f"Thanh toán {ws['H' + str(row)].value or ''} - {ws.cell(16, col).value or 'Tổng'}"
    parts = []
    for c in ("A", "B"):
        v = ws[f"{c}{row}"].value
        if v not in (None, "") and not is_formula(v):
            parts.append(str(v).strip())
    return " ".join(parts)[:90]


def col_header(ws, col):
    """Tên cột: đọc dòng tiêu đề trong 20 dòng đầu."""
    names = []
    for r in range(8, 18):
        v = ws.cell(r, col).value
        if isinstance(v, str) and not v.startswith("=") and len(v) < 60 and not re.fullmatch(r"\[\d+\]", v.strip()):
            names.append(re.sub(r"\s+", " ", v).strip())
    return " / ".join(names[-2:])


def doc_for(sheet, col_letter, kind, formula=""):
    c = CFG
    if kind == "File ngoài":
        return "E07 file Excel đang liên kết"
    if kind.startswith("Tham số"):
        if "1.08" in kind or "thuế" in formula.lower():
            return "F01 căn cứ thuế suất GTGT"
        if sheet == c["DG"] or sheet == c["BBNT"]:
            return "C05 xác nhận % tiến độ của TVGS"
        return "A01 điều khoản HĐ (tạm ứng / thu hồi / giữ lại / hóa đơn)"
    if kind != "Nhập tay":
        return ""
    if sheet == c["TH"]:
        return "E01-E04 hồ sơ thanh toán các đợt đã duyệt, UNC"
    if sheet == c["GTHT"] and col_letter == "I":
        return "A02 BOQ / A04 phụ lục (đơn giá)"
    if col_letter == "E" and sheet in (c["GTHT"], c["BTH"], c["BBNT"]):
        return "A02 BOQ / A04 phụ lục (KL hợp đồng)"
    if sheet == c["BBNT"] and col_letter == "F":
        return "E03 hồ sơ Đợt 3 đã ký (KL lũy kế kỳ trước)"
    if sheet == c["BBNT"] and col_letter == "G":
        return "C02 BB nghiệm thu / D01 BB giao nhận thiết bị"
    if sheet == c["DG"] and col_letter == "D":
        return "A02 BOQ / A04 phụ lục (KL hợp đồng)"
    if sheet == c["DG"] and col_letter == "L":
        return "B06 diễn giải KL / B03 hoàn công"
    if sheet == c["DG"]:
        return "B01 bản vẽ thi công / B03 hoàn công (kích thước)"
    if sheet == c["TKT"]:
        return "B05 bản vẽ kết cấu, thống kê thép"
    return "Cần xác định nguồn"


def norm_pattern(f):
    f = f.replace("$", "").replace("+", "").replace(" ", "")
    return re.sub(r"([A-Z]{1,3})\d+", r"\1#", f)


def lineage(wb, wv, sheet=None, coord="D38", max_depth=16, items=None):
    """Phân rã ô đích thành cây. Dải nhiều ô: phân rã 1 dòng đại diện cho mỗi mẫu công thức."""
    sheet = sheet or CFG["TH"]
    rows, done = [], {}

    def val(sh, cd):
        try:
            return wv[sh][cd].value
        except KeyError:
            return None

    def add(level, text, cell, formula, value, kind, doc, note=""):
        rows.append(dict(level=level, text=text, cell=cell, formula=formula, value=value,
                         kind=kind, doc=doc, note=note))
        return len(rows) + 1  # số dòng trên sheet (có tiêu đề)

    def visit(sh, col, r, level, prefix=""):
        cd = f"{col}{r}"
        key = f"{sh}!{cd}"
        if sh not in wb.sheetnames:
            add(level, prefix + "Tham chiếu ngoài workbook", key, "", None, "File ngoài", doc_for(sh, col, "File ngoài"))
            return
        ws = wb[sh]
        f = ws[cd].value
        lab = row_label(ws, r, ws[cd].column)
        hdr = col_header(ws, ws[cd].column)
        text = prefix + (lab or "") + (f"  [{hdr}]" if hdr else "")
        if key in done:
            add(level, text, key, "", val(sh, cd), "Đã phân rã", "", f"xem dòng {done[key]}")
            return
        if not is_formula(f):
            kind = "Nhập tay" if isinstance(f, (int, float)) else ("Trống" if f in (None, "") else "Văn bản")
            if isinstance(f, str) and num(f) is not None:
                kind = "Nhập tay"
            done[key] = add(level, text, key, "", f, kind, doc_for(sh, col, kind))
            return
        done[key] = add(level, text, key, f, val(sh, cd), "Công thức", "")
        if level >= max_depth:
            add(level + 1, "… dừng do quá sâu", "", "", None, "", "")
            return
        for p in embedded_params(f):
            add(level + 1, f"Tham số trong công thức: {p}", key, "", p, f"Tham số {p}",
                doc_for(sh, col, f"Tham số {p}", lab or ""))
        seen_refs = set()
        for (rs, c1, r1, c2, r2, ext) in parse_refs(f, sh):
            if (rs, c1, r1, c2, r2) in seen_refs:
                continue
            seen_refs.add((rs, c1, r1, c2, r2))
            if ext:
                add(level + 1, f"Liên kết file ngoài {rs}", f"{rs}!{c1}{r1}", "", None, "File ngoài",
                    doc_for(rs, c1, "File ngoài"))
                continue
            if (c1, r1) == (c2, r2):
                visit(rs, c1, r1, level + 1)
                continue
            # dải ô: gom theo mẫu công thức, phân rã 1 dòng đại diện cho mỗi mẫu
            wsr = wb[rs]
            groups = collections.OrderedDict()
            n_input = 0
            from openpyxl.utils import column_index_from_string
            for cc in range(column_index_from_string(c1), column_index_from_string(c2) + 1):
                for rr in range(r1, r2 + 1):
                    v = wsr.cell(rr, cc).value
                    if is_formula(v):
                        if re.match(r"=\+?(SUBTOTAL|SUM)\(", v.replace(" ", ""), re.I):
                            continue  # dòng tổng nhóm
                        groups.setdefault(norm_pattern(v), []).append((cc, rr))
                    elif isinstance(v, (int, float)):
                        n_input += 1
            rng_txt = f"{rs}!{c1}{r1}:{c2}{r2}"
            add(level + 1, f"Tổng dải {rng_txt}", rng_txt, "", None, "Dải",
                "", f"{sum(len(x) for x in groups.values())} ô công thức theo {len(groups)} mẫu; {n_input} ô nhập tay")
            if items and rs == CFG["GTHT"] and c1 == c2 and c1 in ("K", "L", "M"):
                sub = [it for it in items if r1 <= it["gtht_row"] <= r2]
                agg = collections.OrderedDict()
                for it in sub:
                    if c1 == "K":
                        k = ("KL kỳ trước: " + KT_SRC[it["kt_src"]],
                             doc_for(CFG["BBNT"], "F", "Nhập tay") if it["kt_src"] == "NHAPTAY" else
                             ("C05 xác nhận % tiến độ của TVGS" if it["kt_src"] == "TYLE" else "Xem sheet Chi tiet"))
                    else:
                        desc = LK_SRC[it["lk_src"]][0]
                        k = (("KL kỳ này: " if c1 == "L" else "KL lũy kế: ") + desc, LK_DOC.get(it["lk_src"], ""))
                    a = agg.setdefault(k, [0, 0.0])
                    a[0] += 1
                    a[1] += it[c1]
                for (desc, d), (n, v) in sorted(agg.items(), key=lambda x: -x[1][1]):
                    if v or n:
                        add(level + 2, f"Nhóm nguồn - {desc} x Đơn giá (cột I)", rng_txt, "", v, "Nhóm nguồn",
                            d, f"{n} hạng mục")
            for pat, cells in groups.items():
                cc, rr = cells[0]
                from openpyxl.utils import get_column_letter as gcl
                visit(rs, gcl(cc), rr, level + 2, prefix=f"Mẫu '{pat}' x {len(cells)} dòng, đại diện: ")

    visit(sheet, re.match(r"[A-Z]+", coord).group(0), int(re.search(r"\d+", coord).group(0)), 0)
    return rows


def write_lineage(wb_out, rows, first=True):
    ws = wb_out.create_sheet("Cay truy vet", 0 if first else None)
    ws["A1"] = "CÂY TRUY VẾT SỐ ĐỀ NGHỊ THANH TOÁN: DỮ LIỆU NÀO x DỮ LIỆU NÀO RA SỐ ĐÓ"
    ws["A1"].font = Font(name="Arial", bold=True, size=12)
    ws["A2"] = ("Đọc từ trên xuống: mỗi cấp thụt vào là thành phần tạo nên dòng phía trên. Nút lá (Nhập tay, Tham số, "
                "File ngoài) là chỗ cần tài liệu chứng minh. Dải nhiều dòng được đại diện bằng 1 dòng cho mỗi mẫu công thức; "
                "chi tiết từng hạng mục xem sheet 'Chi tiet'.")
    ws["A2"].font = Font(name="Arial", size=9, italic=True)
    ws["A2"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 42
    H = ["Cấp", "Thành phần", "Ô", "Công thức", "Giá trị", "Loại nút", "Tài liệu chứng minh", "Ghi chú"]
    hdr(ws, 3, H, [6, 70, 22, 48, 18, 14, 40, 34])
    fills = {"Nhập tay": "F8D7D3", "File ngoài": "F8D7D3", "Dải": "E3EBF5", "Nhóm nguồn": "FCEBCB"}
    for i, rw in enumerate(rows, 4):
        vals = [rw["level"], "    " * rw["level"] + (rw["text"] or ""), rw["cell"], rw["formula"],
                rw["value"], rw["kind"], rw["doc"], rw["note"]]
        for j, v in enumerate(vals, 1):
            cl = put(ws, i, j, v)
            cl.font = BOLD if rw["level"] == 0 else BASE
        ws.cell(i, 5).number_format = '#,##0.###;(#,##0.###);"-"'
        k = rw["kind"]
        f = fills.get(k) or ("FCEBCB" if k.startswith("Tham số") else None)
        if f:
            ws.cell(i, 6).fill = PatternFill("solid", fgColor=f)
    ws.freeze_panes = "C4"
    # tóm tắt nút lá
    r = len(rows) + 6
    ws.cell(r, 2, "LOẠI DỮ LIỆU GỐC VÀ TÀI LIỆU CHỨNG MINH (đếm theo nút trên cây, không phải số hạng mục)").font = BOLD
    leaves = collections.Counter((rw["kind"], rw["doc"]) for rw in rows
                                 if rw["kind"] in ("Nhập tay", "File ngoài") or rw["kind"].startswith("Tham số"))
    for (k, d), n in sorted(leaves.items()):
        r += 1
        put(ws, r, 2, f"{k}: {n} nút").font = BASE
        put(ws, r, 7, d).font = BASE
    return ws

# --------------------------------------------------------------------------
F1_VALUE = None
THIN = Side(style="thin", color="BFC5C2")
HDR_FILL = PatternFill("solid", fgColor="1F4A7A")
HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BASE = Font(name="Arial", size=10)
BLUE = Font(name="Arial", size=10, color="0000FF")
BOLD = Font(name="Arial", size=10, bold=True)
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
RISK_FILL = {"Cao": "F8D7D3", "Trung bình": "FCEBCB", "Thấp": "DDEFE3"}
NUMF = '#,##0;(#,##0);"-"'
QF = '#,##0.000;(#,##0.000);"-"'


def put(ws, r, c, v):
    """Ghi giá trị; chuỗi bắt đầu bằng '=' được ghi dạng text, không thành công thức."""
    cl = ws.cell(r, c, v)
    if isinstance(v, str) and v.startswith("="):
        cl.data_type = "s"
    return cl


def hdr(ws, row, headers, widths=None):
    for j, h in enumerate(headers, 1):
        cl = ws.cell(row, j, h)
        cl.font, cl.fill = HDR_FONT, HDR_FILL
        cl.alignment = Alignment(wrap_text=True, vertical="center")
        cl.border = Border(bottom=THIN)
        if widths:
            ws.column_dimensions[get_column_letter(j)].width = widths[j - 1]
    ws.row_dimensions[row].height = 42


def write_report(items, th_rows, integ, src_path, out_path, f1_value=None, tree=None):
    global F1_VALUE
    F1_VALUE = f1_value
    wb = openpyxl.Workbook()

    # ---------------- Tham so
    p = wb.active
    p.title = "Tham so"
    p["A1"] = "THAM SỐ TÍNH ẢNH HƯỞNG LÊN SỐ ĐỀ NGHỊ THANH TOÁN (F1)"
    p["A1"].font = Font(name="Arial", bold=True, size=12)
    params = [
        ("Thuế suất GTGT", CFG["VAT"], "2. TH.: các dòng A x 1,08"),
        ("Tỷ lệ giữ lại phần xây dựng", CFG["RET_XD"], "2. TH.!D30: 5% x A.6.1"),
        ("Tỷ lệ giữ lại phần công nghệ (TB + PLHĐ)", CFG["RET_CN"], "2. TH.!D30: 30% x A.5.2"),
        ("Hệ số đơn giá k (suy ra)", CFG["K_GIA"], "GTHT!I440 = 100.000 x k; cần đối chiếu HĐ"),
    ]
    p["A3"], p["B3"], p["C3"] = "Tham số", "Giá trị", "Nguồn trong file hồ sơ"
    for j in range(1, 4):
        p.cell(3, j).font, p.cell(3, j).fill = HDR_FONT, HDR_FILL
    for i, (a, b, cc) in enumerate(params, 4):
        p.cell(i, 1, a).font = BASE
        x = p.cell(i, 2, b)
        x.font, x.fill = BLUE, INPUT_FILL
        x.number_format = "0.0%" if b < 0.5 else "0.000000000"
        p.cell(i, 3, cc).font = BASE
    p["A9"] = "Hệ số ảnh hưởng lên F1 (1 đồng giá trị trước VAT bị loại thì F1 giảm bao nhiêu)"
    p["A9"].font = BOLD
    p["A10"], p["B10"] = "Phần xây dựng", "=(1+B4)*(1-B5)"
    p["A11"], p["B11"] = "Phần công nghệ - giá trị kỳ này", "=(1+B4)*(1-B6)"
    p["A12"], p["B12"] = "Phần công nghệ - giá trị kỳ trước", "=(1+B4)"
    for r in (10, 11, 12):
        p[f"A{r}"].font = BASE
        p[f"B{r}"].font = BASE
        p[f"B{r}"].number_format = "0.0000"
    p["A14"] = ("Ghi chú: F1 = A.6 - B.1 - B.2. Khoản đã thanh toán (B.2) cố định, nên mỗi đồng giá trị lũy kế "
                "bị loại làm F1 giảm theo hệ số trên. Sửa ô vàng nếu hợp đồng quy định khác.")
    p["A15"] = f"File nguồn: {src_path}"
    for r in (14, 15):
        p[f"A{r}"].font = Font(name="Arial", size=9, italic=True)
    p.column_dimensions["A"].width = 52
    p.column_dimensions["B"].width = 16
    p.column_dimensions["C"].width = 48

    # ---------------- Chi tiet
    ws = wb.create_sheet("Chi tiet")
    H = ["GTHT dòng", "Phần", "Hạng mục", "ĐVT", "KL HĐ", "Đơn giá", "Dạng đơn giá",
         "KL kỳ trước", "KL kỳ này", "KL lũy kế",
         "GT kỳ trước", "GT kỳ này", "GT lũy kế",
         "Nguồn KL lũy kế", "Mô tả nguồn", "Mức rủi ro", "Ô nguồn lũy kế", "Nội dung ô nguồn",
         "Nhánh công thức BBNT!H", "Nguồn KL kỳ trước", "Ô nguồn kỳ trước", "Nội dung ô kỳ trước",
         "KL theo DG KL", "KL vượt HĐ", "GT vượt HĐ (bị cắt)", "GT kỳ trước cao hơn diễn giải",
         "Ảnh hưởng F1 - phần kỳ này", "Ảnh hưởng F1 - toàn bộ lũy kế", "Cờ"]
    W = [9, 6, 42, 8, 11, 13, 16, 11, 11, 11, 15, 15, 15, 18, 30, 11, 16, 26, 26, 14, 14, 26, 11, 11, 15, 15, 16, 16, 50]
    hdr(ws, 1, H, W)
    for i, it in enumerate(items, 2):
        desc, risk = LK_SRC[it["lk_src"]]
        vals = [it["gtht_row"], it["part"], it["name"][:120], it["unit"], it["E"], it["I"], it["price_pattern"],
                it["F"], it["G"], it["H"], it["K"], it["L"], it["M"],
                it["lk_src"], desc, risk, it["lk_cell"], it["lk_formula"], it.get("branch", ""),
                KT_SRC[it["kt_src"]], it["kt_cell"], it["kt_formula"], it["dgL"]]
        for j, v in enumerate(vals, 1):
            cl = put(ws, i, j, v)
            cl.font = BASE
        # công thức
        ws.cell(i, 24, f'=IF(W{i}="","",MAX(0,W{i}-E{i}))')
        ws.cell(i, 25, f'=IF(X{i}="",0,X{i}*F{i})')
        ws.cell(i, 26, f'=IF(W{i}="",0,MAX(0,H{i}-W{i})*F{i})')
        ws.cell(i, 27, (f"=L{i}*'Tham so'!$B$10" if it["part"] == "XD"
                        else f"=L{i}*'Tham so'!$B$11"))
        ws.cell(i, 28, (f"=M{i}*'Tham so'!$B$10" if it["part"] == "XD"
                        else f"=L{i}*'Tham so'!$B$11+K{i}*'Tham so'!$B$12"))
        put(ws, i, 29, "; ".join(it["flags"]))
        for j in (24, 25, 26, 27, 28, 29):
            ws.cell(i, j).font = BASE
        for j in (5, 8, 9, 10, 23, 24):
            ws.cell(i, j).number_format = QF
        for j in (6, 11, 12, 13, 25, 26, 27, 28):
            ws.cell(i, j).number_format = NUMF
        ws.cell(i, 16).fill = PatternFill("solid", fgColor=RISK_FILL[risk])
    last = len(items) + 1
    tab = Table(displayName="ChiTiet", ref=f"A1:{get_column_letter(len(H))}{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=True)
    ws.add_table(tab)
    ws.freeze_panes = "D2"

    # ---------------- Tong hop
    s = wb.create_sheet("Tong hop", 0)
    s["A1"] = "TỔNG HỢP GIÁ TRỊ PHỤ THUỘC SỐ NHẬP TAY / NGUỒN CHƯA CÓ CĂN CỨ"
    s["A1"].font = Font(name="Arial", bold=True, size=12)
    s["A2"] = "Đơn vị: đồng, giá trị trước VAT (trừ cột ảnh hưởng F1 đã gồm VAT và giữ lại). Tất cả là công thức trên sheet 'Chi tiet'."
    s["A2"].font = Font(name="Arial", size=9, italic=True)
    CT = "'Chi tiet'"
    R = f"$2:${last}"

    def rng(col):
        return f"{CT}!${col}$2:${col}${last}"

    # Bảng 1: theo nguồn KL lũy kế
    hdr(s, 4, ["Phần", "Nguồn KL lũy kế", "Mô tả", "Mức rủi ro", "Số hạng mục",
               "GT kỳ này", "GT lũy kế", "F1 giảm nếu loại phần kỳ này", "F1 giảm nếu loại toàn bộ lũy kế", "% GT lũy kế"],
        [8, 20, 34, 12, 11, 17, 17, 19, 19, 11])
    combos = []
    seen = set()
    for it in items:
        k = (it["part"], it["lk_src"])
        if k not in seen:
            seen.add(k)
            combos.append(k)
    order = {"XD": 0, "TB": 1, "PL": 2}
    rorder = {"Cao": 0, "Trung bình": 1, "Thấp": 2}
    combos.sort(key=lambda k: (order[k[0]], rorder[LK_SRC[k[1]][1]], k[1]))
    r = 5
    first = r
    for part, src in combos:
        desc, risk = LK_SRC[src]
        s.cell(r, 1, part)
        s.cell(r, 2, src)
        s.cell(r, 3, desc)
        s.cell(r, 4, risk).fill = PatternFill("solid", fgColor=RISK_FILL[risk])
        s.cell(r, 5, f'=COUNTIFS({rng("B")},A{r},{rng("N")},B{r})')
        s.cell(r, 6, f'=SUMIFS({rng("L")},{rng("B")},A{r},{rng("N")},B{r})')
        s.cell(r, 7, f'=SUMIFS({rng("M")},{rng("B")},A{r},{rng("N")},B{r})')
        s.cell(r, 8, f'=SUMIFS({rng("AA")},{rng("B")},A{r},{rng("N")},B{r})')
        s.cell(r, 9, f'=SUMIFS({rng("AB")},{rng("B")},A{r},{rng("N")},B{r})')
        s.cell(r, 10, f"=IF($G${first + len(combos)}=0,0,G{r}/$G${first + len(combos)})")
        r += 1
    tot1 = r
    s.cell(r, 1, "Tổng")
    for j, col in ((5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I"), (10, "J")):
        s.cell(r, j, f"=SUM({col}{first}:{col}{r - 1})")
    for rr in range(first, r + 1):
        for j in range(1, 11):
            s.cell(rr, j).font = BOLD if rr == r else BASE
        for j in (6, 7, 8, 9):
            s.cell(rr, j).number_format = NUMF
        s.cell(rr, 10).number_format = "0.0%"
    r += 1
    s.cell(r, 1, "Đối chiếu: tổng GT kỳ này và tổng GT lũy kế phải bằng dòng cộng cuối bảng 3. GTHT (cột L, M).").font = Font(name="Arial", size=9, italic=True)

    # Bảng 2: kịch bản
    r += 2
    s.cell(r, 1, "KỊCH BẢN ẢNH HƯỞNG LÊN F1").font = BOLD
    r += 1
    hdr(s, r, ["", "Kịch bản", "Mức rủi ro", "", "Số hạng mục", "GT kỳ này", "GT lũy kế",
               "F1 giảm nếu loại phần kỳ này", "F1 giảm nếu loại toàn bộ lũy kế", ""])
    r += 1
    scen = [
        ("Nguồn KL rủi ro Cao", "Cao"),
        ("Nguồn KL rủi ro Trung bình", "Trung bình"),
        ("Có diễn giải kích thước", "Thấp"),
    ]
    s0 = r
    for name, lvl in scen:
        s.cell(r, 2, name)
        s.cell(r, 3, lvl)
        s.cell(r, 5, f'=COUNTIFS({rng("P")},C{r})')
        s.cell(r, 6, f'=SUMIFS({rng("L")},{rng("P")},C{r})')
        s.cell(r, 7, f'=SUMIFS({rng("M")},{rng("P")},C{r})')
        s.cell(r, 8, f'=SUMIFS({rng("AA")},{rng("P")},C{r})')
        s.cell(r, 9, f'=SUMIFS({rng("AB")},{rng("P")},C{r})')
        r += 1
    s.cell(r, 2, "F1 theo hồ sơ (2. TH.!D38)")
    s.cell(r, 8, F1_VALUE)
    s.cell(r, 9, F1_VALUE)
    s.cell(r, 10, "Giá trị lấy từ file hồ sơ")
    f1row = r
    r += 1
    s.cell(r, 2, "F1 còn lại nếu loại phần rủi ro Cao chưa được chứng minh")
    s.cell(r, 8, f"=H{f1row}-H{s0}")
    s.cell(r, 9, f"=I{f1row}-I{s0}")
    s.cell(r, 10, "Âm nghĩa là tổng đã thanh toán vượt phần giá trị có căn cứ")
    for rr in range(s0, r + 1):
        for j in range(1, 11):
            s.cell(rr, j).font = BOLD if rr == r else BASE
        for j in (6, 7, 8, 9):
            s.cell(rr, j).number_format = NUMF
    s.cell(f1row, 8).font = BLUE
    s.cell(f1row, 9).font = BLUE

    # Bảng 3: nguồn KL kỳ trước
    r += 3
    s.cell(r, 1, "NGUỒN KHỐI LƯỢNG KỲ TRƯỚC (phải khớp hồ sơ Đợt 3 đã ký)").font = BOLD
    r += 1
    hdr(s, r, ["Phần", "Nguồn KL kỳ trước", "", "", "Số hạng mục", "GT kỳ trước", "", "", ""])
    r += 1
    k0 = r
    kt = sorted({(it["part"], KT_SRC[it["kt_src"]]) for it in items}, key=lambda k: (order[k[0]], k[1]))
    for part, src in kt:
        s.cell(r, 1, part)
        put(s, r, 2, src)
        s.cell(r, 5, f'=COUNTIFS({rng("B")},A{r},{rng("T")},B{r})')
        s.cell(r, 6, f'=SUMIFS({rng("K")},{rng("B")},A{r},{rng("T")},B{r})')
        r += 1
    s.cell(r, 1, "Tổng")
    s.cell(r, 5, f"=SUM(E{k0}:E{r - 1})")
    s.cell(r, 6, f"=SUM(F{k0}:F{r - 1})")
    for rr in range(k0, r + 1):
        for j in range(1, 10):
            s.cell(rr, j).font = BOLD if rr == r else BASE
        s.cell(rr, 6).number_format = NUMF

    # Bảng 4: KL vượt HĐ
    r += 3
    s.cell(r, 1, "KHỐI LƯỢNG DIỄN GIẢI VƯỢT KL HỢP ĐỒNG (bị công thức BBNT!H cắt về KL HĐ)").font = BOLD
    r += 1
    hdr(s, r, ["Phần", "", "", "", "Số hạng mục", "GT vượt HĐ", "", "", ""])
    r += 1
    v0 = r
    for part in ("XD",):
        s.cell(r, 1, part)
        s.cell(r, 5, f'=COUNTIFS({rng("B")},A{r},{rng("Y")},">0")')
        s.cell(r, 6, f'=SUMIFS({rng("Y")},{rng("B")},A{r})')
        for j in range(1, 10):
            s.cell(r, j).font = BASE
        s.cell(r, 6).number_format = NUMF
        r += 1
    s.cell(r, 1, ("Không cộng vào giá trị đề nghị. Nếu HĐ theo đơn giá, đây là phần có thể phát sinh; "
                  "nếu HĐ trọn gói, cần giải thích vì sao diễn giải khác BOQ.")).font = Font(name="Arial", size=9, italic=True)

    # Bảng 5: dạng đơn giá
    r += 3
    s.cell(r, 1, "DẠNG ĐƠN GIÁ (100% đơn giá là số nhập tay, cần đối chiếu BOQ)").font = BOLD
    r += 1
    hdr(s, r, ["Phần", "Dạng đơn giá", "", "", "Số hạng mục", "GT kỳ này", "GT lũy kế", "", ""])
    r += 1
    p0 = r
    pp = sorted({(it["part"], it["price_pattern"]) for it in items}, key=lambda k: (order[k[0]], k[1]))
    for part, pat in pp:
        s.cell(r, 1, part)
        s.cell(r, 2, pat)
        s.cell(r, 5, f'=COUNTIFS({rng("B")},A{r},{rng("G")},B{r})')
        s.cell(r, 6, f'=SUMIFS({rng("L")},{rng("B")},A{r},{rng("G")},B{r})')
        s.cell(r, 7, f'=SUMIFS({rng("M")},{rng("B")},A{r},{rng("G")},B{r})')
        for j in range(1, 10):
            s.cell(r, j).font = BASE
        for j in (6, 7):
            s.cell(r, j).number_format = NUMF
        r += 1
    s.freeze_panes = "A5"

    # ---------------- Nhap tay TH
    th = wb.create_sheet("Nhap tay TH")
    hdr(th, 1, ["Ô", "Dòng", "Loại", "Nội dung ô", "Giá trị đã lưu", "Ảnh hưởng / cần đối chiếu"],
        [14, 44, 30, 60, 18, 60])
    notes = {
        "I21": "Đợt 3 đã thanh toán: trừ trực tiếp 1:1 vào F1. Cần đề nghị TT đã duyệt + UNC.",
        "D17": "A.2.1 = 0: xác nhận PL05 không có phần xây dựng.",
    }
    for i, (cell, label, kind, content, val) in enumerate(th_rows, 2):
        coord = cell.split("!")[1]
        note = notes.get(coord, "")
        if not note and kind.startswith("Liên kết"):
            note = "Giá trị phụ thuộc file khác; cần file gốc hoặc đổi thành liên kết nội bộ."
        if not note and kind.startswith("Tham số"):
            note = "Đối chiếu tỷ lệ với điều khoản HĐ (tạm ứng, thu hồi, giữ lại, hóa đơn, VAT)."
        row = [cell, label[:80], kind, str(content)[:200], val, note]
        for j, v in enumerate(row, 1):
            cl = put(th, i, j, v)
            cl.font = BASE
            cl.alignment = Alignment(wrap_text=True, vertical="top")
        th.cell(i, 5).number_format = NUMF
    th.freeze_panes = "A2"

    # ---------------- Toan ven
    tv = wb.create_sheet("Toan ven file")
    hdr(tv, 1, ["Loại", "Vị trí", "Mô tả", "Hạng mục"], [22, 26, 70, 50])
    for i, row in enumerate(integ, 2):
        for j, v in enumerate(row, 1):
            put(tv, i, j, v).font = BASE
    tv.freeze_panes = "A2"
    if tree:
        write_lineage(wb, tree)
    wb.save(out_path)


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "tai-lieu-001.xlsx"
    out = sys.argv[2] if len(sys.argv) > 2 else "ket-qua-so-nhap-tay.xlsx"
    wb, wv, items, issues, _ = analyse(src)
    th_rows = th_inputs(wb, wv)
    integ = [(a, b, cc, dd) for a, b, cc, dd in issues] + integrity(wb, items)
    f1 = wv[CFG["TH"]]["D38"].value
    tree = lineage(wb, wv, items=items)
    write_report(items, th_rows, integ, src, out, f1, tree)
    # tóm tắt ra màn hình
    agg = collections.defaultdict(lambda: [0, 0.0, 0.0])
    for it in items:
        k = (it["part"], it["lk_src"])
        agg[k][0] += 1
        agg[k][1] += it["L"]
        agg[k][2] += it["M"]
    print(f"{len(items)} hạng mục")
    for k, v in sorted(agg.items()):
        print(f"{k[0]:3} {k[1]:18} {v[0]:4d}  kỳ này {v[1]:>18,.0f}  lũy kế {v[2]:>18,.0f}")
    print("Tổng kỳ này", f"{sum(i['L'] for i in items):,.0f}", "lũy kế", f"{sum(i['M'] for i in items):,.0f}")
    print("Đã ghi", out)


if __name__ == "__main__":
    main()
