# -*- coding: utf-8 -*-
"""
App kiểm tra hồ sơ thanh toán xây dựng - chạy trên máy.
    pip install -r requirements.txt
    streamlit run app.py
Tuỳ chọn local AI: cài Ollama, chạy `ollama pull qwen3.6:35b`, bật ở thanh bên.
"""
import io
from collections import Counter
from pathlib import Path

import pandas as pd
import streamlit as st

import core

st.set_page_config(page_title="Kiểm tra hồ sơ TTKL", layout="wide")
CHECK = core.load_checklist()
CODES = [""] + [r["ma"] for r in CHECK]
NAME = {r["ma"]: r["ten"] for r in CHECK}

# ------------------------------------------------------------- thanh bên
with st.sidebar:
    st.header("Thiết lập")
    du_an = st.text_input("Tên kho hồ sơ", "Hop_dong_1")
    kho = core.Kho(Path("kho_ho_so") / du_an)
    st.caption(f"File được lưu tại: `{kho.root.resolve()}`")
    st.divider()
    dung_ai = st.toggle("Dùng local AI (Ollama)", value=False)
    llm = None
    if dung_ai:
        model = st.text_input("Model", "qwen3.6:35b")
        host = st.text_input("Địa chỉ Ollama", "http://localhost:11434")
        llm = core.LocalLLM(model, host)
        if llm.ok():
            st.success("Đã kết nối Ollama")
        else:
            st.error("Không kết nối được Ollama. App vẫn chạy bằng từ khoá.")
    st.divider()
    st.caption("Mẹo: đặt tên file bắt đầu bằng mã danh mục, ví dụ `D01_BB giao nhan bom.pdf`, "
               "app sẽ phân loại chắc chắn.")

st.title("Kiểm tra hồ sơ thanh toán")
tab_ms, tab_up, tab_st, tab_qa, tab_hc = st.tabs(["0. Master hợp đồng", "1. Upload và phân loại",
                                                  "2. Tình trạng hồ sơ", "3. Hỏi đáp", "4. Quét số nhập tay"])

# ------------------------------------------------------------- tab 1
with tab_up:
    ups = st.file_uploader("Kéo thả hồ sơ (PDF, Word, Excel, ảnh)", accept_multiple_files=True,
                           type=["pdf", "docx", "xlsx", "xlsm", "png", "jpg", "jpeg", "tif", "tiff", "txt", "csv"])
    done = st.session_state.setdefault("done", set())
    if ups:
        new = [u for u in ups if (u.name, u.size) not in done]
        if new:
            bar = st.progress(0.0, "Đang đọc và phân loại…")
            for i, u in enumerate(new, 1):
                kho.add(u.name, u.getvalue(), CHECK, llm)
                done.add((u.name, u.size))
                bar.progress(i / len(new), f"{i}/{len(new)}: {u.name}")
            bar.empty()
            st.success(f"Đã thêm {len(new)} file.")

    if kho.index:
        df = pd.DataFrame([
            dict(file=fn, ma=i.get("ma", ""), ten_muc=NAME.get(i.get("ma", ""), "Không xác định"),
                 cach=i.get("cach", ""), goi_y=", ".join(i.get("goi_y", [])),
                 trang=i.get("trang"), ghi_chu=i.get("ghi_chu", ""), xac_nhan=i.get("xac_nhan", False))
            for fn, i in kho.index.items()])
        st.caption("Sửa cột **Mã** nếu app phân loại sai, rồi bấm Lưu. File đã sửa được coi là đã xác nhận.")
        ed = st.data_editor(
            df, hide_index=True, width="stretch", key="editor",
            column_config={
                "file": st.column_config.TextColumn("File", disabled=True),
                "ma": st.column_config.SelectboxColumn("Mã", options=CODES),
                "ten_muc": st.column_config.TextColumn("Tài liệu trong danh mục", disabled=True),
                "cach": st.column_config.TextColumn("Cách phân loại", disabled=True),
                "goi_y": st.column_config.TextColumn("Mã gợi ý khác", disabled=True),
                "trang": st.column_config.NumberColumn("Số trang", disabled=True),
                "ghi_chu": st.column_config.TextColumn("Ghi chú", disabled=True),
                "xac_nhan": st.column_config.CheckboxColumn("Đã xác nhận"),
            })
        c1, c2 = st.columns([1, 1])
        if c1.button("Lưu phân loại", type="primary"):
            for _, r in ed.iterrows():
                old = kho.index[r["file"]]
                if r["ma"] != old.get("ma") or bool(r["xac_nhan"]) != bool(old.get("xac_nhan")):
                    kho.set_code(r["file"], r["ma"])
                    kho.index[r["file"]]["xac_nhan"] = bool(r["xac_nhan"]) or r["ma"] != old.get("ma")
            kho.save()
            st.success("Đã lưu.")
            st.rerun()
        xoa = c2.selectbox("Xoá file khỏi kho", [""] + list(kho.index))
        if xoa and c2.button("Xoá"):
            kho.remove(xoa)
            st.rerun()
    else:
        st.info("Chưa có file nào. Upload hồ sơ để bắt đầu.")

# ------------------------------------------------------------- tab 2
with tab_st:
    rows = core.tinh_trang(CHECK, kho.index)
    cnt = Counter(r["tinh_trang"] for r in rows)
    bb = [r for r in rows if r["bat_buoc"]]
    c = st.columns(4)
    c[0].metric("Đã có", cnt.get("Đã có", 0))
    c[1].metric("Cần xác nhận", cnt.get("Cần xác nhận", 0))
    c[2].metric("Thiếu (bắt buộc)", cnt.get("Thiếu", 0))
    c[3].metric("Tài liệu bắt buộc đã đủ", f"{sum(r['tinh_trang'] == 'Đã có' for r in bb)}/{len(bb)}")

    nhom = st.multiselect("Lọc nhóm", sorted({r["nhom"] for r in rows}))
    chi_thieu = st.checkbox("Chỉ hiện mục thiếu / cần xác nhận")
    view = [r for r in rows if (not nhom or r["nhom"] in nhom)
            and (not chi_thieu or r["tinh_trang"] in ("Thiếu", "Cần xác nhận"))]
    sdf = pd.DataFrame(view)

    def color(v):
        return {"Thiếu": "background-color:#F8D7D3", "Cần xác nhận": "background-color:#FCEBCB",
                "Đã có": "background-color:#DDEFE3"}.get(v, "")

    if not sdf.empty:
        st.dataframe(sdf.style.map(color, subset=["tinh_trang"]), hide_index=True, width="stretch",
                     column_config={"ma": "Mã", "nhom": "Nhóm", "ten": "Tài liệu", "bat_buoc": "Bắt buộc",
                                    "uu_tien": "Ưu tiên", "tinh_trang": "Tình trạng", "so_file": "Số file",
                                    "file": "File", "muc_dich": "Để chứng minh", "mat_xich": "Ô / sheet liên quan"})
    md = core.xuat_request(rows)
    buf = io.BytesIO()
    pd.DataFrame([r for r in rows if r["tinh_trang"] in ("Thiếu", "Cần xác nhận")]).to_excel(buf, index=False)
    d1, d2 = st.columns(2)
    d1.download_button("Tải danh sách request (.md)", md, "danh_sach_request.md")
    d2.download_button("Tải danh sách request (.xlsx)", buf.getvalue(), "danh_sach_request.xlsx")

# ------------------------------------------------------------- tab 3
with tab_qa:
    st.caption("Câu hỏi về thiếu/đủ được trả lời từ bảng tình trạng. Câu hỏi về nội dung sẽ tìm trong các file "
               "đã upload; nếu bật local AI, AI sẽ tóm tắt câu trả lời kèm tên file nguồn.")
    goi_y = ["Hồ sơ còn thiếu gì?", "Nhóm E còn thiếu gì?", "E01 đã có chưa?",
             "Tỷ lệ tạm ứng theo hợp đồng là bao nhiêu?", "Tỷ lệ giữ lại bảo hành là bao nhiêu?"]
    cols = st.columns(len(goi_y))
    picked = None
    for col, g in zip(cols, goi_y):
        if col.button(g, width="stretch"):
            picked = g
    hist = st.session_state.setdefault("hist", [])
    q = st.chat_input("Hỏi về hồ sơ…") or picked
    if q:
        ans, hits = core.tra_loi(q, kho, CHECK, llm)
        hist.append((q, ans, hits))
    for q, ans, hits in hist:
        with st.chat_message("user"):
            st.write(q)
        with st.chat_message("assistant"):
            st.markdown(ans)
            if hits:
                with st.expander(f"{len(hits)} đoạn trích nguồn"):
                    for s, fn, ch in hits:
                        st.markdown(f"**{fn}** (điểm {s:.1f})")
                        st.text(ch[:900])

# ------------------------------------------------------------- tab 4
with tab_hc:
    st.caption("Chạy script quét số nhập tay trên file hồ sơ thanh toán Excel đã upload "
               "(cần có các sheet 2. TH., 3. GTHT, 4. BTHKLTH, 5. BBNT, 6. DG KL).")
    xl = [fn for fn, i in kho.index.items() if "3. GTHT" in (i.get("sheets") or [])]
    if not xl:
        st.info("Chưa có file Excel hồ sơ TTKL trong kho.")
    else:
        f = st.selectbox("File hồ sơ", xl)
        if st.button("Quét số nhập tay", type="primary"):
            import hardcode_audit as ha
            with st.spinner("Đang lần theo công thức…"):
                src = kho.files / f
                wb, wv, items, issues, _ = ha.analyse(src)
                th_rows = ha.th_inputs(wb, wv)
                integ = issues + ha.integrity(wb, items)
                tree = ha.lineage(wb, wv, items=items)
                out = io.BytesIO()
                ha.write_report(items, th_rows, integ, str(src), out, wv[ha.CFG["TH"]]["D38"].value, tree)
            st.subheader("Cây truy vết số đề nghị thanh toán")
            st.caption("Mỗi cấp thụt vào là thành phần tạo nên dòng phía trên. Nút lá là dữ liệu gốc cần chứng minh.")
            tdf = pd.DataFrame([dict(thanh_phan="\u2003" * r["level"] + str(r["text"]), o=r["cell"],
                                     cong_thuc=r["formula"], gia_tri=(f"{r['value']:,.2f}" if isinstance(r["value"], (int, float)) else str(r["value"] or "")), loai=r["kind"],
                                     tai_lieu=r["doc"], ghi_chu=r["note"]) for r in tree])
            st.dataframe(tdf, hide_index=True, width="stretch", height=420)
            st.subheader("Phân loại nguồn khối lượng")
            df = pd.DataFrame(items)
            df["muc_rui_ro"] = df["lk_src"].map(lambda s: ha.LK_SRC[s][1])
            df["mo_ta"] = df["lk_src"].map(lambda s: ha.LK_SRC[s][0])
            tong = (df.groupby(["part", "mo_ta", "muc_rui_ro"])
                    .agg(so_hang_muc=("gtht_row", "count"), gt_ky_nay=("L", "sum"), gt_luy_ke=("M", "sum"))
                    .reset_index())
            tong["_p"] = tong["part"].map({"XD": 0, "TB": 1, "PL": 2})
            tong["_r"] = tong["muc_rui_ro"].map({"Cao": 0, "Trung bình": 1, "Thấp": 2})
            tong = tong.sort_values(["_p", "_r"]).drop(columns=["_p", "_r"])
            st.dataframe(tong.style.format({"gt_ky_nay": "{:,.0f}", "gt_luy_ke": "{:,.0f}"}),
                         hide_index=True, width="stretch")
            st.download_button("Tải file kết quả chi tiết (.xlsx)", out.getvalue(), "ket-qua-so-nhap-tay.xlsx")


# ------------------------------------------------------------- tab 0: master hợp đồng
with tab_ms:
    import tempfile
    import master_hd as mh

    def _luu(up):
        d = Path(tempfile.mkdtemp())
        f = d / up.name
        f.write_bytes(up.getvalue())
        return f

    st.caption("Master file là nguồn gốc cho KL hợp đồng, đơn giá, tham số thanh toán, KL kỳ trước và "
               "lịch sử thanh toán. Mỗi hợp đồng một master. Hồ sơ đợt nào cũng đối chiếu vào master.")
    m1, m2, m3 = st.tabs(["Tạo master", "Đối chiếu hồ sơ đợt mới", "Nạp đợt đã duyệt"])
    with m1:
        hd = st.file_uploader("Hợp đồng (.doc, .docx, .pdf)", type=["doc", "docx", "pdf"], key="hd")
        bq = st.file_uploader("Bảng chi tiết giá trị HĐ / BOQ (.xlsx) - tuỳ chọn", type=["xlsx", "xlsm"], key="bq")
        if hd and st.button("Tạo master", type="primary"):
            with st.spinner("Đang đọc hợp đồng…"):
                paras = mh.doc_hop_dong(_luu(hd))
                rows = mh.trich_xuat(paras, mh.LocalLLM(model, host) if (dung_ai and llm and llm.ok()) else None)
                boq = mh.tim_bang_chi_tiet(__import__("openpyxl").load_workbook(_luu(bq), data_only=True)) if bq else None
                buf = io.BytesIO()
                mh.ghi_master(buf, hd.name, rows, boq, bq.name if bq else "")
            n = sum(1 for r in rows if r["gia_tri"] is not None)
            st.success(f"Trích được {n}/{len(rows)} tham số" + (f"; BOQ {sum(1 for i in boq['items'] if i['dvt'])} hạng mục" if boq else ""))
            st.dataframe(pd.DataFrame([dict(ma=r["ma"], tham_so=r["ten"], gia_tri=str(r["gia_tri"]) if r["gia_tri"] is not None else "(không tìm thấy)",
                                            cach=r["cach"], dieu=r["dieu"], ghi_chu=r["ghi_chu"]) for r in rows]),
                         hide_index=True, width="stretch")
            st.download_button("Tải master (.xlsx)", buf.getvalue(), f"master_{Path(hd.name).stem[:40]}.xlsx")
            st.info("Mở master, đọc cột 'Trích dẫn nguyên văn', sửa nếu sai và chọn 'Đã xác nhận'. "
                    "Đợt đầu: nhập KL kỳ trước và lịch sử thanh toán (nếu có).")
    with m2:
        ms = st.file_uploader("Master hợp đồng (.xlsx)", type=["xlsx"], key="ms2")
        hs = st.file_uploader("Hồ sơ thanh toán đợt mới (.xlsx)", type=["xlsx", "xlsm"], key="hs2")
        if ms and hs and st.button("Đối chiếu", type="primary"):
            with st.spinner("Đang đối chiếu…"):
                out = Path(tempfile.mkdtemp()) / "doi_chieu.xlsx"
                rows, tong, cnt = mh.doi_chieu(_luu(ms), _luu(hs), out)
            st.write("**Số tổng hợp tìm thấy trong hồ sơ**")
            st.dataframe(pd.DataFrame([dict(chi_tieu=k, gia_tri=f"{v['gia_tri']:,.0f}", o=v["o"], nhan=v["nhan"]) for k, v in tong.items()]),
                         hide_index=True, width="stretch")
            st.write("**Hạng mục lệch so với master**")
            lech = [dict(dong=x["dong"], ma=x["ma"], ten=x["ten"], lech="; ".join(x["loi"])) for x in rows if x["loi"]]
            st.dataframe(pd.DataFrame(lech) if lech else pd.DataFrame([{"ket_qua": "Không có hạng mục lệch"}]),
                         hide_index=True, width="stretch")
            st.download_button("Tải kết quả đối chiếu (.xlsx)", out.read_bytes(), f"doi_chieu_{Path(hs.name).stem[:40]}.xlsx")
    with m3:
        ms3 = st.file_uploader("Master hợp đồng (.xlsx)", type=["xlsx"], key="ms3")
        dd = st.file_uploader("Hồ sơ đợt đã được duyệt (.xlsx)", type=["xlsx", "xlsm"], key="dd3")
        so_dot = st.number_input("Số đợt", min_value=1, step=1, value=1)
        if ms3 and dd and st.button("Nạp vào master", type="primary"):
            f = _luu(ms3)
            ok, miss, tong = mh.nap_dot(f, _luu(dd), int(so_dot))
            st.success(f"Nạp {ok} hạng mục làm KL kỳ trước; {len(miss)} hạng mục không khớp BOQ.")
            if miss:
                st.warning("Không khớp: " + "; ".join(miss[:20]))
            st.download_button("Tải master đã cập nhật (.xlsx)", f.read_bytes(), ms3.name)
