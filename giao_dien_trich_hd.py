# -*- coding: utf-8 -*-
"""
giao_dien_trich_hd.py - MỘT cửa sổ cho việc trích tham số hợp đồng bằng Local AI trong mạng LAN.
(Gộp 5_trich_hop_dong_AI, 14_cau_hinh_may_AI, 15_trich_lo_hop_dong cũ.)

Tab "Trích hợp đồng": chọn / kéo thả một hoặc nhiều hợp đồng (.doc, .docx, .pdf), bấm Bắt đầu.
    Mỗi hợp đồng: trích bằng quy tắc, hỏi AI từng nhóm tham số (8 nhóm), so trích dẫn, ghi
    master_<HĐ>.xlsx, so_sanh_AI_<HĐ>.xlsx và TONG_HOP_<ngày giờ>.xlsx vào ket_qua_AI\\trich_lo\\.
Tab "Máy AI": gõ IP hoặc tự tìm máy AI trong mạng, kiểm tra kết nối, chọn model, lưu.

Mở bằng 5_trich_hop_dong_AI.bat (kéo thả file vào .bat cũng được: file được đưa sẵn vào danh sách).
Kéo thả vào cửa sổ cần thư viện tkinterdnd2 (file .bat tự cài); thiếu thì dùng nút Thêm file.
"""
import os
import queue
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

import may_ai as A

THU_MUC = Path(__file__).resolve().parent
KET_QUA = THU_MUC / "ket_qua_AI" / "trich_lo"
DUOI = (".doc", ".docx", ".pdf")

NEN = "#F4F5F2"
CHU = "#1A1D21"
PHU = "#5A6068"
XANH = "#0B5E33"
DO = "#9B1C14"
VANG = "#7A4100"
NHAN = "#1D4E89"

try:  # kéo thả file vào cửa sổ (tuỳ chọn)
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:
    TkinterDnD = None


def mo(path):
    """Mở file / thư mục bằng chương trình mặc định của Windows."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def phut_giay(s):
    s = int(s)
    return f"{s // 60} phút {s % 60:02d} giây" if s >= 60 else f"{s} giây"


# =====================================================================================================
class TabMayAI(ttk.Frame):
    """Nhập / tự tìm IP máy AI, kiểm tra kết nối, chọn model, lưu cau_hinh_may_ai.json."""

    def __init__(self, cha, app):
        super().__init__(cha, padding=(18, 14))
        self.app = app
        self.dang_chay = False
        self.da_kiem = None
        cfg = A.doc_cau_hinh()
        ip, cong = A.tach_host(cfg["may_ai"])
        self.columnconfigure(1, weight=1)

        ttk.Label(self, text=f"Máy này: {', '.join(A.ip_may_nay()) or 'không xác định'}", style="Phu.TLabel").grid(
            row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        ttk.Label(self, text="Địa chỉ IP máy AI", style="Muc.TLabel").grid(row=1, column=0, sticky="w")
        self.v_ip = tk.StringVar(value=ip)
        self.e_ip = ttk.Entry(self, textvariable=self.v_ip, width=24, font=("Segoe UI", 11))
        self.e_ip.grid(row=1, column=1, sticky="we", padx=(10, 8), ipady=3)
        ttk.Label(self, text="Cổng").grid(row=1, column=2, sticky="e")
        self.v_cong = tk.StringVar(value=str(cong))
        ttk.Entry(self, textvariable=self.v_cong, width=7, font=("Segoe UI", 11)).grid(row=1, column=3, sticky="w",
                                                                                        padx=(6, 0), ipady=3)
        ttk.Label(self, text="Ví dụ 192.168.5.67. Xem IP trên máy AI bằng lệnh ipconfig (dòng IPv4 Address).",
                  style="Phu.TLabel", wraplength=520).grid(row=2, column=1, columnspan=3, sticky="w", padx=(10, 0),
                                                           pady=(2, 8))
        hang = ttk.Frame(self)
        hang.grid(row=3, column=1, columnspan=3, sticky="w", padx=(10, 0), pady=(0, 12))
        self.b_kiem = ttk.Button(hang, text="Kiểm tra kết nối", command=self.kiem_tra)
        self.b_kiem.pack(side="left")
        self.b_tim = ttk.Button(hang, text="Tự tìm máy AI trong mạng", command=self.tim)
        self.b_tim.pack(side="left", padx=(8, 0))

        ttk.Label(self, text="Model", style="Muc.TLabel").grid(row=4, column=0, sticky="w")
        self.v_model = tk.StringVar(value=cfg["model"])
        self.cb_model = ttk.Combobox(self, textvariable=self.v_model, width=30, font=("Segoe UI", 11))
        self.cb_model.grid(row=4, column=1, columnspan=3, sticky="we", padx=(10, 0), ipady=2)
        ttk.Label(self, text="Danh sách model hiện ra sau khi kiểm tra kết nối. Nên dùng qwen3:8b.",
                  style="Phu.TLabel", wraplength=520).grid(row=5, column=1, columnspan=3, sticky="w", padx=(10, 0),
                                                           pady=(2, 8))
        self.v_an = tk.BooleanVar(value=bool(cfg.get("an_danh", True)))
        ttk.Checkbutton(self, text="Che tên các bên, mã số thuế, số tài khoản trước khi gửi sang máy AI (nên bật)",
                        variable=self.v_an).grid(row=6, column=0, columnspan=4, sticky="w", pady=(0, 12))

        ttk.Label(self, text="Máy AI tìm thấy trong mạng (bấm đúp để chọn)", style="Muc.TLabel").grid(
            row=7, column=0, columnspan=4, sticky="w")
        self.ds = tk.Listbox(self, height=4, font=("Consolas", 10), activestyle="none", relief="solid", bd=1)
        self.ds.grid(row=8, column=0, columnspan=4, sticky="we", pady=(4, 12))
        self.ds.bind("<Double-Button-1>", self.chon_tu_ds)
        self.ds_ip = []

        ttk.Label(self, text="Kết quả", style="Muc.TLabel").grid(row=9, column=0, columnspan=4, sticky="w")
        self.kq = tk.Text(self, height=7, wrap="word", font=("Segoe UI", 10), relief="solid", bd=1, bg="#FFFFFF",
                          padx=8, pady=6, state="disabled")
        self.kq.grid(row=10, column=0, columnspan=4, sticky="nsew", pady=(4, 12))
        self.rowconfigure(10, weight=1)
        for tag, mau in (("ok", XANH), ("loi", DO), ("chu_y", VANG), ("phu", PHU)):
            self.kq.tag_configure(tag, foreground=mau)
        self.b_luu = ttk.Button(self, text="Lưu", style="Chinh.TButton", command=self.luu)
        self.b_luu.grid(row=11, column=3, sticky="e")
        self.e_ip.bind("<Return>", lambda e: self.kiem_tra())
        self.ghi("Gõ IP máy AI rồi bấm “Kiểm tra kết nối”, hoặc bấm “Tự tìm máy AI trong mạng”.", "phu")

    def ghi(self, msg, tag=None, xoa=False):
        self.kq.configure(state="normal")
        if xoa:
            self.kq.delete("1.0", "end")
        self.kq.insert("end", msg + "\n", tag or ())
        self.kq.see("end")
        self.kq.configure(state="disabled")

    def _ban(self, dang):
        self.dang_chay = dang
        for b in (self.b_kiem, self.b_tim, self.b_luu):
            b.state(["disabled"] if dang else ["!disabled"])

    def _host(self):
        ip, cong = self.v_ip.get().strip(), self.v_cong.get().strip() or str(A.CONG_MAC_DINH)
        if not ip or not A.ip_hop_le(ip):
            self.ghi("Địa chỉ IP chưa đúng dạng, ví dụ 192.168.5.67.", "loi", xoa=True)
            return None
        if not cong.isdigit():
            self.ghi("Cổng phải là số (thường là 11434).", "loi", xoa=True)
            return None
        return A.chuan_host(ip, int(cong))

    def _nen(self, viec):
        self._ban(True)
        threading.Thread(target=lambda: self.app.q.put(("goi", self._xong(viec()))), daemon=True).start()

    def _xong(self, ham):
        def f():
            self._ban(False)
            ham()
        return f

    def kiem_tra(self):
        if self.dang_chay:
            return
        host = self._host()
        if not host:
            return
        model = self.v_model.get().strip()
        self.ghi(f"Đang kiểm tra {host} …", "phu", xoa=True)

        def viec():
            kq = A.kiem_tra(host, None)
            if kq["ok"] and model and (model in kq["models"] or f"{model}:latest" in kq["models"]):
                self.app.q.put(("goi", lambda: self.ghi(
                    f"Kết nối được. Đang hỏi thử model {model} (lần đầu có thể mất 10–60 giây) …", "phu")))
                kq = A.kiem_tra(host, model)
            return lambda: self._sau_kiem_tra(host, model, kq)
        self._nen(viec)

    def _sau_kiem_tra(self, host, model, kq):
        if kq["models"]:
            self.cb_model.configure(values=[m for m in kq["models"] if "embed" not in m])
        if not kq["models"]:
            self.ghi("✗ " + kq["loi"], "loi")
            self.ghi("Kiểm tra: máy AI đã bật và Ollama đang chạy? IP đúng chưa (có thể đã đổi)? "
                     "Thử bấm “Tự tìm máy AI trong mạng”.", "phu")
            self.da_kiem = None
            return
        self.ghi(f"✓ Kết nối được {host}", "ok")
        self.ghi("Model trên máy AI: " + ", ".join(kq["models"]), "phu")
        if not model:
            self.ghi("Chọn một model trong ô Model rồi bấm kiểm tra lại.", "chu_y")
        elif kq["ok"]:
            self.ghi(f"✓ Model {model} trả lời sau {kq['giay']} giây. Bấm “Lưu” để dùng.", "ok")
            self.da_kiem = (host, model)
        else:
            self.ghi(f"✗ {kq['loi']}. Chọn model khác trong ô Model rồi bấm kiểm tra lại.", "loi")

    def tim(self):
        if self.dang_chay:
            return
        try:
            cong = int(self.v_cong.get().strip() or A.CONG_MAC_DINH)
        except ValueError:
            cong = A.CONG_MAC_DINH
        mang = ", ".join(ip.rsplit(".", 1)[0] + ".x" for ip in A.ip_may_nay()) or "?"
        self.ghi(f"Đang dò các máy trong mạng {mang} có mở cổng {cong} (khoảng 5–15 giây) …", "phu", xoa=True)
        self.ds.delete(0, "end")
        self.ds_ip = []

        def viec():
            kq = A.tim_may_ai(cong)
            return lambda: self._sau_tim(kq, cong)
        self._nen(viec)

    def _sau_tim(self, kq, cong):
        if not kq:
            self.ghi("✗ Không tìm thấy máy nào chạy Ollama trong mạng này.", "loi")
            self.ghi("Máy AI phải bật, Ollama đang chạy và cho phép kết nối từ mạng LAN "
                     "(OLLAMA_HOST=0.0.0.0, tường lửa mở cổng 11434).", "phu")
            return
        model = self.v_model.get().strip()
        for ip, models in kq:
            co = "✓ có " + model if model in models else "chưa có " + (model or "model")
            self.ds.insert("end", f"{ip:16} {co:22} | {', '.join(models)}")
            self.ds_ip.append((ip, models))
        self.ghi(f"✓ Tìm thấy {len(kq)} máy AI. Bấm đúp vào một dòng để chọn.", "ok")
        hop = [x for x in kq if model in x[1]]
        if len(hop) == 1:
            self.v_ip.set(hop[0][0])
            self.v_cong.set(str(cong))
            self.ghi(f"Đã điền sẵn {hop[0][0]} (máy duy nhất có {model}). Bấm “Kiểm tra kết nối” rồi “Lưu”.", "chu_y")

    def chon_tu_ds(self, _e=None):
        sel = self.ds.curselection()
        if sel:
            ip, models = self.ds_ip[sel[0]]
            self.v_ip.set(ip)
            self.cb_model.configure(values=[m for m in models if "embed" not in m])
            self.kiem_tra()

    def luu(self):
        host = self._host()
        if not host:
            return
        model = self.v_model.get().strip()
        if not model:
            self.ghi("Chưa chọn model.", "loi")
            return
        cfg = A.doc_cau_hinh()
        cfg.update(may_ai=host, model=model, an_danh=bool(self.v_an.get()))
        A.luu_cau_hinh(cfg)
        self.app.cap_nhat_trang_thai_ai()
        if self.da_kiem == (host, model):
            self.ghi(f"✓ Đã lưu: {host} · {model}.", "ok")
        else:
            self.ghi(f"Đã lưu {host} · {model}, nhưng CHƯA kiểm tra thành công địa chỉ này. "
                     "Bấm “Kiểm tra kết nối” để chắc chắn.", "chu_y")


# =====================================================================================================
class TabTrich(ttk.Frame):
    """Chọn hợp đồng, chạy trích (quy tắc + AI theo nhóm), theo dõi tiến độ, mở kết quả."""

    def __init__(self, cha, app, file_ban_dau):
        super().__init__(cha, padding=(18, 14))
        self.app = app
        self.files = []
        self.dung = None
        self.dang_chay = False
        self.t0 = None
        self.tong_hop = None
        self.columnconfigure(0, weight=1)

        # --- danh sách hợp đồng
        dau = ttk.Frame(self)
        dau.grid(row=0, column=0, sticky="we")
        ttk.Label(dau, text="Hợp đồng cần trích", style="Muc.TLabel").pack(side="left")
        ttk.Label(dau, text=("  – kéo thả file .doc/.docx/.pdf vào danh sách, hoặc bấm Thêm file"
                             if TkinterDnD else "  – bấm Thêm file / Thêm thư mục"), style="Phu.TLabel").pack(side="left")

        khung = ttk.Frame(self)
        khung.grid(row=1, column=0, sticky="nsew", pady=(6, 6))
        khung.columnconfigure(0, weight=1)
        khung.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.tv = ttk.Treeview(khung, columns=("stt", "file", "tt"), show="headings", height=7, selectmode="extended")
        for c, t, w, st in (("stt", "#", 40, False), ("file", "Hợp đồng", 520, True), ("tt", "Trạng thái", 220, False)):
            self.tv.heading(c, text=t, anchor="w")
            self.tv.column(c, width=w, stretch=st, anchor="w")
        self.tv.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(khung, orient="vertical", command=self.tv.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.tv.configure(yscrollcommand=sb.set)
        self.tv.tag_configure("loi", foreground=DO)
        self.tv.tag_configure("xong", foreground=XANH)
        self.tv.tag_configure("chay", foreground=NHAN)
        if TkinterDnD:
            self.tv.drop_target_register(DND_FILES)
            self.tv.dnd_bind("<<Drop>>", self._tha)

        nut = ttk.Frame(self)
        nut.grid(row=2, column=0, sticky="w", pady=(0, 10))
        self.b_them = ttk.Button(nut, text="Thêm file…", command=self.them_file)
        self.b_them.pack(side="left")
        self.b_thu_muc = ttk.Button(nut, text="Thêm thư mục…", command=self.them_thu_muc)
        self.b_thu_muc.pack(side="left", padx=(8, 0))
        self.b_bo = ttk.Button(nut, text="Bỏ file đã chọn", command=self.bo_chon)
        self.b_bo.pack(side="left", padx=(8, 0))
        self.b_xoa = ttk.Button(nut, text="Xoá hết", command=self.xoa_het)
        self.b_xoa.pack(side="left", padx=(8, 0))

        # --- tuỳ chọn + chạy
        tc = ttk.Frame(self)
        tc.grid(row=3, column=0, sticky="we", pady=(0, 8))
        self.v_lam_lai = tk.BooleanVar(value=False)
        self.v_khong_ai = tk.BooleanVar(value=False)
        ttk.Checkbutton(tc, text="Hỏi lại AI kể cả hợp đồng đã có kết quả", variable=self.v_lam_lai).pack(side="left")
        ttk.Checkbutton(tc, text="Chỉ dùng quy tắc (không hỏi AI)", variable=self.v_khong_ai).pack(side="left",
                                                                                                     padx=(16, 0))
        chay = ttk.Frame(self)
        chay.grid(row=4, column=0, sticky="we", pady=(0, 6))
        self.b_bat_dau = ttk.Button(chay, text="Bắt đầu trích", style="Chinh.TButton", command=self.bat_dau)
        self.b_bat_dau.pack(side="left")
        self.b_dung = ttk.Button(chay, text="Dừng", command=self.dung_lai)
        self.b_dung.pack(side="left", padx=(8, 0))
        self.b_dung.state(["disabled"])
        self.b_mo_th = ttk.Button(chay, text="Mở file tổng hợp", command=lambda: self.tong_hop and mo(self.tong_hop))
        self.b_mo_th.pack(side="right")
        self.b_mo_th.state(["disabled"])
        ttk.Button(chay, text="Mở thư mục kết quả",
                   command=lambda: (KET_QUA.mkdir(parents=True, exist_ok=True), mo(KET_QUA))).pack(side="right",
                                                                                                padx=(0, 8))
        self.pb = ttk.Progressbar(self, mode="determinate")
        self.pb.grid(row=5, column=0, sticky="we")
        self.lb_tien_do = ttk.Label(self, text="Chưa chạy.", style="Phu.TLabel")
        self.lb_tien_do.grid(row=6, column=0, sticky="w", pady=(4, 6))

        ttk.Label(self, text="Nhật ký", style="Muc.TLabel").grid(row=7, column=0, sticky="w")
        self.log = tk.Text(self, height=12, wrap="word", font=("Consolas", 9), relief="solid", bd=1, bg="#FFFFFF",
                           padx=8, pady=6, state="disabled")
        self.log.grid(row=8, column=0, sticky="nsew", pady=(4, 0))
        self.rowconfigure(8, weight=2)
        for tag, mau in (("ok", XANH), ("loi", DO), ("chu_y", VANG), ("phu", PHU)):
            self.log.tag_configure(tag, foreground=mau)

        self.them(file_ban_dau)
        self.ghi("Thêm hợp đồng rồi bấm “Bắt đầu trích”. Mỗi hợp đồng hỏi AI 8 nhóm tham số; máy AI hiện tại "
                 "mất khoảng 5–8 phút / hợp đồng. Có thể để chạy và làm việc khác.", "phu")

    # ------------------------------------------------ danh sách
    def them(self, paths):
        moi = 0
        for p in paths:
            p = Path(p)
            ds = [f for f in sorted(p.rglob("*")) if f.suffix.lower() in DUOI and not f.name.startswith("~$")] \
                if p.is_dir() else ([p] if p.suffix.lower() in DUOI and p.exists() else [])
            if not ds and not p.is_dir():
                self.ghi(f"Bỏ qua (không phải hợp đồng .doc/.docx/.pdf): {p.name}", "chu_y")
            for f in ds:
                if all(str(f.resolve()).lower() != str(x.resolve()).lower() for x in self.files):
                    self.files.append(f)
                    moi += 1
        self._ve_ds()
        return moi

    def _ve_ds(self):
        self.tv.delete(*self.tv.get_children())
        for i, f in enumerate(self.files):
            self.tv.insert("", "end", iid=str(i), values=(i + 1, f.name, "Chờ"))
        self.lb_tien_do.configure(text=f"{len(self.files)} hợp đồng trong danh sách." if self.files else "Chưa có hợp đồng.")

    def _tha(self, e):
        self.them(self.tk.splitlist(e.data))

    def them_file(self):
        fs = filedialog.askopenfilenames(title="Chọn hợp đồng", filetypes=[("Hợp đồng", "*.doc *.docx *.pdf"),
                                                                          ("Tất cả", "*.*")])
        self.them(fs)

    def them_thu_muc(self):
        d = filedialog.askdirectory(title="Chọn thư mục chứa hợp đồng")
        if d:
            n = self.them([d])
            self.ghi(f"Thêm {n} hợp đồng từ thư mục {d}", "phu")

    def bo_chon(self):
        bo = {int(i) for i in self.tv.selection()}
        self.files = [f for i, f in enumerate(self.files) if i not in bo]
        self._ve_ds()

    def xoa_het(self):
        self.files = []
        self._ve_ds()

    # ------------------------------------------------ chạy
    def ghi(self, msg, tag=None):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag or ())
        self.log.see("end")
        self.log.configure(state="disabled")

    def _ban(self, dang):
        self.dang_chay = dang
        for b in (self.b_them, self.b_thu_muc, self.b_bo, self.b_xoa, self.b_bat_dau):
            b.state(["disabled"] if dang else ["!disabled"])
        self.b_dung.state(["!disabled"] if dang else ["disabled"])
        self.app.khoa_tab_ai(dang)

    def bat_dau(self):
        if self.dang_chay:
            return
        if not self.files:
            self.ghi("Chưa có hợp đồng nào trong danh sách.", "loi")
            return
        import trich_lo_hop_dong as T
        self._ve_ds()
        self.dung = threading.Event()
        self.tong_hop = None
        self.b_mo_th.state(["disabled"])
        self.pb.configure(maximum=len(self.files) * 8, value=0)
        self.t0 = time.time()
        self.hd_i, self.nhom_i = 0, 0
        self._ban(True)
        self.ghi(f"=== Bắt đầu {time.strftime('%H:%M')} – {len(self.files)} hợp đồng ===", "phu")
        q = self.app.q

        def bao(msg):
            q.put(("log", msg))

        def tien_do(i, tt):
            q.put(("tt", i, tt))

        def viec():
            try:
                out = T.chay([str(f) for f in self.files], KET_QUA, self.v_lam_lai.get(), self.v_khong_ai.get(),
                             bao=bao, dung=self.dung, tien_do=tien_do)
                q.put(("xong", str(out)))
            except T.LoiChay as e:
                q.put(("loi_chay", str(e)))
            except Exception as e:
                q.put(("loi_chay", f"{e.__class__.__name__}: {e}"))
        threading.Thread(target=viec, daemon=True).start()
        self._dong_ho()

    def _dong_ho(self):
        if self.dang_chay:
            n = len(self.files)
            self.lb_tien_do.configure(text=f"Hợp đồng {min(self.hd_i + 1, n)}/{n} · nhóm {self.nhom_i}/8 · "
                                           f"đã chạy {phut_giay(time.time() - self.t0)}")
            self.after(1000, self._dong_ho)

    def dung_lai(self):
        if self.dung is not None:
            self.dung.set()
            self.ghi("Đang dừng… (chờ máy AI trả lời xong nhóm đang hỏi)", "chu_y")
            self.b_dung.state(["disabled"])

    def nhan(self, loai, *gt):
        if loai == "log":
            msg = gt[0]
            m = re.search(r"nhóm (\d+)/(\d+)", msg)
            if m:
                self.nhom_i = int(m.group(1))
                self.pb.configure(value=self.hd_i * 8 + self.nhom_i - 1)
            tag = "loi" if "LỖI" in msg else ("chu_y" if "CẢNH BÁO" in msg or "thử lại" in msg
                                              else ("ok" if msg.strip().startswith("xong") else None))
            self.ghi(msg.rstrip(), tag)
        elif loai == "tt":
            i, tt = gt
            self.hd_i = i if tt == "Đang chạy" else i + 1
            if tt == "Đang chạy":
                self.nhom_i = 0
            self.pb.configure(value=self.hd_i * 8)
            if self.tv.exists(str(i)):
                self.tv.set(str(i), "tt", tt)
                self.tv.item(str(i), tags=("loi",) if tt.startswith(("Lỗi", "Đã dừng")) else
                             (("chay",) if tt == "Đang chạy" else ("xong",)))
        elif loai == "xong":
            self.tong_hop = gt[0]
            self._ban(False)
            self.b_mo_th.state(["!disabled"])
            self.pb.configure(value=self.pb["maximum"])
            self.lb_tien_do.configure(text=f"Xong sau {phut_giay(time.time() - self.t0)}. "
                                           f"Bấm “Mở file tổng hợp”, xem sheet Can xac nhan trước.")
        elif loai == "loi_chay":
            self._ban(False)
            self.ghi("DỪNG: " + gt[0], "loi")
            self.lb_tien_do.configure(text="Dừng – xem nhật ký.")
            if "máy AI" in gt[0]:
                self.app.nb.select(1)


# =====================================================================================================
class UngDung:
    def __init__(self, root, file_ban_dau, tab=0):
        self.root = root
        self.q = queue.Queue()
        root.title("Trích tham số hợp đồng – Local AI")
        root.configure(bg=NEN)
        root.minsize(820, 660)
        st = ttk.Style(root)
        try:
            st.theme_use("vista" if "vista" in st.theme_names() else "clam")
        except tk.TclError:
            pass
        st.configure(".", font=("Segoe UI", 10))
        for w in ("TFrame", "TLabel", "TCheckbutton", "TNotebook"):
            st.configure(w, background=NEN)
        st.configure("TLabel", foreground=CHU)
        st.configure("Phu.TLabel", foreground=PHU)
        st.configure("Tieu.TLabel", font=("Segoe UI", 15, "bold"))
        st.configure("Muc.TLabel", font=("Segoe UI", 10, "bold"))
        st.configure("Chinh.TButton", font=("Segoe UI", 10, "bold"))
        st.configure("TNotebook.Tab", padding=(14, 6), font=("Segoe UI", 10))

        dau = ttk.Frame(root, padding=(18, 12, 18, 4))
        dau.pack(fill="x")
        ttk.Label(dau, text="Trích tham số hợp đồng", style="Tieu.TLabel").pack(side="left")
        self.lb_ai = ttk.Label(dau, style="Phu.TLabel")
        self.lb_ai.pack(side="right")

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True, padx=12, pady=(4, 12))
        self.tab_trich = TabTrich(self.nb, self, file_ban_dau)
        self.tab_ai = TabMayAI(self.nb, self)
        self.nb.add(self.tab_trich, text="Trích hợp đồng")
        self.nb.add(self.tab_ai, text="Máy AI")
        self.nb.select(tab)
        self.cap_nhat_trang_thai_ai()
        root.after(100, self._doc_q)
        root.protocol("WM_DELETE_WINDOW", self.dong)

    def khoa_tab_ai(self, khoa):
        self.nb.tab(1, state="disabled" if khoa else "normal")

    def cap_nhat_trang_thai_ai(self):
        cfg = A.doc_cau_hinh()
        self.lb_ai.configure(text=f"Máy AI: {cfg['may_ai']} · {cfg['model']} · đang kiểm tra…", foreground=PHU)

        def viec():
            kq = A.kiem_tra(cfg["may_ai"], cfg["model"], thu_model=False)
            return lambda: self.lb_ai.configure(
                text=f"Máy AI: {cfg['may_ai']} · {cfg['model']} · " + ("● kết nối được" if kq["ok"] else "● " + kq["loi"]),
                foreground=XANH if kq["ok"] else DO)
        threading.Thread(target=lambda: self.q.put(("goi", viec())), daemon=True).start()

    def _doc_q(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == "goi":
                    m[1]()
                else:
                    self.tab_trich.nhan(*m)
        except queue.Empty:
            pass
        self.root.after(150, self._doc_q)

    def dong(self):
        if self.tab_trich.dang_chay:
            from tkinter import messagebox
            if not messagebox.askyesno("Đang chạy", "Đang trích hợp đồng. Đóng cửa sổ sẽ dừng giữa chừng "
                                                    "(hợp đồng đã xong vẫn giữ kết quả). Đóng?"):
                return
        self.root.destroy()


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    tab = 1 if "--may-ai" in sys.argv else 0
    try:  # chữ nét trên màn hình độ phân giải cao (Windows); gọi trước khi tạo cửa sổ
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = TkinterDnD.Tk() if TkinterDnD else tk.Tk()

    def bao_loi(*exc):  # mở bằng pythonw không có cửa sổ đen: lỗi phải hiện ra và ghi lại
        import traceback
        from tkinter import messagebox
        txt = "".join(traceback.format_exception(*exc))
        (THU_MUC / "loi_giao_dien.txt").write_text(txt, encoding="utf-8")
        messagebox.showerror("Lỗi", f"{exc[1]}\n\nChi tiết đã ghi vào loi_giao_dien.txt (gửi file này để sửa).")
    root.report_callback_exception = bao_loi
    try:
        UngDung(root, args, tab)
    except Exception:
        bao_loi(*sys.exc_info())
        root.destroy()
        return
    root.mainloop()


if __name__ == "__main__":
    main()
