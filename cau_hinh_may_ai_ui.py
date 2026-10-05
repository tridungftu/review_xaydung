# -*- coding: utf-8 -*-
"""
cau_hinh_may_ai_ui.py - Cửa sổ nhập / dò địa chỉ máy Local AI trong mạng LAN.

Mở bằng 14_cau_hinh_may_AI.bat. Đổi mạng, đổi IP máy AI: mở cửa sổ này, bấm "Tự tìm máy AI"
hoặc gõ IP mới, bấm "Kiểm tra kết nối", rồi "Lưu". Chỉ dùng thư viện chuẩn Python (tkinter).
"""
import queue
import threading
import tkinter as tk
from tkinter import ttk

import may_ai as A

NEN = "#F4F5F2"
CHU = "#1A1D21"
PHU = "#5A6068"
XANH = "#0B5E33"
DO = "#9B1C14"
VANG = "#7A4100"


class UngDung:
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.dang_chay = False
        self.da_kiem = None  # (host, model) đã kiểm tra thành công
        cfg = A.doc_cau_hinh()
        ip, cong = A.tach_host(cfg["may_ai"])

        root.title("Cấu hình máy Local AI")
        root.configure(bg=NEN)
        root.minsize(600, 560)
        st = ttk.Style(root)
        try:
            st.theme_use("vista" if "vista" in st.theme_names() else "clam")
        except tk.TclError:
            pass
        f_chu = ("Segoe UI", 10)
        st.configure(".", font=f_chu)
        st.configure("TFrame", background=NEN)
        st.configure("TLabel", background=NEN, foreground=CHU)
        st.configure("Phu.TLabel", foreground=PHU)
        st.configure("Tieu.TLabel", font=("Segoe UI", 15, "bold"))
        st.configure("Muc.TLabel", font=("Segoe UI", 10, "bold"))
        st.configure("TCheckbutton", background=NEN)
        st.configure("Chinh.TButton", font=("Segoe UI", 10, "bold"))

        k = ttk.Frame(root, padding=(20, 16, 20, 16))
        k.pack(fill="both", expand=True)
        k.columnconfigure(1, weight=1)

        ttk.Label(k, text="Máy Local AI", style="Tieu.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        ip_nay = ", ".join(A.ip_may_nay()) or "không xác định"
        ttk.Label(k, text=f"Máy này: {ip_nay}", style="Phu.TLabel").grid(row=1, column=0, columnspan=4, sticky="w")
        self.lb_dang_luu = ttk.Label(k, style="Phu.TLabel")
        self.lb_dang_luu.grid(row=2, column=0, columnspan=4, sticky="w", pady=(0, 12))
        self._hien_dang_luu(cfg)

        ttk.Label(k, text="Địa chỉ IP máy AI", style="Muc.TLabel").grid(row=3, column=0, sticky="w")
        self.v_ip = tk.StringVar(value=ip)
        self.e_ip = ttk.Entry(k, textvariable=self.v_ip, width=24, font=("Segoe UI", 11))
        self.e_ip.grid(row=3, column=1, sticky="we", padx=(10, 8), ipady=3)
        ttk.Label(k, text="Cổng").grid(row=3, column=2, sticky="e")
        self.v_cong = tk.StringVar(value=str(cong))
        ttk.Entry(k, textvariable=self.v_cong, width=7, font=("Segoe UI", 11)).grid(row=3, column=3, sticky="w",
                                                                                     padx=(6, 0), ipady=3)
        ttk.Label(k, text="Ví dụ 192.168.5.67. Xem IP trên máy AI bằng lệnh ipconfig (dòng IPv4 Address).",
                  style="Phu.TLabel", wraplength=430).grid(row=4, column=1, columnspan=3, sticky="w", padx=(10, 0), pady=(2, 8))

        hang = ttk.Frame(k)
        hang.grid(row=5, column=1, columnspan=3, sticky="w", padx=(10, 0), pady=(0, 12))
        self.b_kiem = ttk.Button(hang, text="Kiểm tra kết nối", command=self.kiem_tra)
        self.b_kiem.pack(side="left")
        self.b_tim = ttk.Button(hang, text="Tự tìm máy AI trong mạng", command=self.tim)
        self.b_tim.pack(side="left", padx=(8, 0))

        ttk.Label(k, text="Model", style="Muc.TLabel").grid(row=6, column=0, sticky="w")
        self.v_model = tk.StringVar(value=cfg["model"])
        self.cb_model = ttk.Combobox(k, textvariable=self.v_model, width=30, font=("Segoe UI", 11))
        self.cb_model.grid(row=6, column=1, columnspan=3, sticky="we", padx=(10, 0), ipady=2)
        ttk.Label(k, text="Danh sách model hiện ra sau khi kiểm tra kết nối. Nên dùng qwen3:8b.",
                  style="Phu.TLabel", wraplength=430).grid(row=7, column=1, columnspan=3, sticky="w", padx=(10, 0), pady=(2, 8))

        self.v_an = tk.BooleanVar(value=bool(cfg.get("an_danh", True)))
        ttk.Checkbutton(k, text="Che tên các bên, mã số thuế, số tài khoản trước khi gửi sang máy AI (nên bật)",
                        variable=self.v_an).grid(row=8, column=0, columnspan=4, sticky="w", pady=(0, 12))

        ttk.Label(k, text="Máy AI tìm thấy trong mạng (bấm đúp để chọn)", style="Muc.TLabel").grid(
            row=9, column=0, columnspan=4, sticky="w")
        self.ds = tk.Listbox(k, height=4, font=("Consolas", 10), activestyle="none", relief="solid", bd=1)
        self.ds.grid(row=10, column=0, columnspan=4, sticky="we", pady=(4, 12))
        self.ds.bind("<Double-Button-1>", self.chon_tu_ds)
        self.ds.bind("<Return>", self.chon_tu_ds)
        self.ds_ip = []

        ttk.Label(k, text="Kết quả", style="Muc.TLabel").grid(row=11, column=0, columnspan=4, sticky="w")
        self.kq = tk.Text(k, height=7, wrap="word", font=("Segoe UI", 10), relief="solid", bd=1, bg="#FFFFFF",
                          padx=8, pady=6, state="disabled")
        self.kq.grid(row=12, column=0, columnspan=4, sticky="nsew", pady=(4, 12))
        k.rowconfigure(12, weight=1)
        for tag, mau in (("ok", XANH), ("loi", DO), ("chu_y", VANG), ("phu", PHU)):
            self.kq.tag_configure(tag, foreground=mau)

        cuoi = ttk.Frame(k)
        cuoi.grid(row=13, column=0, columnspan=4, sticky="e")
        ttk.Button(cuoi, text="Đóng", command=root.destroy).pack(side="right")
        self.b_luu = ttk.Button(cuoi, text="Lưu", style="Chinh.TButton", command=self.luu)
        self.b_luu.pack(side="right", padx=(0, 8))

        self.e_ip.focus_set()
        self.e_ip.bind("<Return>", lambda e: self.kiem_tra())
        root.after(100, self._doc_q)
        self.ghi("Gõ IP máy AI rồi bấm “Kiểm tra kết nối”, hoặc bấm “Tự tìm máy AI trong mạng”.", "phu")

    # ---------------------------------------------------------- tiện ích
    def _hien_dang_luu(self, cfg):
        self.lb_dang_luu.configure(text=f"Đang lưu: {cfg['may_ai']} · model {cfg['model']}")

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
        self.root.configure(cursor="watch" if dang else "")

    def _host(self):
        ip, cong = self.v_ip.get().strip(), self.v_cong.get().strip() or str(A.CONG_MAC_DINH)
        if not ip or not A.ip_hop_le(ip):
            self.ghi("Địa chỉ IP chưa đúng dạng, ví dụ 192.168.5.67.", "loi", xoa=True)
            return None
        if not cong.isdigit():
            self.ghi("Cổng phải là số (thường là 11434).", "loi", xoa=True)
            return None
        return A.chuan_host(ip, int(cong))

    def _chay_nen(self, ham):
        self._ban(True)
        threading.Thread(target=lambda: self.q.put(("xong", ham())), daemon=True).start()

    def _doc_q(self):
        try:
            while True:
                loai, gt = self.q.get_nowait()
                if loai == "tien_do":
                    self.ghi(gt, "phu")
                elif loai == "xong":
                    self._ban(False)
                    if callable(gt):
                        gt()
        except queue.Empty:
            pass
        self.root.after(100, self._doc_q)

    # ---------------------------------------------------------- nút
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
                self.q.put(("tien_do", f"Kết nối được. Đang hỏi thử model {model} (lần đầu có thể mất 10–60 giây) …"))
                kq = A.kiem_tra(host, model)
            return lambda: self._sau_kiem_tra(host, model, kq)
        self._chay_nen(viec)

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
        elif not kq["co_model"] and kq["loi"]:
            self.ghi(f"✗ {kq['loi']}. Chọn model khác trong ô Model rồi bấm kiểm tra lại.", "loi")
        elif kq["ok"]:
            self.ghi(f"✓ Model {model} trả lời sau {kq['giay']} giây. Bấm “Lưu” để dùng.", "ok")
            self.da_kiem = (host, model)
        else:
            self.ghi("✗ " + kq["loi"], "loi")

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

        def tien_do(i, n):
            if i == n or i % 64 == 0:
                self.q.put(("tien_do", f"  đã dò {i}/{n} địa chỉ"))

        def viec():
            kq = A.tim_may_ai(cong, tien_do=tien_do)
            return lambda: self._sau_tim(kq, cong)
        self._chay_nen(viec)

    def _sau_tim(self, kq, cong):
        if not kq:
            self.ghi("✗ Không tìm thấy máy nào chạy Ollama trong mạng này.", "loi")
            self.ghi("Máy AI phải bật, Ollama đang chạy và cho phép kết nối từ mạng LAN "
                     "(OLLAMA_HOST=0.0.0.0, tường lửa mở cổng 11434).", "phu")
            return
        model = self.v_model.get().strip()
        for ip, models in kq:
            co = "✓ có " + model if model in models else "chưa có " + (model or "model đã chọn")
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
        if not sel:
            return
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
        self._hien_dang_luu(cfg)
        if self.da_kiem == (host, model):
            self.ghi(f"✓ Đã lưu: {host} · {model}. Các công cụ trích hợp đồng sẽ dùng địa chỉ này.", "ok")
        else:
            self.ghi(f"Đã lưu {host} · {model}, nhưng CHƯA kiểm tra thành công địa chỉ này. "
                     "Bấm “Kiểm tra kết nối” để chắc chắn.", "chu_y")


def main():
    try:  # chữ nét trên màn hình độ phân giải cao (Windows); phải gọi trước khi tạo cửa sổ
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = tk.Tk()
    UngDung(root)
    root.mainloop()


if __name__ == "__main__":
    main()
