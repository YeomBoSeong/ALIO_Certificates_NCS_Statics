"""
ALIO 공기업 채용공고 수집기 - GUI
opendata.alio.go.kr 공개 API로 채용공고를 수집하고
NCS 및 자격증 정보를 추출해 Excel 파일로 저장합니다.

수집/분석 핵심 로직은 collector_core.py에 있습니다 (GUI 없는 환경에서도 재사용 가능).
"""
import os
import sys
import threading
import queue
from pathlib import Path
from typing import Optional

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext

from collector_core import (
    HAS_REQUESTS, HAS_PDF, HAS_HWP, HAS_EXCEL,
    ALIO_JOB_TYPES, ALIO_DEFAULT_ON, NCS_CATEGORIES,
    load_config, save_config, collect_worker,
)

# PyInstaller(onefile)로 빌드된 exe에서는 __file__이 실행 시 압축이 풀리는
# 임시 폴더(_MEIPASS)를 가리켜서, 프로그램 종료 시 그 폴더가 삭제되며 기본
# 저장 위치도 함께 사라진다. exe로 실행 중일 때는 실제 exe가 있는 위치를
# 기준으로 잡아야 한다.
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent

DEFAULT_OUTPUT_DIR = APP_DIR / "ALIO_Downloads"

# ════════════════════════════════════════════════════════════════
# GUI
# ════════════════════════════════════════════════════════════════

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ALIO 공기업 채용공고 수집기")
        self.geometry("800x660")
        self.minsize(700, 560)
        self.configure(bg="#f5f5f5")

        self._cfg        = load_config()
        self._stop_ev    = threading.Event()
        self._worker: Optional[threading.Thread] = None
        self._msg_q: queue.Queue = queue.Queue()
        self._last_excel = ""

        self._build_ui()
        self._load_settings()
        self._check_deps()
        self.after(100, self._poll_queue)

    # ── 패키지 확인 ──────────────────────────────────────────────
    def _check_deps(self):
        missing = []
        if not HAS_REQUESTS: missing.append("requests")
        if not HAS_PDF:      missing.append("pdfplumber")
        if not HAS_HWP:      missing.append("olefile")
        if not HAS_EXCEL:    missing.append("pandas, openpyxl")
        if missing:
            warn = "일부 패키지가 없습니다:\n  " + "\n  ".join(missing)
            warn += "\n\nsetup.bat을 실행하여 설치하세요."
            messagebox.showwarning("패키지 누락", warn)

    # ── UI 구성 ──────────────────────────────────────────────────
    def _build_ui(self):
        # 헤더
        hdr = tk.Frame(self, bg="#1a237e", height=56)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="ALIO 공기업 채용공고 수집기",
                 font=("맑은 고딕", 15, "bold"), bg="#1a237e", fg="white").pack(side="left", padx=16, pady=14)
        tk.Label(hdr, text="opendata.alio.go.kr",
                 font=("맑은 고딕", 9), bg="#1a237e", fg="#90caf9").pack(side="right", padx=16)

        # 탭
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=("맑은 고딕", 10), padding=[12, 4])

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=10, pady=8)

        self.tab_set  = ttk.Frame(self.nb)
        self.tab_prog = ttk.Frame(self.nb)
        self.nb.add(self.tab_set,  text="  설정  ")
        self.nb.add(self.tab_prog, text="  수집 진행  ")

        self._build_settings_tab()
        self._build_progress_tab()

    # ── 설정 탭 ──────────────────────────────────────────────────
    def _build_settings_tab(self):
        tab = self.tab_set
        tab.columnconfigure(0, weight=1)

        canvas = tk.Canvas(tab, bg="#f5f5f5", highlightthickness=0)
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=canvas.yview)
        self._settings_frame = ttk.Frame(canvas)

        self._settings_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=self._settings_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        frm = self._settings_frame
        frm.columnconfigure(0, weight=1)
        row = 0

        # ── 키워드 선택 ──────────────────────────────────────────
        kw_lf = ttk.LabelFrame(frm, text="  검색 키워드  ", padding=(10, 6))
        kw_lf.grid(row=row, column=0, sticky="ew", padx=12, pady=(10, 4))
        kw_lf.columnconfigure(0, weight=1)
        row += 1

        kw_nb = ttk.Notebook(kw_lf)
        kw_nb.pack(fill="both", expand=True)

        tab_alio   = ttk.Frame(kw_nb, padding=8)
        tab_ncs    = ttk.Frame(kw_nb, padding=8)
        tab_custom = ttk.Frame(kw_nb, padding=8)
        kw_nb.add(tab_alio,   text="  ALIO 직종  ")
        kw_nb.add(tab_ncs,    text="  NCS 분류  ")
        kw_nb.add(tab_custom, text="  직접 입력  ")

        # ── ALIO 직종 탭 ─────────────────────────────────────────
        self._alio_vars: dict = {}
        cols = 5
        for i, kw in enumerate(ALIO_JOB_TYPES):
            var = tk.BooleanVar(value=(kw in ALIO_DEFAULT_ON))
            self._alio_vars[kw] = var
            ttk.Checkbutton(tab_alio, text=kw, variable=var).grid(
                row=i // cols, column=i % cols, sticky="w", padx=8, pady=2)

        btn_row_alio = ttk.Frame(tab_alio)
        btn_row_alio.grid(row=(len(ALIO_JOB_TYPES) // cols) + 1,
                          column=0, columnspan=cols, sticky="w", pady=(6, 0))
        ttk.Button(btn_row_alio, text="전체 선택",
                   command=lambda: [v.set(True)  for v in self._alio_vars.values()]).pack(side="left", padx=4)
        ttk.Button(btn_row_alio, text="전체 해제",
                   command=lambda: [v.set(False) for v in self._alio_vars.values()]).pack(side="left", padx=4)

        # ── NCS 분류 탭 ──────────────────────────────────────────
        self._ncs_vars: dict = {}
        cols_ncs = 3
        for i, kw in enumerate(NCS_CATEGORIES):
            var = tk.BooleanVar(value=False)
            self._ncs_vars[kw] = var
            ttk.Checkbutton(tab_ncs, text=kw, variable=var).grid(
                row=i // cols_ncs, column=i % cols_ncs, sticky="w", padx=8, pady=2)

        btn_row_ncs = ttk.Frame(tab_ncs)
        btn_row_ncs.grid(row=(len(NCS_CATEGORIES) // cols_ncs) + 1,
                         column=0, columnspan=cols_ncs, sticky="w", pady=(6, 0))
        ttk.Button(btn_row_ncs, text="전체 선택",
                   command=lambda: [v.set(True)  for v in self._ncs_vars.values()]).pack(side="left", padx=4)
        ttk.Button(btn_row_ncs, text="전체 해제",
                   command=lambda: [v.set(False) for v in self._ncs_vars.values()]).pack(side="left", padx=4)

        # ── 직접 입력 탭 ─────────────────────────────────────────
        tk.Label(tab_custom,
                 text="※ ALIO 직종 / NCS 분류에서 복수 선택 시 키워드 간에는 기본적으로 OR 검색입니다.",
                 fg="#1565c0", font=("맑은 고딕", 9)).pack(anchor="w", pady=(0, 8))

        tk.Label(tab_custom, text="키워드 직접 입력  (쉼표로 구분, 키워드 간 OR 검색)",
                 font=("맑은 고딕", 9)).pack(anchor="w")
        self._custom_kw = ttk.Entry(tab_custom)
        self._custom_kw.pack(fill="x", pady=(2, 2))
        tk.Label(tab_custom, text="예: 전산, 데이터, IT직",
                 fg="#888", font=("맑은 고딕", 8)).pack(anchor="w", pady=(0, 10))

        ttk.Separator(tab_custom, orient="horizontal").pack(fill="x", pady=(0, 8))

        flt_frm = ttk.Frame(tab_custom)
        flt_frm.pack(fill="x")
        flt_frm.columnconfigure(1, weight=1)
        flt_frm.columnconfigure(3, weight=1)

        tk.Label(flt_frm, text="포함 단어:").grid(row=0, column=0, sticky="e", padx=(0, 6), pady=4)
        self._include_words = ttk.Entry(flt_frm)
        self._include_words.grid(row=0, column=1, sticky="ew", pady=4)

        tk.Label(flt_frm, text="제외 단어:", padx=12).grid(row=0, column=2, sticky="e", pady=4)
        self._exclude_words = ttk.Entry(flt_frm)
        self._exclude_words.grid(row=0, column=3, sticky="ew", pady=4)

        tk.Label(flt_frm,
                 text="ALIO 공개 API(opendata.alio.go.kr)에서 받아온 공고를 이 프로그램이 추가로 걸러냅니다.  적용 대상: 기관명 + 공고 제목\n"
                      "API 자체는 단순 키워드 검색만 지원하므로, AND 조합·제외는 여기서만 가능합니다.\n"
                      "| = OR (둘 중 하나)   & = AND (둘 다 포함)     예) 전산 | IT직     예) 병원 | 의료원",
                 fg="#888", font=("맑은 고딕", 8), justify="left"
                 ).grid(row=1, column=0, columnspan=4, sticky="w")

        # ── 수집 옵션 ─────────────────────────────────────────────
        opt_lf = ttk.LabelFrame(frm, text="  수집 옵션  ", padding=10)
        opt_lf.grid(row=row, column=0, sticky="ew", padx=12, pady=4)
        row += 1

        r0 = ttk.Frame(opt_lf)
        r0.pack(fill="x", pady=2)

        tk.Label(r0, text="키워드당 최대").pack(side="left")
        self._max_per_kw = ttk.Spinbox(r0, from_=10, to=9999, width=7)
        self._max_per_kw.set(500)
        self._max_per_kw.pack(side="left", padx=6)
        tk.Label(r0, text="개").pack(side="left")

        tk.Label(r0, text="    ").pack(side="left")
        tk.Label(r0, text="최근").pack(side="left")
        self._days_back = ttk.Spinbox(r0, from_=30, to=1825, width=6)
        self._days_back.set(365)
        self._days_back.pack(side="left", padx=6)
        tk.Label(r0, text="일 이내").pack(side="left")

        r1 = ttk.Frame(opt_lf)
        r1.pack(fill="x", pady=4)
        self._dl_files = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            r1,
            text="첨부파일 다운로드  (HWP·PDF 파싱 → NCS·자격증 정확도 향상, 시간 더 소요)",
            variable=self._dl_files,
        ).pack(side="left")

        # ── 저장 위치 ─────────────────────────────────────────────
        out_lf = ttk.LabelFrame(frm, text="  결과 저장 위치  ", padding=10)
        out_lf.grid(row=row, column=0, sticky="ew", padx=12, pady=4)
        out_lf.columnconfigure(0, weight=1)
        row += 1

        self._output_dir = ttk.Entry(out_lf)
        self._output_dir.insert(0, str(DEFAULT_OUTPUT_DIR))
        self._output_dir.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(out_lf, text="찾아보기", command=self._browse_output).grid(row=0, column=1)

        # ── 버튼 ─────────────────────────────────────────────────
        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=row, column=0, pady=12)

        self._start_btn = ttk.Button(
            btn_frame, text="  수집 시작  ", command=self._start,
            style="Accent.TButton")
        self._start_btn.pack(side="left", padx=8, ipadx=10, ipady=4)
        ttk.Button(btn_frame, text="설정 저장", command=self._save_settings).pack(side="left", padx=4)

    # ── 진행 탭 ──────────────────────────────────────────────────
    def _build_progress_tab(self):
        tab = self.tab_prog
        tab.columnconfigure(0, weight=1)
        tab.rowconfigure(2, weight=1)

        # ── 진행바 영역 ───────────────────────────────────────────
        prog_frm = ttk.Frame(tab, padding=(12, 8))
        prog_frm.grid(row=0, column=0, sticky="ew")
        prog_frm.columnconfigure(0, weight=1)

        def _make_pbar_section(parent, title, row_start):
            """(label_status, label_pct, progressbar, label_count) 반환"""
            # 섹션 제목
            ttk.Separator(parent, orient="horizontal").grid(
                row=row_start, column=0, columnspan=2, sticky="ew", pady=(6, 2))
            tk.Label(parent, text=title, font=("맑은 고딕", 9, "bold"),
                     fg="#1565c0", bg="#f5f5f5", anchor="w").grid(
                row=row_start + 1, column=0, sticky="w")

            header = ttk.Frame(parent)
            header.grid(row=row_start + 2, column=0, columnspan=2, sticky="ew")
            header.columnconfigure(0, weight=1)

            lbl_status = tk.Label(header, text="대기 중", font=("맑은 고딕", 9),
                                  anchor="w", bg="#f5f5f5", fg="#333")
            lbl_status.grid(row=0, column=0, sticky="w")

            lbl_pct = tk.Label(header, text="0%", font=("맑은 고딕", 10, "bold"),
                               fg="#1565c0", bg="#f5f5f5")
            lbl_pct.grid(row=0, column=1, sticky="e")

            pbar = ttk.Progressbar(parent, mode="determinate", maximum=100)
            pbar.grid(row=row_start + 3, column=0, columnspan=2, sticky="ew", pady=2)

            lbl_count = tk.Label(parent, text="", font=("맑은 고딕", 8),
                                 fg="#666", bg="#f5f5f5", anchor="w")
            lbl_count.grid(row=row_start + 4, column=0, columnspan=2, sticky="w")

            return lbl_status, lbl_pct, pbar, lbl_count

        (self._dl_lbl_status, self._dl_lbl_pct,
         self._dl_pbar, self._dl_lbl_count) = _make_pbar_section(
            prog_frm, "공고 수집 · 첨부파일 다운로드", 0)

        (self._an_lbl_status, self._an_lbl_pct,
         self._an_pbar, self._an_lbl_count) = _make_pbar_section(
            prog_frm, "NCS · 자격증 분석", 6)

        # 로그
        self._log = scrolledtext.ScrolledText(
            tab, wrap="word", font=("Consolas", 9), state="disabled",
            bg="#1e1e2e", fg="#cdd6f4", insertbackground="white",
            selectbackground="#3d3d5c",
        )
        self._log.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 4))

        # 하단 버튼
        btn_frm = ttk.Frame(tab)
        btn_frm.grid(row=2, column=0, pady=6)

        self._stop_btn = ttk.Button(btn_frm, text="중지", command=self._stop, state="disabled")
        self._stop_btn.pack(side="left", padx=8)

        self._open_dir_btn = ttk.Button(btn_frm, text="폴더 열기",
                                        command=self._open_dir, state="disabled")
        self._open_dir_btn.pack(side="left", padx=4)

        self._open_xl_btn = ttk.Button(btn_frm, text="Excel 열기",
                                       command=self._open_excel, state="disabled")
        self._open_xl_btn.pack(side="left", padx=4)

    # ── 이벤트 ───────────────────────────────────────────────────
    def _browse_output(self):
        d = filedialog.askdirectory(initialdir=self._output_dir.get() or str(Path.home()))
        if d:
            self._output_dir.delete(0, "end")
            self._output_dir.insert(0, d)

    def _get_keywords(self) -> list:
        kws = [k for k, v in self._alio_vars.items() if v.get()]
        # NCS 분류명은 "경영·회계·사무" 형태라 그대로 검색하면 0개 나옴
        # → "·" 기준으로 분리해서 각각 키워드로 사용
        for k, v in self._ncs_vars.items():
            if v.get():
                kws.extend(part.strip() for part in k.split("·") if part.strip())
        kws += [k.strip() for k in self._custom_kw.get().split(",") if k.strip()]
        return list(dict.fromkeys(kws))  # 순서 유지하며 중복 제거

    def _start(self):
        if not HAS_REQUESTS:
            messagebox.showerror("오류", "requests 패키지가 없습니다.\nsetup.bat을 먼저 실행하세요.")
            return
        if not HAS_EXCEL:
            messagebox.showerror("오류", "pandas/openpyxl이 없습니다.\nsetup.bat을 먼저 실행하세요.")
            return

        keywords = self._get_keywords()
        if not keywords:
            messagebox.showwarning("경고", "키워드를 최소 1개 선택하세요.")
            return

        output_dir = self._output_dir.get().strip()
        if not output_dir:
            messagebox.showwarning("경고", "저장 위치를 지정하세요.")
            return

        try:
            max_per_kw = int(self._max_per_kw.get())
        except ValueError:
            messagebox.showerror("오류", "최대 공고 수를 올바르게 입력하세요.")
            return

        params = {
            "keywords":      keywords,
            "include_expr":  self._include_words.get().strip(),
            "exclude_expr":  self._exclude_words.get().strip(),
            "max_per_kw":    max_per_kw,
            "output_dir":    output_dir,
            "download_files": self._dl_files.get(),
        }

        self._stop_ev.clear()
        self._clear_log()
        self._last_excel = ""
        # 다운로드 진행바 초기화
        self._dl_pbar["value"] = 0
        self._dl_lbl_pct.config(text="0%")
        self._dl_lbl_status.config(text="대기 중")
        self._dl_lbl_count.config(text="")
        # 분석 진행바 초기화
        self._an_pbar["value"] = 0
        self._an_lbl_pct.config(text="0%")
        self._an_lbl_status.config(text="대기 중")
        self._an_lbl_count.config(text="")

        self._start_btn.config(state="disabled")
        self._stop_btn.config(state="normal")
        self._open_dir_btn.config(state="disabled")
        self._open_xl_btn.config(state="disabled")

        self.nb.select(self.tab_prog)

        self._worker = threading.Thread(
            target=collect_worker,
            args=(params, self._msg_q, self._stop_ev),
            daemon=True,
        )
        self._worker.start()

    def _stop(self):
        self._stop_ev.set()
        self._stop_btn.config(state="disabled")
        self._log_write("⚠  중지 요청... 현재 작업이 끝나면 멈춥니다.")

    def _open_dir(self):
        d = self._output_dir.get()
        if d and Path(d).exists():
            os.startfile(d)

    def _open_excel(self):
        if self._last_excel and Path(self._last_excel).exists():
            os.startfile(self._last_excel)

    # ── 큐 폴링 ──────────────────────────────────────────────────
    def _poll_queue(self):
        try:
            while True:
                msg = self._msg_q.get_nowait()
                kind = msg[0]

                if kind == "log":
                    self._log_write(msg[1])

                elif kind == "prog_dl":
                    _, cur, total, label = msg
                    pct = int(cur / total * 100) if total > 0 else 0
                    self._dl_pbar["value"] = pct
                    self._dl_lbl_pct.config(text=f"{pct}%")
                    self._dl_lbl_status.config(text=(label[:65] if label else ""))
                    self._dl_lbl_count.config(text=f"{cur:,} / {total:,}개")

                elif kind == "prog_an":
                    _, cur, total, label = msg
                    pct = int(cur / total * 100) if total > 0 else 0
                    self._an_pbar["value"] = pct
                    self._an_lbl_pct.config(text=f"{pct}%")
                    self._an_lbl_status.config(text=(label[:65] if label else ""))
                    self._an_lbl_count.config(text=f"{cur:,} / {total:,}개")

                elif kind == "done":
                    _, excel_path, n = msg
                    self._last_excel = excel_path
                    self._dl_pbar["value"] = 100
                    self._dl_lbl_pct.config(text="100%")
                    self._an_pbar["value"] = 100
                    self._an_lbl_pct.config(text="100%")
                    self._an_lbl_status.config(text="완료!")
                    self._an_lbl_count.config(text=f"총 {n:,}개 공고 분석 완료")
                    self._log_write(f"\n✅ Excel 저장 완료: {excel_path}")
                    self._start_btn.config(state="normal")
                    self._stop_btn.config(state="disabled")
                    self._open_dir_btn.config(state="normal")
                    self._open_xl_btn.config(state="normal")
                    messagebox.showinfo("완료", f"{n}개 공고 분석 완료!\n\n{excel_path}")

                elif kind == "error":
                    self._log_write(f"\n❌ 오류: {msg[1]}")
                    self._start_btn.config(state="normal")
                    self._stop_btn.config(state="disabled")
                    messagebox.showerror("오류 발생", msg[1][:400])

        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    # ── 로그 헬퍼 ────────────────────────────────────────────────
    def _log_write(self, text: str):
        self._log.config(state="normal")
        self._log.insert("end", text + "\n")
        self._log.see("end")
        self._log.config(state="disabled")

    def _clear_log(self):
        self._log.config(state="normal")
        self._log.delete("1.0", "end")
        self._log.config(state="disabled")

    # ── 설정 저장/로드 ────────────────────────────────────────────
    def _save_settings(self):
        cfg = {
            "alio_keywords":  [k for k, v in self._alio_vars.items() if v.get()],
            "ncs_keywords":   [k for k, v in self._ncs_vars.items() if v.get()],
            "custom_kw":      self._custom_kw.get(),
            "include_expr":   self._include_words.get(),
            "exclude_expr":   self._exclude_words.get(),
            "max_per_kw":     self._max_per_kw.get(),
            "days_back":      self._days_back.get(),
            "download_files": self._dl_files.get(),
            "output_dir":     self._output_dir.get(),
        }
        save_config(cfg)
        messagebox.showinfo("저장", "설정이 저장되었습니다.")

    def _load_settings(self):
        cfg = self._cfg
        if not cfg:
            return
        for kw, var in self._alio_vars.items():
            var.set(kw in cfg.get("alio_keywords", list(ALIO_DEFAULT_ON)))
        for kw, var in self._ncs_vars.items():
            var.set(kw in cfg.get("ncs_keywords", []))
        for attr, key in [
            ("_custom_kw",     "custom_kw"),
            ("_include_words", "include_expr"),
            ("_exclude_words", "exclude_expr"),
            ("_output_dir",    "output_dir"),
        ]:
            if key in cfg:
                w = getattr(self, attr)
                w.delete(0, "end")
                w.insert(0, cfg[key])
        if "max_per_kw" in cfg:
            self._max_per_kw.delete(0, "end")
            self._max_per_kw.insert(0, cfg["max_per_kw"])
        if "days_back" in cfg:
            self._days_back.delete(0, "end")
            self._days_back.insert(0, cfg["days_back"])
        if "download_files" in cfg:
            self._dl_files.set(cfg["download_files"])


# ════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
