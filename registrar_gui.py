r"""
부자재 표준견본 등록기 (GUI) — 2단계
 ① 초도분 분석 + 엑셀출력 : 입고검사 엑셀의 품목들을 품목재고조회(ITEMINVv1)에서 초도입고일 조회 →
    초도입고 == 입고일이면 초도분. 결과를 화면 표 + 검토 엑셀로 출력.
 ② 웹앱에 등록 : 초도분만 register.json 생성 → 로컬 웹앱에 자동 등록(브라우저).

⚠️ Windows + 관리자 권한 / iERP 로그인 + '품목재고조회' 창을 띄워둔 상태에서 ① 실행.
실행:  등록기.bat 더블클릭 (또는 pyw registrar_gui.py)
"""
from __future__ import annotations
import datetime
import os
import webbrowser
from pathlib import Path
from urllib.parse import urlparse

import sample_export as se
import sample_register as sr

DEFAULT_XLSXDIR = r"C:\sample-register"            # 검토 엑셀 저장 폴더
DEFAULT_JSON = r"C:\sample-list\register.json"     # register.json 저장 경로(웹앱 폴더 안)
DEFAULT_WEBURL = "http://localhost:5175"


def _ts() -> str:
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def _server_up(url: str) -> bool:
    import urllib.request
    try:
        urllib.request.urlopen(url, timeout=1.0)
        return True
    except Exception:
        return False


_HTTPD = {"srv": None}


def _ensure_server(webdir: Path, url: str):
    """웹이 안 떠 있으면 webdir를 서빙하는 로컬 서버를 띄운다(127.0.0.1, 방화벽 팝업 회피)."""
    if _server_up(url):
        return True, "웹 서버 이미 실행 중"
    if _HTTPD["srv"] is not None:
        return True, "웹 서버 실행 중"
    import functools, http.server, socketserver, threading
    port = urlparse(url).port or 5175
    handler = functools.partial(http.server.SimpleHTTPRequestHandler, directory=str(webdir))
    try:
        httpd = socketserver.TCPServer(("127.0.0.1", port), handler)
    except Exception as e:
        return False, f"웹 서버 시작 실패(포트 {port}): {e}"
    threading.Thread(target=httpd.serve_forever, daemon=True).start()
    _HTTPD["srv"] = httpd
    return True, f"로컬 웹 서버 시작 (127.0.0.1:{port})"


def main():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog

    root = tk.Tk()
    root.title("부자재 표준견본 등록기   (v2026-06-24a · UIA번들수정)")
    try:
        import sv_ttk          # Windows 11 스타일 테마 (pip install sv-ttk)
        sv_ttk.set_theme("light")
    except Exception:
        pass
    # 전체 글씨체: 맑은 고딕 8 (named 폰트 + ttk 스타일 둘 다 강제 → 표/버튼/라벨 통일)
    FONT = ("맑은 고딕", 8)
    try:
        import tkinter.font as tkfont
        for fn in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
                   "TkLabelFont", "TkButtonFont", "TkEntryFont", "TkFixedFont"):
            try:
                tkfont.nametofont(fn).configure(family="맑은 고딕", size=8)
            except Exception:
                pass
    except Exception:
        pass
    try:
        style = ttk.Style()
        for st in ("TLabel", "TButton", "TCheckbutton", "TEntry",
                   "TLabelframe.Label", "Treeview", "Treeview.Heading"):
            style.configure(st, font=FONT)
    except Exception:
        pass

    # 필수 패키지 점검 (pyw는 콘솔이 없어 import 오류가 안 보이므로 안내 표시)
    _need = {"openpyxl": "openpyxl", "pywinauto": "pywinauto", "win32com.client": "pywin32"}
    _missing = []
    for _mod, _pkg in _need.items():
        try:
            __import__(_mod)
        except Exception:
            _missing.append(_pkg)
    if _missing:
        messagebox.showwarning(
            "패키지 필요",
            "다음 패키지가 필요합니다:  " + ", ".join(sorted(set(_missing))) +
            "\n\n명령 프롬프트에서 설치:\n  py -m pip install " + " ".join(sorted(set(_missing))) + " sv-ttk")

    pad = {"padx": 8, "pady": 6}
    state = {"rows": [], "analysis": [], "records": []}

    # ① 입고검사 대상 (자동조회 / 엑셀 직접선택)
    f1 = ttk.LabelFrame(root, text="① 입고검사 대상")
    f1.pack(fill="x", **pad)
    mode = tk.StringVar(value="auto")

    ttk.Radiobutton(f1, text="iERP 자동조회 (품질검사대상조회 · 검사파트=자재파트)",
                    variable=mode, value="auto").grid(row=0, column=0, sticky="w", padx=6, pady=(6, 0))
    fp = ttk.Frame(f1)
    fp.grid(row=1, column=0, sticky="w", padx=26)
    ttk.Label(fp, text="입고일자").pack(side="left", padx=(0, 4))
    e_from = ttk.Entry(fp, width=12); e_from.pack(side="left")
    ttk.Label(fp, text="~").pack(side="left", padx=4)
    e_to = ttk.Entry(fp, width=12); e_to.pack(side="left")
    ttk.Label(fp, text="(YYYY-MM-DD)").pack(side="left", padx=6)

    ttk.Radiobutton(f1, text="엑셀 직접 선택 (이미 받아둔 엑셀)",
                    variable=mode, value="manual").grid(row=2, column=0, sticky="w", padx=6, pady=(8, 0))
    fe = ttk.Frame(f1)
    fe.grid(row=3, column=0, sticky="we", padx=26)
    e_excel = ttk.Entry(fe, width=50); e_excel.pack(side="left", fill="x", expand=True)

    def pick_excel():
        p = filedialog.askopenfilename(title="입고검사 엑셀 선택", parent=root,
                                       filetypes=[("Excel", "*.xlsx"), ("All", "*.*")])
        if p:
            e_excel.delete(0, "end"); e_excel.insert(0, p)

    ttk.Button(fe, text="찾기", command=pick_excel).pack(side="left", padx=6)

    sub_only = tk.BooleanVar(value=False)
    ttk.Checkbutton(f1, text="부자재(접미사 7/5)만", variable=sub_only).grid(row=4, column=0, sticky="w", padx=6, pady=(8, 6))
    f1.columnconfigure(0, weight=1)

    # 진행 상태
    status = tk.StringVar(value="① 조회+초도분 분석 (품질검사대상조회·품목재고조회 자동으로 열림) → ② 웹앱 등록")
    ttk.Label(root, textvariable=status).pack(anchor="w", **pad)

    # 결과 표
    f2 = ttk.LabelFrame(root, text="분석 결과 (초도분 O = 등록 대상)")
    f2.pack(fill="both", expand=True, **pad)
    cols = ("chodo", "code", "name", "customer", "mat_in", "first_in")
    heads = {"chodo": "초도분", "code": "품목코드", "name": "품목명",
             "customer": "고객사", "mat_in": "입고일", "first_in": "초도입고일"}
    widths = {"chodo": 50, "code": 110, "name": 260, "customer": 150, "mat_in": 90, "first_in": 90}
    tv = ttk.Treeview(f2, columns=cols, show="headings", height=6)
    for c in cols:
        tv.heading(c, text=heads[c]); tv.column(c, width=widths[c], anchor="w")
    vs = ttk.Scrollbar(f2, orient="vertical", command=tv.yview)
    tv.configure(yscroll=vs.set)
    tv.pack(side="left", fill="both", expand=True)
    vs.pack(side="right", fill="y")
    tv.tag_configure("yes", background="#e7f7ec")
    tv.tag_configure("manual", foreground="#d00000")   # 수동조회요망 = 빨간 글씨

    # 저장 위치 (검토 엑셀 폴더 / register.json 경로 직접 지정)
    fs = ttk.LabelFrame(root, text="저장 위치")
    fs.pack(fill="x", **pad)
    ttk.Label(fs, text="검토 엑셀 폴더").grid(row=0, column=0, padx=6, pady=4, sticky="w")
    e_xlsxdir = ttk.Entry(fs, width=44)
    e_xlsxdir.grid(row=0, column=1, padx=6, sticky="we"); e_xlsxdir.insert(0, DEFAULT_XLSXDIR)

    def pick_xlsxdir():
        p = filedialog.askdirectory(title="검토 엑셀 저장 폴더", parent=root)
        if p:
            e_xlsxdir.delete(0, "end"); e_xlsxdir.insert(0, p)

    ttk.Button(fs, text="찾기", command=pick_xlsxdir).grid(row=0, column=2, padx=6)

    ttk.Label(fs, text="register.json 경로").grid(row=1, column=0, padx=6, pady=4, sticky="w")
    e_json = ttk.Entry(fs, width=44)
    e_json.grid(row=1, column=1, padx=6, sticky="we"); e_json.insert(0, DEFAULT_JSON)

    def pick_json():
        p = filedialog.asksaveasfilename(title="register.json 저장 위치", parent=root,
                                         defaultextension=".json", initialfile="register.json",
                                         filetypes=[("JSON", "*.json"), ("All", "*.*")])
        if p:
            e_json.delete(0, "end"); e_json.insert(0, p)

    ttk.Button(fs, text="찾기", command=pick_json).grid(row=1, column=2, padx=6)
    ttk.Label(fs, text="※ register.json 은 웹앱 폴더(index.html 있는 곳)에 둬야 ②가 자동 등록됩니다",
              foreground="#888").grid(row=2, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 4))
    fs.columnconfigure(1, weight=1)

    # ② 웹앱 주소
    f3 = ttk.LabelFrame(root, text="② 웹앱 주소")
    f3.pack(fill="x", **pad)
    ttk.Label(f3, text="웹 주소").grid(row=0, column=0, padx=6, pady=4)
    e_url = ttk.Entry(f3, width=40)
    e_url.grid(row=0, column=1, padx=6, sticky="we"); e_url.insert(0, DEFAULT_WEBURL)
    f3.columnconfigure(1, weight=1)

    # 버튼 (① 조회+초도분 분석 한번에 → ② 웹앱 등록)
    bar = ttk.Frame(root)
    bar.pack(side="bottom", pady=10)
    btn1 = ttk.Button(bar, text="①  조회 + 초도분 분석")
    btn1.pack(side="left", padx=5)
    btn_stop = ttk.Button(bar, text="중단", state="disabled")
    btn_stop.pack(side="left", padx=5)
    btn3 = ttk.Button(bar, text="②  웹앱에 등록", state="disabled")
    btn3.pack(side="left", padx=5)

    def set_status(m):
        status.set(m); root.update_idletasks()

    stop_flag = {"stop": False}

    def _report_error(e, where=""):
        """에러 종류+메시지+추적을 팝업에 표시하고 로그파일에도 저장 (콘솔 없는 EXE 대비)."""
        import traceback
        tb = traceback.format_exc()
        head = f"{type(e).__name__}: {e}".strip()
        logpath = None
        try:
            d = Path(e_xlsxdir.get().strip() or DEFAULT_XLSXDIR)
            d.mkdir(parents=True, exist_ok=True)
            logpath = d / "오류로그.txt"
            logpath.write_text(f"[{where}]\n{head}\n\n{tb}", encoding="utf-8")
        except Exception:
            pass
        set_status("오류: " + (head[:60] or "(상세는 팝업 참고)"))
        msg = head + "\n\n" + tb[-1400:]
        if logpath:
            msg += f"\n\n(로그 저장: {logpath})"
        messagebox.showerror("오류", msg)

    def on_stop():
        stop_flag["stop"] = True
        set_status("중단 중... (현재 품목 끝나면 멈춤)")
        btn_stop.configure(state="disabled")

    def _fill_table(analysis, looked_up=False):
        tv.delete(*tv.get_children())
        for a in sorted(analysis, key=lambda x: (not x["is_chodo"], x["code"])):
            first = a["first_in"]
            if a["is_chodo"]:
                flag, fin, tag = "O", first, "yes"
            elif first:
                flag, fin, tag = "X", first, ""
            elif looked_up:                      # 조회했는데 초도입고 빈값 = 수동조회요망
                flag, fin, tag = "수동", "수동조회요망", "manual"
            else:                                # ① 단계: 아직 초도입고 조회 전
                flag, fin, tag = "—", "—", ""
            tv.insert("", "end", tags=(tag,),
                      values=(flag, a["code"], a["name"], a["customer"], a["mat_in"], fin))

    # ① 조회 + 초도분 분석 (통합): 품질검사대상조회 자동조회+엑셀출력 → 품목재고조회 초도입고 → 분석
    def on_run():
        stop_flag["stop"] = False
        btn1.configure(state="disabled"); btn3.configure(state="disabled")
        btn_stop.configure(state="normal")
        tv.delete(*tv.get_children())
        state["rows"] = []; state["analysis"] = []; state["records"] = []
        try:
            # --- 1단계: 품목 리스트 확보 (자동조회 엑셀출력 / 엑셀 직접선택) ---
            if mode.get() == "auto":
                df, dt = e_from.get().strip(), e_to.get().strip()
                if not df:
                    raise RuntimeError("입고일자 시작일을 입력하세요 (YYYY-MM-DD).")
                try:
                    se.close_iteminv()    # 품목재고조회 열려있으면 닫기(품질검사대상조회 가림 방지)
                except Exception:
                    pass
                set_status("품질검사대상조회 자동 조회 + 엑셀출력 중...")
                root.update()
                path = str(se.export_receipt(df, dt or df))
                set_status(f"엑셀 받음: {Path(path).name} — 읽는 중...")
            else:
                path = e_excel.get().strip()
                if not path or not Path(path).exists():
                    raise RuntimeError("엑셀 파일을 선택하세요.")
                set_status("엑셀 읽는 중...")
            rows = se.load_receipt_excel(path)
            if not rows:
                raise RuntimeError("엑셀에서 품목을 읽지 못했습니다. (형식 확인)")
            if mode.get() == "manual":
                df, dt = e_from.get().strip(), e_to.get().strip()
                if df or dt:
                    def _inrange(r):
                        d = se._norm_date(r.get("mat_in"))
                        return d and (not df or d >= df) and (not dt or d <= dt)
                    rows = [r for r in rows if _inrange(r)]
                    if not rows:
                        raise RuntimeError("입고일 기간에 해당하는 품목이 없습니다.")
            state["rows"] = rows
            if stop_flag["stop"]:
                set_status("중단됨"); return

            # --- 2단계: 품목재고조회 자동 열고 초도입고 조회 ---
            set_status("품목재고조회 자동으로 여는 중... (그리드 우클릭 → 재고현황조회)")
            root.update()
            se.ensure_iteminv_open()
            set_status(f"{len(rows)}행 — 초도입고 조회 중... (마우스/키보드 건드리지 마세요)")

            def prog(i, total, code, d):
                set_status(f"초도입고 조회 {i}/{total} : {code} → {d or '(없음)'}")
                root.update()

            fmap = se.lookup_first_in([r.get("code") for r in rows], progress=prog,
                                      should_stop=lambda: stop_flag["stop"])
            if stop_flag["stop"]:
                set_status("중단됨"); messagebox.showinfo("중단", "조회를 중단했습니다.")
                return
            for r in rows:
                r["first_in"] = fmap.get(str(r.get("code") or "").strip(), "")

            # --- 분석 + 검토엑셀 ---
            analysis = sr.analyze(rows)
            state["analysis"] = analysis
            state["records"] = sr.records_from_analysis(analysis, subsidiary_only=sub_only.get())
            _fill_table(analysis, looked_up=True)

            xdir = Path(e_xlsxdir.get().strip() or DEFAULT_XLSXDIR)
            xdir.mkdir(parents=True, exist_ok=True)
            out_xlsx = xdir / f"초도분분석_{_ts()}.xlsx"
            sr.write_review_excel(analysis, out_xlsx)
            n = len(state["records"]); nch = sum(1 for a in analysis if a["is_chodo"])
            nman = sum(1 for a in analysis if not a["first_in"])
            set_status(f"완료 — 초도분 {nch}건 / 등록대상 {n}건"
                       + (f" / 수동조회요망 {nman}건(빨강)" if nman else "")
                       + f" · 검토엑셀: {out_xlsx.name}")
            try:
                os.startfile(str(out_xlsx))
            except Exception:
                pass
            btn3.configure(state=("normal" if n else "disabled"))
        except SystemExit as e:
            set_status("중단"); messagebox.showerror("안내", str(e) or "(중단됨)")
        except Exception as e:
            _report_error(e, "①조회+분석")
        finally:
            btn1.configure(state="normal"); btn_stop.configure(state="disabled")

    def on_register():
        recs = state.get("records") or []
        if not recs:
            messagebox.showwarning("확인", "등록할 초도분이 없습니다. 먼저 ①(조회 + 초도분 분석)을 실행하세요.")
            return
        jsonpath = Path(e_json.get().strip() or DEFAULT_JSON)
        url = (e_url.get().strip() or DEFAULT_WEBURL).rstrip("/")
        webdir = jsonpath.parent                       # register.json 이 놓인 폴더 = 웹 서빙 폴더
        try:
            sr.write_register_json(recs, jsonpath)     # 사용자가 지정한 위치에 저장
            if not (webdir / "index.html").exists():
                messagebox.showwarning(
                    "주의", f"{webdir}\n에 웹앱(index.html)이 없어 자동 등록 페이지가 안 열릴 수 있습니다.\n"
                    "register.json 을 웹앱 폴더(예: C:\\sample-list)에 저장하세요.\n"
                    "(파일은 저장됐으니 웹에서 직접 '가져오기'로 올릴 수도 있습니다.)")
            ok, msg = _ensure_server(webdir, url)
            set_status(msg)
            if not ok:
                messagebox.showerror("오류", msg)
                return
            webbrowser.open(f"{url}/?import={jsonpath.name}&ts={_ts()}")
            set_status(f"웹앱 열림 — 자동 등록 {len(recs)}건 (중복 자동 제외)")
            messagebox.showinfo("등록", f"register.json 저장: {jsonpath}\n웹앱에서 {len(recs)}건을 가져옵니다.")
        except Exception as e:
            _report_error(e, "②웹앱등록")

    btn1.configure(command=on_run)
    btn_stop.configure(command=on_stop)
    btn3.configure(command=on_register)

    # 내용(버튼 포함)에 맞춰 기본 창 크기 자동 계산 → 처음 열릴 때 ①조회 버튼까지 보이게
    root.update_idletasks()
    w = max(620, root.winfo_reqwidth() + 24)
    h = root.winfo_reqheight() + 24
    root.geometry(f"{w}x{h}")
    root.minsize(560, 420)
    root.resizable(True, True)
    root.mainloop()


if __name__ == "__main__":
    main()
