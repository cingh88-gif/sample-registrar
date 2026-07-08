r"""
iERP 부자재 입고 화면 → '초도분' 품목 추출.

초도분 판정(사용자 확정 규칙):
  해당 날짜로 조회 → 품목 우클릭 → '초도입고일' == '자재입고일' 이면 초도분.

두 가지 경로를 제공한다:
  (1) load_receipt_excel(path) + chodo_rows(...)  ← 화면 '엑셀출력'에 두 날짜 컬럼이 있으면 이 경로(오프라인 가능)
  (2) export_receipt(...) ← 품질검사대상조회 자동 조회+엑셀출력 / lookup_first_in(...) ← 품목재고조회 초도입고. (Windows 전용)

⚠️ (2)의 컨트롤 ID는 inspect_ierp.py 정찰(Phase 0) 후 RECEIPT dict에 채워야 한다.
   현재는 플레이스홀더이며, 실행 시 명확한 안내와 함께 중단된다.

의존성: openpyxl (필수), pywinauto (화면 자동화 경로만)
"""
from __future__ import annotations
import datetime
from pathlib import Path

# 화면 자동화 엔진은 기존 도구를 그대로 재사용한다.
try:
    import ierp_export as ie  # connect_window / ensure_screen_open / export_screen / EXPORT_DIR / _ts
except Exception:  # 단독 import(테스트 등) 시에도 오프라인 함수는 동작하도록
    ie = None


# 엑셀출력 경로용 유연한 컬럼 매핑(한글 변형 흡수)
# ※ iERP 품질검사대상조회 출력은 코드성 컬럼이 [코드, 표시명] 쌍으로 같은 헤더가 2번 나온다.
#    customer/category/supply 는 '표시명'(뒤 컬럼)을 써야 하므로 PREFER_LAST 로 마지막 매칭 선택.
COLMAP = {
    "code":         ["품목코드", "품번", "itemcode"],
    "name":         ["품명", "품목명", "itemname"],
    "customer":     ["고객", "고객사", "거래처명", "거래처", "customer"],
    "category":     ["관리유형", "제형", "품목유형", "유형", "category"],   # 관리유형(단상자/파우치 등) 우선
    "supply":       ["입고유형", "사급", "자급", "사자급", "구분"],           # 사급반입 → 사급
    "first_in":     ["초도입고일", "초도입고일자", "초도입고날짜", "초도입고"],
    "mat_in":       ["입고일시", "자재입고일자", "자재입고일", "입고일자", "입고일"],
}
# 같은 헤더가 [코드,표시명]으로 중복될 때 '표시명'(마지막 매칭)을 쓰는 필드
PREFER_LAST = {"customer", "category", "supply"}


def _norm_date(v) -> str:
    """다양한 날짜 표기를 'YYYY-MM-DD'로 정규화. 비거나 파싱불가면 ''."""
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip()
    if not s:
        return ""
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) >= 8:                      # 20260619 / 2026.06.19 / 2026-06-19 …
        return f"{digits[0:4]}-{digits[4:6]}-{digits[6:8]}"
    return s


def is_chodo(row: dict) -> bool:
    """초도입고일 == 자재입고일 (둘 다 값이 있을 때)."""
    a = _norm_date(row.get("first_in"))
    b = _norm_date(row.get("mat_in"))
    return bool(a) and a == b


def chodo_rows(rows: list[dict]) -> list[dict]:
    """초도분 행만 추린다."""
    return [r for r in rows if is_chodo(r)]


def fill_first_in_from_history(rows: list[dict]) -> list[dict]:
    """엑셀에 '초도입고일' 컬럼이 없을 때: 입고이력(자재입고일)에서
    품목코드별 '최초 입고일'을 계산해 각 행의 first_in 에 채운다.
    (사용자 확정: 초도입고일 == 그 품목코드의 최초 입고일)

    ⚠️ 정확하려면 엑셀이 '전체 이력'(충분히 과거~현재)을 담아야 한다.
       하루치만 뽑으면 모든 행이 초도분으로 잘못 잡힌다 → history_date_range()로 범위 확인."""
    first: dict[str, str] = {}
    for r in rows:
        code = str(r.get("code") or "").strip()
        d = _norm_date(r.get("mat_in"))
        if not code or not d:
            continue
        if code not in first or d < first[code]:
            first[code] = d
    for r in rows:
        code = str(r.get("code") or "").strip()
        r["first_in"] = first.get(code, "")
    return rows


def history_date_range(rows: list[dict]) -> tuple[str, str]:
    """입고일(자재입고일)의 (최소, 최대) — 엑셀이 충분히 넓은 이력인지 점검용."""
    ds = [_norm_date(r.get("mat_in")) for r in rows]
    ds = [d for d in ds if d]
    return (min(ds), max(ds)) if ds else ("", "")


# ── 품목재고조회(ITEMINVv1)에서 품목별 '초도입고일' 읽기 (정찰 확정값) ──────────
# 정찰 덤프(win32_…ITEMINVv1….txt) 기준 컨트롤 ID:
#   품목코드 입력 txtMITEM / 초도입고 txtMREFD / 품목명 txtMPNAM / 사·자급 txtMTYP2D
#   조회 버튼은 toolStrip1 안(uia 백엔드에서 title='조회'로 접근)
ITEMINV = dict(
    title_re=r".*ITEMINVv1.*",     # '품목재고조회 (Location별)v1 (ITEMINVv1) …'
    item_edit="txtMITEM",          # 품목코드 입력
    first_in_edit="txtMREFD",      # 초도입고일 (M-REceived-First-Date)
    name_edit="txtMPNAM",          # 품목명(확인용)
    query_btn="조회",              # 툴바 조회(uia)
)


def _find_window(pats, what="창", hint=""):
    """제목에 pats 중 하나가 포함된 창을 Desktop에서 찾는다. (win32 우선, '보이는 창' 우선)
    중복/숨은 창이 떠 있어도 화면에 보이는 창을 잡도록 한다.
    반환: (window, backend). 못 찾으면 열린 창 목록과 함께 중단."""
    try:
        from pywinauto import Desktop
    except Exception:
        raise SystemExit("pywinauto 미설치 — 'py -m pip install pywinauto' 후 실행하세요.")
    for backend in ("win32", "uia"):
        try:
            wins = Desktop(backend=backend).windows()
        except Exception:
            continue
        matches = []
        for w in wins:
            try:
                t = w.window_text() or ""
            except Exception:
                continue
            if any(p in t for p in pats):
                vis = True
                try:
                    vis = bool(w.is_visible())
                except Exception:
                    pass
                area = 0
                try:
                    r = w.rectangle(); area = max(0, (r.right - r.left)) * max(0, (r.bottom - r.top))
                except Exception:
                    pass
                matches.append((vis, area, w))
        if matches:
            # 보이는 창 우선, 그중 가장 큰(=실제 작업) 창
            matches.sort(key=lambda m: (m[0], m[1]), reverse=True)
            return matches[0][2], backend
    found = []
    for backend in ("win32", "uia"):
        try:
            for w in Desktop(backend=backend).windows():
                t = (w.window_text() or "").strip()
                if t:
                    found.append(f"[{backend}] {t}")
        except Exception:
            pass
    raise SystemExit(
        f"{what} 창을 찾지 못했습니다.\n" + (hint + "\n" if hint else "") +
        "열린 창(일부):\n  " + "\n  ".join(sorted(set(found))[:30]))


def _find_iteminv_window():
    return _find_window(("ITEMINVv1", "품목재고조회"), "품목재고조회(ITEMINVv1)",
                        "→ 품질검사대상조회에서 품목을 우클릭해 '품목재고조회' 창을 띄워둔 뒤 다시 실행하세요.")


# 우클릭 컨텍스트 메뉴에서 품목재고조회를 여는 항목 텍스트 후보(우선순위, 공백 무시 매칭)
ITEMINV_OPEN_MENU = ("재고현황조회", "품목재고조회", "재고조회")
# 메뉴에서 '재고현황조회'가 위에서 몇 번째인지 (키보드 네비 백업용): 품목기본정보(1) → 재고현황조회(2)
ITEMINV_OPEN_INDEX = 2


def _click_context_menu(owner_win, texts):
    """우클릭 후 뜬 컨텍스트 메뉴(별도 팝업창)에서 texts 중 하나의 MenuItem을 클릭.
    모든 창의 MenuItem을 모아 texts 우선순위대로 매칭한다."""
    import time
    from pywinauto import Desktop
    time.sleep(0.3)
    items = []
    try:
        for w in Desktop(backend="uia").windows():
            try:
                items += w.descendants(control_type="MenuItem")
            except Exception:
                pass
    except Exception:
        pass
    def norm(s):
        return (s or "").replace(" ", "")
    for t in texts:                              # 우선순위 순서 (공백 무시)
        tn = norm(t)
        for it in items:
            try:
                wt = norm(it.window_text())
            except Exception:
                continue
            if tn and tn in wt:
                for action in ("click_input", "invoke", "select"):
                    try:
                        getattr(it, action)()
                        return True
                    except Exception:
                        continue
    return False


def ensure_iteminv_open(timeout: float = 10.0):
    """품목재고조회(ITEMINVv1)가 열려 있으면 그 창을, 없으면 품질검사대상조회 그리드를
    우클릭해 '재고현황조회'를 눌러 자동으로 연다. 반환: (window, backend)."""
    import time
    try:
        return _find_iteminv_window()          # 이미 열려 있으면 그대로
    except SystemExit:
        pass
    qc, backend = _find_window(RECEIPT_WIN, "품질검사대상조회",
                               "→ 먼저 ① 조회로 품질검사대상조회를 띄우고 조회하세요.")
    try:
        qc.set_focus()
    except Exception:
        pass
    time.sleep(0.3)
    ctl = _control_map(qc)
    grid = ctl.get(RECEIPT_CFG.get("grid", "dataGridView1"))
    if grid is None:
        raise SystemExit("그리드를 찾지 못해 품목재고조회를 자동으로 열 수 없습니다.")
    try:
        grid.set_focus()
    except Exception:
        pass
    # 첫 데이터 행을 우클릭 (헤더 아래)
    try:
        grid.right_click_input(coords=(45, 50))
    except Exception as e:
        raise SystemExit(f"그리드 우클릭 실패: {e}")
    time.sleep(1.0)                       # 우클릭 후 1초만 대기

    def _wait(sec):
        end = time.time() + sec
        while time.time() < end:
            try:
                return _find_iteminv_window()
            except SystemExit:
                time.sleep(0.3)
        return None

    # 시도1: 키보드로 재고현황조회(위에서 N번째) 선택 — 빠르고 안정적
    try:
        from pywinauto.keyboard import send_keys
        send_keys("{DOWN}" * ITEMINV_OPEN_INDEX + "{ENTER}")
    except Exception:
        pass
    win = _wait(2.5)
    if win:
        return win
    # 시도2: uia 메뉴 항목 텍스트 클릭(폴백)
    _click_context_menu(qc, ITEMINV_OPEN_MENU)
    win = _wait(timeout)
    if win:
        return win
    raise SystemExit(
        "재고현황조회 메뉴를 눌러 품목재고조회를 열지 못했습니다.\n"
        "→ '재고현황조회'가 우클릭 메뉴에서 위에서 몇 번째인지 알려주세요(현재 2번째로 설정).")


def close_iteminv():
    """품목재고조회가 열려 있으면 닫는다 (품질검사대상조회 조작 전 방해 제거)."""
    import time
    try:
        win, backend = _find_iteminv_window()
    except SystemExit:
        return
    try:
        win.close()
        time.sleep(0.5)
    except Exception:
        pass


def _control_map(win) -> dict:
    """창의 모든 하위 컨트롤을 automation_id → 컨트롤 래퍼로 매핑(정찰 도구와 동일 방식).
    Desktop().windows() 가 주는 래퍼는 child_window()가 없어 descendants()로 직접 찾는다."""
    m = {}
    try:
        ctrls = win.descendants()
    except Exception:
        ctrls = []
    for c in ctrls:
        try:
            aid = getattr(c.element_info, "automation_id", None)
        except Exception:
            aid = None
        if aid and aid not in m:
            m[aid] = c
    return m


def _ensure_controls(win, backend):
    """컨트롤(auto_id)이 보이는 백엔드를 보장한다. win32에선 auto_id가 0개로 안 보이는
    PC가 있어, 비면 핸들로 uia/win32 재연결해서 보이는 쪽을 쓴다. 반환 (win, backend, ctl)."""
    ctl = _control_map(win)
    if ctl:
        return win, backend, ctl
    handle = getattr(win, "handle", None)
    if handle is not None:
        from pywinauto import Application
        for be in ("uia", "win32"):
            try:
                w2 = Application(backend=be).connect(handle=handle).window(handle=handle).wrapper_object()
                c2 = _control_map(w2)
                if c2:
                    return w2, be, c2
            except Exception:
                continue
    return win, backend, ctl


def _set_edit(ctl, value):
    try:
        ctl.set_edit_text(value)
    except Exception:
        try:
            ctl.set_text(value)
        except Exception:
            ctl.type_keys("^a{DEL}" + value, with_spaces=True, set_foreground=True)


def _text(ctl) -> str:
    try:
        return (ctl.window_text() or "").strip()
    except Exception:
        return ""


# 조회 버튼은 toolStrip1(자동ID 'toolStrip1') 안의 '맨 왼쪽' 버튼.
# win32 백엔드에선 개별 버튼이 안 잡혀 toolStrip 영역 좌측을 좌표로 클릭한다.
QUERY_BTN_X_OFFSET = 28   # toolStrip1 좌측에서 '조회' 아이콘까지 (안 맞으면 조정)


def _click_query(ctl, win_handle=None) -> bool:
    """toolStrip1 좌측의 조회 버튼 클릭. 실패 시 uia(핸들연결)로 title='조회' 버튼 시도."""
    ts = ctl.get("toolStrip1")
    if ts is not None:
        try:
            r = ts.rectangle()
            ts.click_input(coords=(QUERY_BTN_X_OFFSET, (r.bottom - r.top) // 2))
            return True
        except Exception:
            pass
    if win_handle is not None:                         # 폴백: uia로 핸들 연결 → 조회 버튼
        try:
            from pywinauto import Application
            app = Application(backend="uia").connect(handle=win_handle)
            app.window(handle=win_handle).child_window(title="조회", control_type="Button").click_input()
            return True
        except Exception:
            pass
    return False


def lookup_first_in(codes, query_wait: float = 0.8, settle_timeout: float = 9.0,
                    progress=None, should_stop=None) -> dict:
    """이미 열려 있는 품목재고조회(ITEMINVv1) 창에 품목코드를 하나씩 넣어 '초도입고일'을 읽는다.
    반환: {품목코드: 'YYYY-MM-DD' 또는 ''}.  (Windows + pywinauto, iERP 로그인 필요)

    progress(i, total, code, first_in_date): 매 건 후 호출(GUI 진행표시/UI 갱신용).
    조회 완료 판정은 '품목명(txtMPNAM) 변경'으로 한다(코드 입력 후 DB 조회가 끝나면 이름이 바뀜).
    ★ 전제: 품목재고조회 화면을 띄워둔 상태."""
    import time
    win, backend = _find_iteminv_window()
    win, backend, ctl = _ensure_controls(win, backend)   # 컨트롤 보이는 백엔드 보장
    print(f"  품목재고조회 연결됨 (backend={backend}, 컨트롤 {len(ctl)}개)")
    try:
        win.set_focus()
    except Exception:
        pass
    time.sleep(0.3)

    item_edit = ctl.get(ITEMINV["item_edit"])
    first_edit = ctl.get(ITEMINV["first_in_edit"])
    name_edit = ctl.get(ITEMINV["name_edit"])
    if item_edit is None or first_edit is None:
        have = ", ".join(sorted(k for k in ctl if k)[:30])
        raise SystemExit("품목코드/초도입고 컨트롤(txtMITEM/txtMREFD)을 찾지 못했습니다.\n"
                         f"  찾은 auto_id 일부: {have}")
    win_handle = getattr(win, "handle", None)
    # 첫 코드 직전, 조회 버튼 클릭 경로가 동작하는지 1회 점검
    if ctl.get("toolStrip1") is None and win_handle is None:
        print("  ※ 조회 버튼(toolStrip1)을 찾지 못했습니다 — 클릭이 안 되면 알려주세요.")

    uniq, seen = [], set()
    for c in codes:
        c = str(c or "").strip()
        if c and c not in seen:
            seen.add(c); uniq.append(c)

    result = {}
    for i, code in enumerate(uniq, 1):
        if should_stop and should_stop():
            print("  중단 요청 — 초도입고 조회를 멈춥니다.")
            break
        try:
            prev_name = _text(name_edit) if name_edit is not None else ""
            item_edit.set_focus()
            _set_edit(item_edit, code)               # 품목코드 입력
            time.sleep(0.3)                          # 입력값 반영 대기(조회 전)
            # 입고된 품목은 초도입고일이 반드시 있음 → 빈값이면 최대 4회 재조회
            val = ""
            tries = 0
            for attempt in range(4):
                tries = attempt + 1
                mult = attempt + 1                       # 1차=1배, 재조회 2차=2배, 3차=3배, 4차=4배
                clicked = _click_query(ctl, win_handle)  # ★ 조회 버튼 클릭
                if not clicked and i == 1 and attempt == 0:
                    print("  ※ 조회 버튼 클릭 실패 — 좌표 보정이 필요할 수 있습니다.")
                if attempt == 0 and name_edit is not None:
                    end = time.time() + settle_timeout   # 첫 시도: 새 품목 로드(품목명 변경) 대기
                    while time.time() < end:
                        c = _text(name_edit)
                        if c and c != prev_name:
                            break
                        time.sleep(0.2)
                else:
                    end = time.time() + 3.0 * mult       # 재조회: 배수만큼 더 길게 대기
                    while time.time() < end:
                        if _norm_date(_text(first_edit)):
                            break
                        time.sleep(0.2)
                time.sleep(query_wait * mult)            # 화면 값 갱신 여유(재조회일수록 길게)
                val = _norm_date(_text(first_edit))
                if val:
                    break
                time.sleep(0.5 * mult)                   # 재조회 전 대기
            result[code] = val
            tag = "" if val else f"  [{tries}회 조회에도 초도입고 없음 — 확인필요]"
            print(f"  ({i}/{len(uniq)}) {code} → 초도입고 {val or '(없음)'} (조회 {tries}회){tag}")
        except Exception as e:
            result[code] = ""
            print(f"  ({i}/{len(uniq)}) {code} → 조회 실패: {e}")
        if progress:
            try:
                progress(i, len(uniq), code, result.get(code, ""))
            except Exception:
                pass
    return result


# ── 품질검사대상조회(PM70111Uv2) 자동 조회 + 엑셀출력 ────────────────────────────
# 입고일자 기간 + 검사파트=자재파트 설정 → 조회 → 엑셀출력 → .xlsx 저장.
# 컨트롤 ID는 inspect_ierp.py "품질검사대상조회" 정찰로 확정함.
RECEIPT_WIN = ("PM70111Uv2", "품질검사대상조회")
RECEIPT_CFG = dict(
    program_id="PM70111Uv2",  # F9 실행창에 입력할 프로그램 ID(풀네임)
    part_combo="cmbRPINSM",   # 검사파트 콤보
    part_value="자재파트",     # 검사파트 = 자재파트
    date_from="dtpFRDATE",    # 입고일자 시작 (DateTimePicker)
    date_to="dtpTODATE",      # 입고일자 종료 (DateTimePicker)
    grid="dataGridView1",     # 그리드 (조회완료 대기용)
    query_x=38,               # 조회 버튼 x (toolStrip 좌표 폴백용)
    export_x=110,             # 엑셀출력 버튼 x (폴백용)
    query_wait=3.5,           # 조회 후 그리드 로딩 대기(초) — 데이터 많으면 늘리세요
    # iEMenu 자동 열기 (메인메뉴 > 품질관리 > 품질검사대상조회)
    menu_path=["품질관리"],            # 트리에서 클릭할 분류
    menu_program="품질검사대상조회",   # 우측 목록에서 더블클릭할 프로그램
)


def _connect_iemenu():
    """iEMenu 메뉴 창에 연결. 트리 탐색이 uia 기준이라 'uia 연결'을 최우선으로 한다.
    (uia 타이틀) → (Desktop으로 핸들 찾아 uia 핸들연결) → (win32 핸들연결) 순. 반환: WindowSpecification."""
    from pywinauto import Application, Desktop
    import re
    # 1) uia 타이틀 연결
    try:
        app = Application(backend="uia").connect(title_re=r"(?i).*iemenu.*", timeout=4)
        return app.window(title_re=r"(?i).*iemenu.*")
    except Exception:
        pass
    # 2) Desktop으로 iEMenu 핸들 찾기 (win32/uia 어느쪽이든)
    handle = None
    for be in ("win32", "uia"):
        try:
            for w in Desktop(backend=be).windows():
                try:
                    t = w.window_text() or ""
                except Exception:
                    continue
                if re.search("iemenu", t, re.I):
                    handle = getattr(w, "handle", None)
                    break
        except Exception:
            pass
        if handle:
            break
    # 3) 핸들로 연결 — uia 먼저(트리 탐색 위해), 안 되면 win32
    if handle:
        for be2 in ("uia", "win32"):
            try:
                return Application(backend=be2).connect(handle=handle, timeout=4).window(handle=handle)
            except Exception:
                continue
    raise SystemExit(
        "iEMenu 메뉴 창을 찾지 못했습니다.\n"
        "→ iERP의 iEMenu(메뉴 트리) 창이 떠 있는지 확인하세요.")


def _open_via_expand_all(tree_nodes, program_name):
    """iEMenu '전체 펼치기' → 왼쪽 트리에서 tree_nodes 더블클릭(예: 품질관리)
    → 오른쪽 목록에서 program_name(예: 품질검사대상조회) 더블클릭해 연다."""
    import time
    menu = _connect_iemenu()
    try:
        menu.set_focus()
    except Exception:
        pass
    try:                                   # 오른쪽을 '선택메뉴' 탭으로 (프로그램 목록 보이게)
        menu.child_window(title="선택메뉴", control_type="TabItem").click_input()
        time.sleep(0.4)
    except Exception:
        pass
    try:                                   # 전체 펼치기
        menu.child_window(title="전체 펼치기", control_type="Button").click_input()
        time.sleep(0.8)
    except Exception:
        pass
    tree = menu
    try:
        tree = menu.child_window(auto_id="mainTreeView", control_type="Tree")
        tree.set_focus()
        ie.send_keys("^{HOME}")            # 트리 맨 위로
        time.sleep(0.2)
    except Exception:
        pass
    # 1) 왼쪽 트리에서 분류 더블클릭 (품질관리) → 오른쪽 목록 채워짐
    for name in tree_nodes:
        item = tree.child_window(title=name, control_type="TreeItem")
        try:
            item.scroll_into_view()
        except Exception:
            pass
        item.double_click_input()
        time.sleep(1.0)
    # 2) 오른쪽 목록에서 프로그램 더블클릭 (품질검사대상조회) — 트리 항목 아님
    prog = None
    for ct in ("DataItem", "ListItem", "Custom", "Text", "TreeItem"):
        try:
            cand = menu.child_window(title=program_name, control_type=ct)
            if cand.exists(timeout=1):
                prog = cand
                break
        except Exception:
            continue
    if prog is None:
        prog = menu.child_window(title=program_name)   # 타입 무관 폴백
    try:
        prog.scroll_into_view()
    except Exception:
        pass
    prog.double_click_input()
    time.sleep(2.0)


def ensure_qc_open():
    """품질검사대상조회가 열려 있으면 그 창을, 없으면 iEMenu '전체 펼치기' 후
    품질관리(트리) 더블클릭 → 품질검사대상조회(우측목록) 더블클릭해 자동으로 연다."""
    try:
        return _find_window(RECEIPT_WIN, "품질검사대상조회")
    except SystemExit:
        pass
    if ie is None:
        raise SystemExit("품질검사대상조회 창이 없습니다 → iERP에서 직접 띄워주세요.")
    tree_nodes = list(RECEIPT_CFG.get("menu_path") or ["품질관리"])
    program = RECEIPT_CFG.get("menu_program") or "품질검사대상조회"
    program_id = RECEIPT_CFG.get("program_id")
    # 1순위: F9 실행창에 프로그램 ID 입력(빠르고 정확)
    if program_id and hasattr(ie, "open_via_f9"):
        print(f"  품질검사대상조회 자동 열기(F9): {program_id}")
        try:
            ie.open_via_f9(program_id)
            return _find_window(RECEIPT_WIN, "품질검사대상조회")
        except SystemExit:
            print("  F9 열기 실패 → 트리 메뉴 탐색으로 폴백")
        except Exception as e:
            print(f"  F9 열기 예외({type(e).__name__}) → 트리 메뉴 탐색으로 폴백")
    # 2순위(폴백): iEMenu '전체 펼치기' 트리 탐색
    print(f"  품질검사대상조회 자동 열기: 전체펼치기 → {' → '.join(tree_nodes)}(트리) → {program}(목록) 더블클릭")
    try:
        _open_via_expand_all(tree_nodes, program)
    except SystemExit:
        raise
    except Exception as e:
        raise SystemExit(
            "품질검사대상조회를 자동으로 열지 못했습니다.\n"
            f"  (원인: {type(e).__name__})\n"
            "→ iERP iEMenu에서 '품질검사대상조회'를 직접 더블클릭해 띄운 뒤, 다시 ①을 누르세요.\n"
            "  (창이 떠 있으면 그 창으로 바로 진행됩니다.)")
    try:
        return _find_window(RECEIPT_WIN, "품질검사대상조회")
    except SystemExit:
        raise SystemExit(
            "자동 열기를 시도했지만 품질검사대상조회 창을 찾지 못했습니다.\n"
            "→ iERP에서 품질검사대상조회를 직접 띄운 뒤 다시 ①을 누르세요.")


def _click_toolbar(win, ctl, title, x_offset) -> bool:
    """툴바 버튼 클릭: uia(핸들연결) title 우선, 실패 시 toolStrip1 좌표(x_offset)."""
    h = getattr(win, "handle", None)
    if h is not None:
        try:
            from pywinauto import Application
            app = Application(backend="uia").connect(handle=h)
            app.window(handle=h).child_window(title=title, control_type="Button").click_input()
            return True
        except Exception:
            pass
    ts = ctl.get("toolStrip1")
    if ts is not None:
        try:
            r = ts.rectangle()
            ts.click_input(coords=(x_offset, (r.bottom - r.top) // 2))
            return True
        except Exception:
            pass
    return False


def _set_date_ctl(ctl, ymd: str):
    """날짜 컨트롤(DateTimePicker/Edit)에 YYYY-MM-DD 설정."""
    if ctl is None or not ymd:
        return
    y, m, d = (ymd.split("-") + ["", "", ""])[:3]
    try:
        ctl.set_time(year=int(y), month=int(m), day=int(d))   # win32 DateTimePicker
        return
    except Exception:
        pass
    try:
        _set_edit(ctl, ymd)
        return
    except Exception:
        pass
    # 키보드: DateTimePicker(yyyy-MM-dd)에 yyyymmdd 입력 → 세그먼트 자동 이동
    try:
        try:
            ctl.click_input()
        except Exception:
            ctl.set_focus()
        from pywinauto.keyboard import send_keys
        send_keys(y + m.zfill(2) + d.zfill(2))
    except Exception:
        pass


def _save_ierp_excel(out_path, L, timeout: float = 45) -> bool:
    """iERP 엑셀출력으로 열린 엑셀 워크북을 찾아 out_path 로 저장. 진단을 L()로 남긴다.
    미저장 워크북 우선, 없으면 데이터 있는 워크북(저장본 포함)도 후보로 본다."""
    out_path = Path(out_path)
    try:
        import win32com.client as win32
        import pythoncom
    except Exception as e:
        L(f"  win32com 사용 불가: {e}")
        return False
    out_path.parent.mkdir(parents=True, exist_ok=True)
    import time
    end = time.time() + timeout
    while time.time() < end:
        try:
            apps = ie._excel_apps(win32, pythoncom)
        except Exception as e:
            L(f"  Excel 인스턴스 열거 실패: {e}")
            apps = []
        L(f"  Excel 인스턴스 {len(apps)}개")
        best, best_score = None, -1
        for xl in apps:
            try:
                cnt = xl.Workbooks.Count
            except Exception as e:
                L(f"   Workbooks 접근 실패: {e}")
                continue
            for i in range(cnt, 0, -1):
                try:
                    wb = xl.Workbooks(i)
                    used = wb.Worksheets(1).UsedRange
                    rows, cols = used.Rows.Count, used.Columns.Count
                    has_path = bool(wb.Path)
                    L(f"   WB '{wb.Name}' 저장됨={'Y' if has_path else 'N'} {rows}x{cols}")
                    if rows < 2 and cols < 2:
                        continue
                    score = (0 if has_path else 10 ** 7) + rows * cols
                    if score > best_score:
                        best, best_score = (xl, wb), score
                except Exception as e:
                    L(f"   WB 검사 실패: {e}")
        if best is not None:
            xl, wb = best
            try:
                xl.DisplayAlerts = False
                wb.SaveAs(str(out_path), FileFormat=51)
                L(f"  저장 성공: {out_path}")
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
                return True
            except Exception as e:
                L(f"  SaveAs 실패: {e}")
        time.sleep(1.0)
    L("  시간초과 — 저장할 엑셀 워크북을 못 찾음")
    return False


def export_receipt(date_from: str, date_to: str | None = None, out_path=None):
    """품질검사대상조회를 입고일자 기간으로 조회(검사파트=자재파트) → 엑셀출력 → .xlsx 저장.
    반환: 저장된 xlsx Path. (Windows + pywinauto + pywin32, iERP 로그인 + 화면 띄움)"""
    if ie is None:
        raise SystemExit("ierp_export 로드 실패 — Windows + pywinauto 환경에서 실행하세요.")
    cfg = RECEIPT_CFG
    if not (cfg.get("date_from") and cfg.get("part_combo")):
        raise SystemExit(
            "품질검사대상조회 자동조회는 정찰(Phase)이 필요합니다.\n"
            "  iERP에서 품질검사대상조회 화면을 띄우고:\n"
            '     py inspect_ierp.py "품질검사대상조회"\n'
            "  결과의 auto_id로 RECEIPT_CFG(part_combo·date_from·date_to)를 채우세요.\n"
            "  (당장은 화면에서 직접 엑셀출력 → '엑셀 직접 선택' 경로를 쓰세요.)")
    import time
    win, backend = ensure_qc_open()        # 없으면 iEMenu로 자동 열기
    log = []
    def L(m):
        print(m); log.append(str(m))

    # 컨트롤(auto_id)이 보이는 백엔드 보장 (win32에서 0개로 안 보이는 PC 대응)
    win, backend, ctl = _ensure_controls(win, backend)

    L(f"품질검사대상조회 사용 backend={backend}, 컨트롤 {len(ctl)}개")
    try:
        win.set_focus()
    except Exception:
        pass
    time.sleep(0.3)

    # 검사파트 = 자재파트
    part = ctl.get(cfg["part_combo"])
    if part is not None:
        L("검사파트 = 자재파트 설정 시도")
        try:
            part.select(cfg["part_value"]); L("  검사파트 select OK")
        except Exception as e:
            L(f"  select 실패({e}) → 텍스트입력 시도")
            try:
                _set_edit(part, cfg["part_value"]); L("  텍스트입력 OK")
            except Exception as e2:
                L(f"  텍스트입력도 실패({e2})")
        time.sleep(0.3)
    else:
        L(f"※ 검사파트 콤보({cfg['part_combo']}) 못 찾음")

    # 입고일자 기간
    L(f"입고일자 {date_from} ~ {date_to or date_from} 설정")
    df_ctl = ctl.get(cfg["date_from"]); dt_ctl = ctl.get(cfg.get("date_to"))
    L(f"  날짜컨트롤 from={'O' if df_ctl else 'X'} to={'O' if dt_ctl else 'X'}")
    _set_date_ctl(df_ctl, date_from)
    _set_date_ctl(dt_ctl, date_to or date_from)
    time.sleep(0.4)

    # 조회
    L("조회 버튼 클릭")
    if not _click_toolbar(win, ctl, "조회", cfg["query_x"]):
        L("  ※ 조회 버튼 클릭 실패")
    try:
        ie._wait_query_done(win, cfg.get("grid", "dataGridView1"))
    except Exception:
        time.sleep(cfg.get("query_wait", 3.5))
    time.sleep(1.0)
    try:
        g = ctl.get(cfg.get("grid", "dataGridView1"))
        L(f"  조회 후 그리드 행수: {g.item_count() if g is not None else '컨트롤없음'}")
    except Exception as e:
        L(f"  그리드 행수 확인 불가: {e}")

    # 엑셀출력
    L("엑셀출력 버튼 클릭")
    if not _click_toolbar(win, ctl, "엑셀출력", cfg["export_x"]):
        L("  ※ 엑셀출력 버튼 클릭 실패")
    time.sleep(1.0)

    out = Path(out_path) if out_path else (ie.EXPORT_DIR / f"receipt_{ie._ts()}.xlsx")
    L("엑셀 저장 시도")
    ok = _save_ierp_excel(out, L, timeout=45)

    logfile = out.parent / "export_log.txt"
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        logfile.write_text("\n".join(log), encoding="utf-8")
    except Exception:
        logfile = None
    if not ok:
        raise SystemExit(
            "엑셀 자동저장 실패.\n"
            + (f"진단로그: {logfile}\n" if logfile else "")
            + "→ 또는 iERP에서 직접 엑셀출력·저장 후 '엑셀 직접 선택' 모드로 진행하세요.")
    L(f"엑셀 저장 완료: {out}")
    try:
        if logfile:
            logfile.write_text("\n".join(log), encoding="utf-8")
    except Exception:
        pass
    return out


# ── (1) 엑셀출력 결과 파싱 (오프라인 가능) ─────────────────────────────────────
def _find_header_indexes(header: list[str]) -> dict:
    idx = {}
    norm = [str(h or "").replace(" ", "").strip() for h in header]
    for key, aliases in COLMAP.items():
        for a in aliases:
            a2 = a.replace(" ", "")
            matches = [i for i, h in enumerate(norm) if h == a2 or (a2 and a2 in h)]
            if matches:
                idx[key] = matches[-1] if key in PREFER_LAST else matches[0]
                break
    return idx


def load_receipt_excel(path: str | Path, header_row: int = 1) -> list[dict]:
    """부자재 입고 엑셀(.xlsx)을 읽어 행 dict 목록으로. (openpyxl)
    반환 dict 키: code,name,customer,category,supply,first_in,mat_in"""
    import openpyxl
    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header = None
    for _ in range(header_row):
        header = next(rows_iter, None)
    if not header:
        return []
    idx = _find_header_indexes(list(header))
    out = []
    for raw in rows_iter:
        if raw is None or all(c is None or str(c).strip() == "" for c in raw):
            continue
        def g(key):
            i = idx.get(key)
            return raw[i] if (i is not None and i < len(raw)) else None
        rec = {
            "code": str(g("code") or "").strip(),
            "name": str(g("name") or "").strip(),
            "customer": str(g("customer") or "").strip(),
            "category": str(g("category") or "").strip(),
            "supply": str(g("supply") or "").strip(),
            "first_in": g("first_in"),
            "mat_in": g("mat_in"),
        }
        if rec["code"] or rec["name"]:
            out.append(rec)
    return out


if __name__ == "__main__":
    # 오프라인 자가점검: 초도분 판정 로직
    sample = [
        {"code": "151201133701", "name": "A 단상자", "first_in": "2026-06-19", "mat_in": "2026-06-19"},
        {"code": "100701055701", "name": "B 라벨",   "first_in": "20260619",   "mat_in": "2026.06.19"},
        {"code": "111801001701", "name": "C 용기",   "first_in": "2026-06-18", "mat_in": "2026-06-19"},
    ]
    keep = chodo_rows(sample)
    print(f"입력 {len(sample)}건 → 초도분 {len(keep)}건")
    for r in keep:
        print("  초도분:", r["code"], r["name"])
