# iERP 초도분 → 표준견본 등록기

iERP 부자재 입고 데이터에서 **초도분**(초도입고일 == 자재입고일)만 골라
표준견본 웹앱용 `register.json` 을 만들어 자동 등록한다.

> ⚠️ Windows 전용. iERP 화면 자동화(`pywinauto`)를 쓰므로 Windows에서만 동작한다.

---

## 빠른 시작 (실무자용 · GUI)

가장 쉬운 방법. 명령어 없이 창에서 클릭으로 처리한다.

1. iERP 로그인 → **품질검사대상조회**에서 날짜 조회 → **엑셀출력**.
2. **품목재고조회 창을 하나 띄워둔다** (아무 품목 우클릭 → 품목재고조회).
3. `등록기.bat` **더블클릭** (또는 `pyw registrar_gui.py`).
4. 창에서 받은 엑셀을 지정하고 **① 엑셀에서 초도분 추출 → ② 웹앱에 등록** 순서로 버튼 클릭.
   - 돌아가는 동안 마우스·키보드를 건드리지 말 것. 품목재고조회 창은 열어둘 것.
   - `register.json` 은 웹앱 폴더(`index.html` 있는 곳, 기본 `C:\sample-list`)에 저장돼야 자동 등록된다.

### exe로 배포하려면
Python 없는 PC에 배포할 실행파일을 만들려면:
```
build_registrar_exe.bat   더블클릭  (인터넷 필요, 수 분 소요)
```
→ `dist\부자재표준견본등록기.exe` 생성.

---

## 명령줄 사용법 (고급 · CLI)

`sample_register.py` 로 직접 `register.json` 을 만든다.

```
# 초도입고를 iERP 품목재고조회에서 자동으로 읽어와 등록 (iERP 화면 띄운 상태)
py sample_register.py --excel <받은.xlsx> --lookup-firstin --out register.json

# 부자재 계열(접미사 7/5)만
py sample_register.py --excel <받은.xlsx> --lookup-firstin --subsidiary-only --out register.json
```

### 옵션
| 옵션 | 설명 |
|---|---|
| `--excel <path>` | 입고 엑셀출력(.xlsx) 입력 |
| `--screen` | 엑셀 대신 화면 자동화(우클릭) 경로 사용 (정찰 완료 필요) |
| `--lookup-firstin` | 품목재고조회(ITEMINVv1) 창에서 코드별 초도입고일을 읽어와 채움 |
| `--subsidiary-only` | 부자재 계열(접미사 7/5)만 등록 |
| `--on YYYY-MM-DD` | 특정 초도입고일 하루만 (= `--from X --to X`) |
| `--from` / `--to` | 초도입고일 날짜 범위 필터 |
| `--out <path>` | 출력 `register.json` 경로 |
| `--min-as-first` | [비권장] 초도입고일을 입고이력 최초일로 추정(부정확) |

### (대안) 초도입고일 컬럼이 이미 있는 엑셀
`--lookup-firstin` 없이 그대로 쓰면 엑셀의 초도입고일을 사용한다.
`--on/--from/--to` 로 날짜 필터 가능.

마지막으로 표준견본 웹앱(로컬)에서 **⇪ iERP 등록 가져오기** → `register.json` 업로드.
이미 등록된 품목코드는 자동 제외, 신규만 '등록' 상태로 추가된다.

---

## 판정 방식: 품목재고조회(ITEMINVv1)의 '초도입고' 값

초도입고일은 제품마다 iERP에 저장된 고유값이다(품질검사대상조회에서 품목 우클릭 →
**품목재고조회** 창의 `초도입고` 필드).
→ 입고검사 엑셀의 각 품목코드를 품목재고조회에 넣어 **초도입고일**을 읽고,
`초도입고 == 입고일(자재입고일)` 이면 그 입고가 첫 입고 = **초도분** → 등록.

> 참고 컨트롤 ID(정찰 확정): 품목코드 `txtMITEM` · 초도입고 `txtMREFD` · 품목명 `txtMPNAM` · 조회는 toolStrip1 좌측.
> 조회 버튼 클릭 좌표가 안 맞으면 `sample_export.py` 의 `QUERY_BTN_X_OFFSET` 조정.

### 자동 매핑 (실제 양식 기준)
| 웹 필드 | iERP 컬럼 |
|---|---|
| 고객사 | 고객 |
| 품목코드 | 품목코드 |
| 품목명 | 품명 |
| 유형 | 관리유형(단상자/파우치 등) |
| 사·자급 | 입고유형(사급반입→사급, 그 외→자급) |
| 등록일자 | 입고일시(자재입고일) |

---

## 파일 구성
| 파일 | 역할 |
|---|---|
| `registrar_gui.py` | GUI 등록기 (실무자용 진입점) |
| `등록기.bat` | GUI 실행 배치 |
| `build_registrar_exe.bat` | exe 빌드 스크립트 |
| `sample_register.py` | CLI — 초도분 판정 후 `register.json` 생성 |
| `sample_export.py` | 초도분 판정/엑셀파싱 + iERP 화면 자동화 골격 |
| `ierp_export.py` | iERP 조회 결과 엑셀 추출 유틸 |
| `customer_map.py` | 품목코드 앞4자리 → 고객사, 접미사 → 품목형태 |

> ⚠️ `register.json` 은 실데이터다 — 공개 저장소/배포에 올리지 말 것(로컬에서만).
> 실제 거래처명 매핑이 든 `customer_map.py` 는 `secret_file/` 에 있으며, 공개본은 비어 있다.

---

## 화면 자동화(우클릭) 경로 — 정찰(Phase 0) 후 사용
엑셀출력에 두 날짜가 없고 **행 우클릭으로만** 확인되는 화면이면 자동화가 필요하다.
정찰로 컨트롤 ID를 확보해 `sample_export.py` 의 `RECEIPT` dict 를 채운다:
```
# iERP 로그인 + 부자재 입고 화면 띄운 상태에서 (관리자 PowerShell)
py inspect_ierp.py "입고"
```
확보할 값: 창제목/메뉴경로, 조회 날짜 입력 auto_id, 그리드 auto_id,
우클릭 컨텍스트 메뉴 항목명, 초도입고일/자재입고일 필드. 채운 뒤:
```
py sample_register.py --screen --from 2026-06-19 [--to 2026-06-19]
```
`extract_via_rightclick()` 의 `# TODO(정찰)` 부분(행 셀 읽기/우클릭 팝업 읽기)을 화면에 맞게 확정.

---

## 의존성
```
py -m pip install openpyxl pywinauto pywin32
```
(엑셀 경로만 쓸 거면 `openpyxl` 만으로도 동작. GUI 테마는 `sv-ttk` 추가.)
