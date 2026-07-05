# iERP 초도분 → 표준견본 등록기

부자재 입고 데이터에서 **초도분**(초도입고일 == 자재입고일)만 골라
표준견본 웹앱용 `register.json` 을 만든다.

## 파일
- `customer_map.py` — 품목코드 앞4자리 → 고객사, 접미사 → 품목형태(부자재 등)
- `sample_export.py` — 초도분 판정/엑셀파싱(오프라인) + 화면 자동화 골격(우클릭, Windows)
- `sample_register.py` — 위를 묶어 `register.json` 생성 (CLI)

## 확정된 방식: 품목재고조회(ITEMINVv1)의 '초도입고' 값 사용
초도입고일은 제품마다 iERP에 저장된 고유값이다(품질검사대상조회에서 품목 우클릭 → **품목재고조회** 창의 `초도입고` 필드).
→ 입고검사 엑셀의 각 품목코드를 품목재고조회에 넣어 **초도입고일**을 읽고,
   `초도입고 == 입고일(자재입고일)`이면 그 입고가 첫 입고 = **초도분** → 등록.

### 흐름
1. iERP **품질검사대상조회**에서 날짜 조회 → **엑셀출력** (그날 입고분).
2. **품목재고조회 창을 띄워둔다** (아무 품목 우클릭 → 품목재고조회 한 번 열기).
3. 실행:
   ```
   py sample_register.py --excel <받은.xlsx> --lookup-firstin --out register.json
   py sample_register.py --excel <받은.xlsx> --lookup-firstin --subsidiary-only   # 부자재(7/5)만
   ```
   - `--lookup-firstin` : 품목재고조회에서 코드별 초도입고를 읽어옴(조회 버튼 자동 클릭).
   - 돌아가는 동안 마우스·키보드 건드리지 말 것. 품목재고조회 창은 열어둘 것.
4. 표준견본 웹앱(로컬)에서 **⇪ iERP 등록 가져오기** → `register.json` 업로드.
   - 이미 등록된 품목코드는 자동 제외, 신규만 '등록' 상태로 추가됨.

> 참고 컨트롤 ID(정찰 확정): 품목코드 `txtMITEM` · 초도입고 `txtMREFD` · 품목명 `txtMPNAM` · 조회는 toolStrip1 좌측.
> 조회 버튼 클릭 좌표가 안 맞으면 `sample_export.py` 의 `QUERY_BTN_X_OFFSET` 조정.

### (대안) 초도입고일 컬럼이 있는 엑셀이면
`--lookup-firstin` 없이 그대로 쓰면 됨(엑셀의 초도입고일 사용). `--on/--from/--to`로 초도입고일 날짜 필터 가능.

### 자동 매핑(실제 양식 기준)
| 웹 필드 | iERP 컬럼 |
|---|---|
| 고객사 | 고객 |
| 품목코드 | 품목코드 |
| 품목명 | 품명 |
| 유형 | 관리유형(단상자/파우치 등) |
| 사·자급 | 입고유형(사급반입→사급, 그 외→자급) |
| 등록일자 | 입고일시(자재입고일) |

> ⚠️ `register.json` 은 실데이터다 — 공개 저장소/배포에 올리지 말 것(로컬에서만).

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

## 의존성
`py -m pip install openpyxl pywinauto pywin32`  (openpyxl만이면 엑셀 경로는 동작)
