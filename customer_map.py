"""
품목코드 → 고객사/품목형태 도출.

9자리 품목코드 = 거래처(4) + 카테고리(2) + 일련(3)
12자리 파생품   = 9자리 + 유형코드 3자리 (첫 글자가 형태)

⚠️ 공개 저장소 버전에서는 거래처명 데이터(CUSTOMER_BY_CODE)가 비어 있다(기밀).
   실제 거래처명 매핑이 든 customer_map.py 는 secret_file/ 폴더에 있으며,
   이 파일 위치(프로그램 폴더)에 덮어쓰면 고객사명이 표시된다.
   비어 있어도 customer_from_code() 는 (코드4자리, matched=False) 로 정상 동작한다.
"""
from __future__ import annotations

# 거래처 4자리 → 고객사명.  ※ 기밀: 공개본은 비움. 실데이터는 secret_file/customer_map.py
CUSTOMER_BY_CODE: dict[str, str] = {}

# 파생 접미사 첫 글자 → 품목형태
FORM_BY_SUFFIX = {
    "Z": "반제품",
    "1": "반제품",
    "7": "부자재",
    "5": "가공부자재",
    "Y": "부업",
}
SUBSIDIARY_SUFFIX = ("7", "5")  # 부자재 계열


def _clean(code) -> str:
    return str(code or "").strip()


def customer_from_code(code) -> tuple[str, bool]:
    """앞 4자리로 고객사명을 찾는다. (이름, 매칭여부) 반환.
    미등록 코드면 (코드4자리, False)."""
    c = _clean(code)
    key = c[:4]
    if key in CUSTOMER_BY_CODE:
        return CUSTOMER_BY_CODE[key], True
    return key, False


def form_from_code(code) -> str:
    """품목형태(완제품/반제품/부자재/가공부자재/부업). 코드만으로 판별.
    ※ 웹의 '유형'(제형: 하이드로겔/시트 등)과는 다름 — 화면 컬럼이 있으면 그쪽 우선."""
    c = _clean(code)
    if len(c) <= 9:
        return "완제품"
    suffix_first = c[9:10]
    return FORM_BY_SUFFIX.get(suffix_first, "")


def is_subsidiary(code) -> bool:
    """부자재 계열(접미사 7=부자재, 5=가공부자재)인지."""
    c = _clean(code)
    return len(c) >= 12 and c[9:10] in SUBSIDIARY_SUFFIX


if __name__ == "__main__":
    for t in ["151201080", "151201133701", "100701055Z01", "1118010015"]:
        print(t, "→", customer_from_code(t), "/", form_from_code(t), "/ 부자재:", is_subsidiary(t))
