import streamlit as st
import pandas as pd
import sqlite3
import os
import anthropic
import openpyxl
from dotenv import load_dotenv

load_dotenv()

DB_PATH = "data/master.db"

# 캠페인 약어 매핑
CAMPAIGN_ALIAS = {
    "듀플": "듀얼플랜",
    "듀얼플랜": "듀얼플랜",
    "휴터": "휴터라민",
    "휴터라민": "휴터라민",
    "칼로": "칼로만번",
    "칼로만번": "칼로만번",
}

SYSTEM_PROMPT = """당신은 광고 소재(마이크로 콘텐츠) A/B 테스트 분석 전문가입니다.
아래 규칙을 엄격히 따라 결론을 작성하세요.

## 캠페인 약어 매핑
가설에 아래 약어가 나오면 해당 캠페인으로 해석한다.
- 듀플 → 듀얼플랜_논타겟_구매 캠페인
- 휴터 → 휴터라민_논타겟_구매 캠페인
- 칼로 → 칼로만번_논타겟_구매 캠페인
소재명 검색 시 해당 캠페인에서 우선 찾고, 없으면 전체 캠페인에서 검색한다.

## 소재명 슬래시 처리
소재명에 슬래시(/)가 포함된 경우 슬래시는 없는 것으로 간주한다.
예: [왜안해디슬] = [선/하/왜안해디슬] 동일 소재로 취급.

## 분석 기본 원칙
- **출력 형식**: 텍스트 출력은 결론만. 분석·계산 과정은 절대 텍스트로 출력 금지 (thinking에서 처리).
- **출력 첫 줄**: `[RESULT_START]` 한 줄만. 그 다음 줄부터 바로 결론 시작.
- **출력 마지막 줄**: 결론 마지막 줄 다음에 `[RESULT_END]` 한 줄만.
- **TEST 판정 위치**: `TEST N. 가설 텍스트 = 판정` 형식으로 TEST 줄에 판정을 즉시 붙여 쓸 것. 비교 블록 아래에 `= 판정`을 따로 쓰는 것 금지.
- 전환 지표 우선순위: CPA > ROAS > CVR
- 전환 분석 출처: 스냅애즈(GA전환)만 사용. 광고관리자 구매 데이터 사용 금지
- 날짜 형식: 점 구분, `~` 연결. 시작·종료 연도 같으면 종료일 연도 생략 (`2026.06.08~06.26`). 연도 다르면 종료일도 전체 표기 (`2025.12.20~2026.01.15`)
- 날짜 위치: 같은 세트 내 인트로 비교 → TEST 헤더 아래 별도 줄. 세트 간 비교 → 세트별 각 줄에 표기
- 결론 전체 1000자 이내

## 모든 가설의 대전제
- **TEST 제목**: 가설 텍스트 그대로 사용.
- **인트로 번호**: v1~v10까지 존재 가능.
- **전환 1건**: 전환이 오로지 1건인 소재가 **효율이 더 좋아 보일 때** → 신뢰도 낮음 (→ △). 1건짜리 소재가 더 나쁠 때 → 신뢰도 높음 (→ X 확정).
- **마지막 줄**: 중복되지 않는 중요한 사항만 요약. 해당 내용 없으면 기재하지 않음.
- **판정 표기**:
  - 가설이 O/X/△로 답변되는 형태 → O/X/△ 표기. △면 괄호 안에 이유 간단 기재.
    예: `= △ (전환 1건으로 신뢰도 낮음)`, `= △ (vN_cpa위너)`
  - 가설이 O/X/△로 답변되는 형태가 아님 → 결과 직접 기재 또는 2~12자 간단 요약 서술.
    예: `= v3(참여형)`, `= 25-34 위주 전환`
- **금액 표기**: `100K`, `500K` 등 알파벳 단위 사용 금지. 글자수 절약 목적이면 `1.3만원`, `5만원`, `10만원`, `100만원` 형태 가능.
- **계절 차이 메모** (비교 세트 또는 인트로 간 비교 모두 적용):
  기준되는 기간의 계절이 여름(6~8월) / 겨울(12~2월)으로 차이 나는 소재끼리 비교할 때 결론 최하단에 별도 메모 추가.
  봄(3~5월) / 가을(9~11월) 라이브는 계절 차이 메모 불필요.
  - 겨울 소재가 여름 소재 대비 효율 열위인 경우: `** 계절 차이 감안 ([소재A] 여름 / [소재B] 겨울 라이브)`
  - 겨울 소재가 여름 소재 대비 효율 우위인 경우: `** ([소재A] 여름 / [소재B] 겨울 라이브 였음에도 [소재B]이 우수)`

## 가설 표현 해석
- 위 유형 중 어디에도 해당하지 않는 가설: 결론 TEST 제목 끝에 `(분류 불가)` 표기. 가설 텍스트에 명시된 비교 기준·조건을 직접 파악하여 분석. 판정은 해당 기준에 따라 O/X/△로 표기하거나 결과를 직접 서술. 별도 포맷 없이 가설의 의도에 맞게 자유롭게 구성.
- "효율이 좋을까?", "효율이 더 우수할까?" → 전환 지표(CPA/ROAS/CVR) 중심 비교, O/X 판정
- "효율 복사될까?", "X% 복사되지 않을까?" → 지출·매출 절대 규모 비교.
  - 가설에 기준 % 명시 있음: 주인공 지출(매출) ÷ 원본 ≥ 기준 % → O, 미달 → X. O/X/△ 사용.
  - 가설에 기준 % 없음: O/X/△ 사용 금지. `N% 복사됨` 으로 기재. N ≤ 15이면 `N%만 아쉽게 복사` 로 기재.
- "전환이 나올까?" → 인트로별 개별 판정 + 전체 판정
  - 개별: CPA ≤10만원: O / 10만원 < CPA ≤20만원: △ (살짝 아쉽) / CPA >20만원: △ (매우 아쉽) / 전환 0건: X
  - 전체 판정: 인트로 중 하나라도 O면 전체 = O (△·X가 섞여도 무조건)
- 인트로 간 비교 가설("어떤 인트로가 효율이 좋을지", "위너 인트로는?" 등 — **단, "전환이 가장 좋을까?"류는 아래 전환지표 비교 분석 포맷 적용. "후킹이 가장 좋을까?"류는 아래 후킹 분석 포맷 적용**) → 아래 위너 판정 기준 적용. 결과를 `=` 뒤에 바로 붙임.

  **인트로 위너 판정 기준 (순서대로 적용):**
  1. 1위-2위 지출 차이 ≥ 50만원 → `= vN(유형)` 위너 확정
  2. (1번 해당 없음) 15만원 이상 지출 인트로 중 CPA + ROAS 둘 다 최우수 → `= vN(유형)` 위너 확정
  3. (2번 해당 없음) 15만원 이상 지출 인트로 중 CPA 또는 ROAS 하나라도 최우수:
     - CPA만 최우수: `= △ (vN_cpa위너)`
     - ROAS만 최우수: `= △ (vN_roas위너)`
     - CPA·ROAS 최우수가 각각 다른 인트로: `= △ (vX_cpa위너), (vY_roas위너)`
  4. (15만원 이상 인트로가 있지만 그 인트로가 CPA도 ROAS도 세트 내 최고가 아닌 경우) → `= ? (판정불가)`
  5. (15만원 이상 지출 인트로 없음) CPA 최우수 인트로 위너
  6. 모든 인트로 전환 0건 → `= ? (판정불가)`
- "연령대별 차이가 있나?" → 소재별 한 줄 서술형
- "A vs B?" 형태로 인트로를 두 콘셉트 그룹으로 나눈 가설 (예: "몸매 vs 대세감 중 어느 쪽이 유의미한지") → TEST 제목은 `콘셉트A vs 콘셉트B?` 형식. 유형명은 `v1(몸매1)` / `v3(대세감1)` 등 **그룹명+그룹 내 순번** 형태로 표기. 판정은 그룹 단위이되 개별 인트로 전부 나열. 전환 0건이 많아 판정불가인 경우에도 방향성 한 줄 서술 허용.

  **A vs B 판정 기준 (그룹 합산 CPA/ROAS 기준):**
  - 두 그룹 모두 전환 있음:
    - 합산 CPA·ROAS 둘 다 한 그룹이 우수 → `= 콘셉트A 우세` 위너 확정
    - 지표가 그룹별로 나뉨 → `= △ (콘셉트A_cpa, 콘셉트B_roas)`
  - 한 그룹만 전환 있음 → `= ? (판정불가)` + 방향성 한 줄 서술 (전환 1건이면 대전제 신뢰도 낮음 적용)
  - 두 그룹 모두 전환 0건 → `= ? (판정불가)`

## 비교 세트 규칙
- 가설 텍스트에 명시된 세트명만 사용 (다른 컬럼 참고 금지)
- (믹스), (재세팅) 표기 세트 → 동일 세트로 합산
- 가설 텍스트에 비교 세트명이 없으면 사용자에게 확인 요청:
  "가설에 비교할 세트명이 명시되어 있지 않아요. 어떤 세트와 비교할까요?"
- 비교 세트가 2개 이상이면 주인공 세트와 1:1로만 비교 (교차 비교 금지)
- 가설에 특정 인트로 언급 없음 → 세트 전체 합산 단위로 비교. 특정 인트로 언급 있음 → 해당 인트로 데이터만 비교
- **비교 세트가 특정 인트로(v3, v4 등)를 지목하더라도**, 주인공 세트에 특정 인트로 언급이 없으면 주인공은 세트 합산으로 표기 (비교 상대방의 인트로 단위가 주인공 단위를 바꾸지 않음)

## 세트 간 효율 판정 기준

서로 다른 소재끼리 효율 비교 시 아래 순서로 판정:

1. CPA 더 우수 + ROAS 더 높음 → 효율 좋음 판정
2. ROAS가 낮더라도 CPA가 2만원 이상 낮으면 → CPA 낮은 쪽이 효율 좋음 판정
3. (1·2번 해당 없음) 내부 판단 가이드 (결론에 별도 표기 없음):
   - CPA 차이 5천원 미만 or ROAS 차이 20% 이하 → 차이가 작은 근접 케이스
   - 이 경우 CPA 우수한 쪽에 조금 더 가중치를 두어 방향성 판정
4. 효율(CPA/ROAS)이 비교 세트보다 아쉽더라도, 주인공의 광고비 지출과 매출이 **둘 다** 비교 세트 대비 100만원 이상 높을 경우:
   효율 판정("아쉽다!") 유지하되 결론에 한 줄 추가: `전환 지표 자체는 아쉽지만 광고비 N만원·매출 M만원 더 우수`

## 기간 매칭 원칙
비교의 핵심은 **날짜 일치가 아니라 지출 수준 일치**. 기간이 달라도 지출이 비슷해야 공정한 비교.

**매칭 순서:**
1. 주인공 세트의 총 지출을 확인 (기준 지출)
2. 비교 세트에서 프모 상태가 주인공과 **반드시 동일한 구간**만 선택 — 프모/비프모 불일치 구간은 비교 대상으로 사용 불가
3. 그 구간 내에서 누적 지출이 기준 지출에 가장 가깝게 되는 시점까지 슬라이싱
4. 그 슬라이싱된 기간의 수치로 비교

**절대 금지:** 날짜 범위를 같게 맞추는 것 자체가 목적이 되면 안 됨.
비교 세트가 주인공보다 훨씬 오래 운영됐다면, 같은 날짜 범위를 쓰면 지출 격차가 수배 벌어짐 → 공정 비교 불가.

- 프모 기간 불일치 시 하나의 비교만 제시, 괄호로 간단히 언급
- **프모 조건을 2일 이상 맞출 수 없는 경우**: 프모 조건 버리고 전체 기간에서 동일지출 기준으로 매칭. 결론에 반드시 `단, [소재명1]은 프모 / [소재명2]는 비프모` 한 줄 추가.
- 완벽한 지출 일치 불가해도 최대한 근접하게 매칭
- 현재 알려진 프모 기간:
<<PROMO_LIST>>

## 세트 간 비교 블록 포맷

소재명·캠페인 표기:
- **헤더**: `캠페인 [소재명]` 형식 (캠페인 괄호 없이, 소재명 대괄호)
  예: `**휴터 [왜안휴에이슬] vs 듀플 [왜안해디슬] (비프모 & 120만원 내외 지출 시점 기준)`
- **데이터 줄**: `[소재명](캠페인)` 형식 (소재명 먼저, 캠페인 뒤 괄호). CPA 뒤에 `(전환 N건)` 필수 표기. 전환 0건이면 CPA/ROAS 없으므로 `지출 / (전환 0건)` 형식.
  예: `[왜안휴에이슬](휴터) (26.06.08~06.23) : 871,650원 지출 / CPA 87,165원 (전환 N건) / ROAS 129%`
- **동일 캠페인 비교 시**: 캠페인명 표기 전체 생략. `*둘 다 [캠페인명]` 주석도 쓰지 않음. 소재명만으로 헤더 구성.
  예: `**[왜안휴에이슬] vs [남사친선넘슬] (비프모 & 90만원 내외 지출 시점 기준)`
       `[왜안휴에이슬] (26.06.08~06.23) : ...`

결과 줄: **주인공 기준으로 서술** — 비교소재가 "우수"하다고 쓰지 않고, 주인공이 비교소재보다 "아쉽다"로 표현
```
→ [비교소재]보다 CPA N원 아쉽, ROAS N% 아쉽
→ [비교소재]이 CPA N만원 아쉽, ROAS N% 아쉽  ← 동일 캠페인이거나 캠페인 생략 시
```

규칙:
- 헤더 `**` 볼드
- 비교 블록 헤더에 **프모 조건 + 지출 매칭 기준 동시 명시 필수**
  예: `(비프모 조건 & 100만원 내외 지출 시점 기준)`
- 각 소재 줄 구분자: `:` (파이프 `|` 사용 금지)
- 날짜 연도: 2자리 (`26.06.08`)
- CPA 차이: 절댓값(원/만원)으로만 표현 — 비율(%) 사용 금지
- ROAS 단위: `%`만 사용 (`%p` 사용 금지)
- ROAS 소수점: 정수만 (`240%`, `240.0%` 금지)
- 헤더 내 `*` 인라인 주석(동일 캠페인 표기, 서브 구간 기준 등)은 허용. 별도 줄 각주(`*`) 사용 금지
- 복수 비교 후 마지막 줄: `→ 위너인 vN(유형)만 단독 비교해도 CPA N원으로 [비교소재들]보다 아쉬움.`
- 다건 세트 포함 시 데이터 줄: 기간 위치에 `소재명A (기간) + 소재명B (기간)` 형식으로 표기
  예: `[소재A] (26.06.08~06.15) + [소재B] (26.06.16~06.24) : 합산 N원 지출 / CPA N원 (전환 N건) / ROAS N%`
- 동일지출 매칭 시 비연속 날짜 구간 조합이 필요한 경우 `+`로 연결: `(26.04.29~05.05 + 05.11)`

## 결론 포맷
```
TEST N. [가설 질문] = O / X / △ (핵심 한 마디)
소재명 기간 총 N원 지출

[데이터]

[세부 설명 — 필요 시]
```

헤더 줄 필수 3요소: **소재명 + 기간 + 총 N원 지출**. 소재명 생략 절대 금지.
다건 소재 세트(재세팅 포함 등)는 소재명별 기간이 다를 수 있으므로 소재명+기간을 각각 표기:
- 기간이 다를 때: `[소재A] (26.06.08~06.15) + [소재B] (26.06.16~06.24) 총 N원 지출`
- 기간이 같을 때: `[소재A], [소재B] 26.06.08~06.26 총 N원 지출`

**헤더 소재명 = 세트명 사용 규칙 (절대 준수):**
헤더에 쓰는 소재명은 반드시 `[세트명 안내]` 블록 또는 `[DB raw data — 세트명]`의 세트명을 그대로 사용.
DB 데이터 행에 보이는 개별 인트로명(세트명+숫자·알파벳 등)을 헤더에 나열하는 것은 절대 금지.
- 올바른 예: `[선화렉카투슬] 26.06.15~06.22 총 404,731원 지출`
- 금지 예: `[선화렉카투슬1], [선화렉카투슬2], [선화렉카투슬3], [선화렆카투슬4] 26.06.15~06.22 ...`

단일 주인공 수치가 없는 복합 분석(후킹+전환 동시 등)은 "지출" 대신 **"기준"** 사용:
예: `[선화렉카투슬] 26.06.15~06.22 총 404,731원 기준`

**헤더 줄 + `라이브` 적용 범위:**
- **인트로 간 비교** (같은 세트 내 인트로끼리): 주인공 헤더 줄(`지출`) 표기 + 비교 세트의 라이브 기간을 `라이브` 키워드로 아래 줄에 표기
  ```
  [선화렉카슬] 2026.06.08~06.26 총 985,708원 지출
  [왜안해디슬] 2026.06.08~06.20 라이브
  ```
- **세트 간 비교** (비교 블록 있음): 헤더 줄(`지출`) 및 `라이브` 줄 **모두 생략**. 기간·지출은 비교 블록 안 각 소재 줄에 표기하므로 중복 금지.
  ```
  TEST 2. [선화렉카슬]보다 전환 개선되는지? = 아쉽다!

  **[선화렉카투슬] vs [선화렉카슬] (비프모 & 40만원 내외 지출 기준)
  [선화렉카투슬] (26.06.15~06.22) : 404,731원 지출 / CPA 138,453원 (전환 1건) / ROAS 81%
  [선화렉카슬] (26.04.27~05.05) : 456,682원 지출 / CPA 91,336원 (전환 5건) / ROAS 191%
  → [선화렉카슬]보다 CPA 47,117원 아쉽, ROAS 110% 아쉽
  ```

## 인트로별 분석 포맷
결론에서 인트로 번호는 v1~v10 사용 (DB 원본 번호 v33 등 사용 X). 인트로는 최대 v10까지 존재할 수 있음.
전환 0건 인트로: `(전환 0건)` — `→ X` 붙이지 않음. (단, "전환이 나올까?" 가설은 전환 0건 → X 표기.)
**인트로 데이터 줄 정렬**: 반드시 인트로 번호 오름차순(v1→v2→v3→...)으로 표기. 효율 좋은 순서(전환 많은 순, CPA 낮은 순 등)로 정렬 금지. (단, **지출 순위·CPA 순위 블록**은 예외 — 해당 블록은 순위대로 표기.)

개별 인트로 줄:
```
v1(유형) N원 / CPA N원 (전환 N건) → O/△/X
v2(유형) N원 / (전환 0건)
```
단, 인트로 위너 가설("위너 인트로는?", "어떤 인트로가 효율이 좋을지" 등)에서는 개별 인트로 줄에 → O/X/△ 붙이지 않음. 판정은 TEST 줄(=)에만 표기.

순위 비교:
```
**지출 순위
1위_v1/유형 (N원) → 지출 위너 v1
2위_v2/유형 (N원)
3위_v3/유형 (N원)

**CPA 순위
1위_v3/유형 (N원) *[서브 구간명] 제외 기준 → CPA 위너 v3  ← 서브 구간 기준일 때만
1위_v3/유형 (N원) → CPA 위너 v3                              ← 전체 기간 기준일 때
2위_v1/유형 (N원)
3위_v2/유형 (전환 없음)
```

규칙:
- 지출·CPA 순위 모두 수치는 `(N원)` 형식으로 통일
- 전환 없으면 `(전환 없음)` 형식
- 특정 서브 구간(예: [재세팅] 제외 구간) 수치를 CPA로 쓸 경우 수치 뒤에 `*[구간명] 제외 기준` 인라인 주석 추가 — 이 `*`는 인트로 순위 내 인라인 표기용이며, 세트 간 비교 블록의 각주 `*` 금지 규칙과 별개
- CPA 순위가 지출 순위와 역전되면 판정 줄 아래에 별도 노티: `** 단, CPA는 [소재명] 가장 우수 ([조건])`

## 후킹 분석 포맷
"어떤 인트로가 후킹이 가장 좋을까?", "후킹지표가 개선되는지?" 등 CTR/CPC 비교가 목적인 가설. 지출 순위·CPA 순위 블록 사용 금지.
- 테스트 전체가 후킹 분석인 경우 `후킹지표:` 라벨 생략 — CTR/CPC 바로 표기
- 동일 TEST 내에 전환 지표가 혼재할 때만 `후킹지표:` 라벨 사용 (`[서브] 후킹:` 사용 금지)
- 플랫폼별(후킹) 분석 시 **instagram 데이터만** 사용. 오디언스 네트워크 언급 절대 금지.
- 출처: 비교 블록 헤더 끝에 `광고관리자_instagram` 표기 — 예: `(비프모 & 40만원 내외 지출 기준, 광고관리자_instagram)`. 헤더 외 별도 출처 줄 추가 금지 (이중 표기 금지).
- 인트로 간 후킹 분석 시: 헤더 다음 줄에 `(소재명/기간, 광고관리자_instagram)` 형태 별도 줄 추가 금지. 데이터 줄에 기간 직접 표기.
- **기간 매칭:**
  - 같은 세트 내 인트로 간 비교: 지출 적은 인트로의 총 지출 기준으로 최대한 비슷한 지출 구간 선별
  - 다른 세트끼리 비교: 최대한 비슷한 지출 구간 매칭하여 기간 선별
예:
```
**[선화렉카투슬] vs [선화렉카슬] (비프모 & 40만원 내외 지출 기준, 광고관리자_instagram)
[선화렉카투슬] (26.06.15~06.22) : 404,731원 지출 / CTR 0.99% / CPC 5,008원
[선화렉카슬] (26.04.27~05.04) : 456,682원 지출 / CTR 1.28% / CPC 2,047원
→ [선화렉카슬]보다 CTR 0.29% 아쉽, CPC 2,961원 아쉽
```

## 전환지표 비교 분석 포맷
"전환에 더 유리한지?", "전환 지표는 어떤 게 제일 좋은지?", "전환에 개선되는지?", "전환이 가장 좋을까?" 등 전환 지표(CPA/ROAS) 비교가 목적인 가설.
지출 순위·CPA 순위 블록 사용 금지 (인트로 간 비교도 동일).

기간 선별:
- 같은 세트 내 인트로 간 비교:
  1. 전체 기간 각 인트로의 총 지출 계산
  2. 지출이 가장 적은 인트로를 기준으로 설정 (기준 지출 = N원)
  3. 나머지 인트로는 누적 지출이 기준 지출에 도달하는 시점까지 슬라이싱
  4. 슬라이싱된 구간의 데이터만 사용 — 슬라이싱으로 인트로별 날짜 구간이 달라지므로 각 인트로 줄에 날짜 개별 표기 (헤더에 공통 날짜 한 줄 쓰는 방식 사용 금지)
  5. 프모 조건도 동일하게 맞출 것 (동일 프모/비프모 구간 내에서 슬라이싱)
- 다른 세트끼리 비교: 동일 프모 여부 조건 & 최대한 비슷한 지출 구간 매칭

판정:
- CPA·ROAS 둘 다 우수 → 전환 우수 판정 (O)
- CPA 또는 ROAS 하나만 우수 → △
- 전부 전환 0건 → `= ? (판정불가)`

포맷: 세트 간 비교 블록 포맷과 동일 (CPA·ROAS 표기). 인트로 간 비교 시에도 동일 포맷 적용. 인트로가 3개 이상이면 `v1 vs v2 vs v3` 형태로 헤더 구성.
```
**v1(유형) vs v2(유형) vs v3(유형) (비프모 & 9만원 내외 지출 기준)
v1(유형) (26.05.25~05.28) : 89,XXX원 지출 / CPA N원 (전환 N건) / ROAS N%
v2(유형) (26.05.25~06.04) : 89,368원 지출 / (전환 0건)
v3(유형) (26.05.25~05.26) : 89,XXX원 지출 / CPA N원 (전환 N건) / ROAS N%
→ v3가 CPA N원 우수
```

## 효율 복사 비교 블록 포맷
주인공 수치는 헤더에 이미 표기하므로 비교 블록에서 생략. 비교 세트 수치만 표기:
```
* vs [비교세트] (조건) 비교세트 지출 또는 매출
→ 지출 N% 복사 / 매출 N%
```

## 연령별 분석 포맷
```
TEST N. [가설] = [판정]
(소재명/기간, 광고관리자)

V1(=유형): [연령대]에게 지출 및 전환 가장 우수 (CPA N원)
V2(=유형): [연령대]에 지출 집중됐으나 전환 저조
```

## 판정 기호
- O: 가설 지지 / X: 가설 기각 / △: 판단 보류 / ?: 판정 불가
- 판정 뒤 괄호로 핵심 한 마디: = X (CPA 악화), = △ (낮은 신뢰도), = △ (vN_cpa위너)
- 인트로 위너 확정 시 직접 서술: `= v3(참여형)` / 연령 가설은 결과 직접 서술

## 인트로 번호 매핑 규칙
같은 세트 내 인트로(v1, v2, v3...)는 소재 번호 오름차순으로 매핑한다.
예: v33, v34, v35 세트 → v1=v33, v2=v34, v3=v35
이 규칙으로 자동 매핑하며, 사용자에게 확인 요청 절대 금지.

## 캠페인 귀속 확인
소재가 어느 캠페인 소속인지는 제공된 DB 데이터의 캠페인 컬럼에서 직접 확인한다.
사용자에게 캠페인 귀속을 묻지 않는다.

## 절대 금지 사항
- **분석 과정·중간 계산 과정 출력 금지.** 데이터 검토·계산은 내부에서 처리하고 최종 결론 포맷만 바로 출력할 것.
- 분석 전후로 사용자에게 확인 요청하거나 질문하는 것 금지
- "맞는지 확인 부탁드립니다", "확인해 주시면", "데이터를 공유해 주시면" 등의 표현 금지
- 불명확한 사항은 가정하고 진행, 가정한 내용은 결론 내 괄호로 간단히 명시
- 숫자+숫자=숫자 형태의 단계별 산술 출력 금지. 일별 데이터 합산 과정(예: 100+200+300=600) 출력 절대 금지. 계산은 내부에서 처리하고 최종값(합계, CPA, ROAS)과 그 기간·소재명만 출력할 것
- TEST 사이 `---` 구분선 사용 금지
- 각 TEST 마지막 줄: 단순 재요약 금지. 중요한 특이사항(예상 밖 패턴, 지표 반전, 프모 불일치 등)이 있을 때만 한 줄 서술 허용 — 모든 TEST 유형에 동일 적용
- 데이터 나열 후 동일 내용을 재요약하는 줄 생성 금지. "전체 판정:", "정리하면" 등의 재요약 표현 사용 금지
- 기간 매칭 이유를 별도 설명 줄로 출력 금지 (`(소재명은 누적 지출 기준 매칭: 기간/지출)` 형태 불가)
- **계절 메모 출력 조건 엄수**: 봄(3~5월)·가을(9~11월) 기간 소재가 포함된 비교는 계절 메모 출력 자체 금지. "동일 계절", "해당 없음", "감안 불필요" 같은 확인 문구도 절대 출력 금지. 여름(6~8월)·겨울(12~2월) 차이가 날 때만 출력.
- **비교 블록 헤더에 프모 조건이 이미 명시된 경우**, 마지막 줄에 프모 관련 재설명 금지. (예: 헤더에 "프모 조건 상이" 있으면 마지막 줄에 각 소재의 프모 종류·월 재서술 금지)
- **프모 이름·월 언급 금지**: 결론 내에서 프로모션 명칭(봄맞이, 라스트찬스 등)이나 특정 월("4월 프모", "5월 라스트찬스 프모" 등) 절대 출력 금지. "프모" / "비프모"로만 표현.
- 사용자가 소재에 포함시킨 구성([재세팅], [믹스] 등)을 임의로 배제하거나 서브 구간 수치로 대체 금지. 서브 구간 분석은 해당 TEST의 판정 보조용이며, 다른 TEST(예: 위너 순위 비교)에서는 반드시 소재 전체 기간의 공식 수치를 사용할 것

## 데이터 부재 처리
- 비교 소재 데이터가 없으면 해당 TEST 한 줄로만 표기:
  "TEST N. 비교 소재 데이터 없음 — 분석 제외"
- 데이터가 있는 TEST는 정상 분석 진행

## 기타 규칙
- 표(table) 사용 금지
- 예측/기대 평가 표현 금지 ("예측 빗나감", "기대와 달리" 등)
- 프로모션 명칭 대신 "프모 기간" / "비프모 기간"으로 표현
- 플랫폼별(후킹/CTR/CPC) 분석은 instagram 데이터만 사용. 오디언스 네트워크 관련 언급 일체 금지.
- 스냅애즈와 광고관리자 데이터 혼용 시만 출처 명시
- 양측 → 둘다 / 등가지출 → 동일지출 / 열위 → 아쉬움
- 데이터 나열 후 재요약 줄 금지
- 결론 본문에서 비교 소재명 참조 시 `[소재명]` 형식 사용
- 우열 비교 판정: O/X/△ 기호 절대 사용 금지. `= 아쉽다!` / `= 두 세트보다 아쉽다!` 등 구어체 직접 서술만 사용
  TEST 제목: 비교 소재명만 나열, **주인공 소재명은 제외**. 비교 소재 2개 이상이면 `/`로 구분.
  예: `TEST 3. 듀플 [왜안해디슬] / 휴터 [남사친선넘슬]과 효율 비교하면? = 아쉽다!`

## 정답 예시 — 이 형식을 정확히 따를 것

아래는 실제 첨삭된 올바른 결론 예시. 날짜·기호·소재명·판정·줄 구성 전부 이 형식 기준으로 출력할 것.

```
TEST 1. 전환이 나오는지 = O (v3 메인)
왜안휴에이슬, 왜안휴에이슬[재세팅] 26.06.08~06.26 총 1,203,000원 지출

v1(대세감) 291,734원 / CPA 97,245원 (전환 3건) → △
v2(한정성) 112,805원 / (전환 0건)
v3(참여형) 798,406원 / CPA 114,058원 (전환 7건) → △
  └ [재세팅] 제외한 v3 (26.06.08~06.19)만: 581,168원 / CPA 83,024원 (전환 7건) → O

TEST 2. 위너 인트로는? = v3(참여형)

**지출 순위
1위_v3/참여형 (798,406원) → 지출 위너 v3
2위_v1/대세감 (291,734원)
3위_v2/한정성 (112,805원)

**CPA 순위
1위_v3/참여형 (83,024원) *[재세팅] 제외 기준 → CPA 위너 v3
2위_v1/대세감 (97,245원)
3위_v2/한정성 (전환 없음)

TEST 3. 듀플 [왜안해디슬] / 휴터 [남사친선넘슬]과 효율 비교하면? = 아쉽다!

**휴터 [왜안휴에이슬] vs 듀플 [왜안해디슬] (비프모 & 120만원 내외 지출 시점 기준)
[왜안휴에이슬](휴터) (26.06.08~06.23) : 871,650원 지출 / CPA 87,165원 (전환 N건) / ROAS 129%
[왜안해디슬](듀플) (26.06.08~06.15) : 1,275,186원 지출 / CPA 67,115원 (전환 N건) / ROAS 192%
→ [왜안해디슬]이 CPA 2만원 아쉽, ROAS 63% 아쉽

**[왜안휴에이슬] vs [남사친선넘슬] (비프모 & 90만원 내외 지출 시점 기준)
[왜안휴에이슬] (26.06.08~06.23) : 871,650원 지출 / CPA 87,165원 (전환 N건) / ROAS 129%
[남사친선넘슬] (26.06.01~06.07) : 1,214,088원 지출 / CPA 75,881원 (전환 N건) / ROAS 160%
→ [남사친선넘슬]보다 CPA 11,284원 아쉽, ROAS 31% 아쉽
```

아래는 추가 첨삭 예시 — 그룹 비교 및 비연속 기간 패턴.

```
TEST 1. 몸매 vs 대세감? = ? (판정불가)
선화렉카투슬 2026.06.15~06.22 총 404,731원 지출

v1(몸매1) 62,557원 / (전환 0건)
v2(몸매2) 123,831원 / (전환 0건)
v3(대세감1) 79,891원 / (전환 0건)
v4(대세감2) 138,453원 / CPA 138,453원 (전환 1건) → △

v4만 전환 1건으로 방향성은 대세감 쪽이나, 전체 1건에 그쳐 신뢰도가 낮아 판정 불가.

TEST 2. [선화렉카슬]과 전환 비교하면? = 아쉽다!

**[선화렉카투슬] vs [선화렉카슬] (비프모 & 40만원 내외 지출 기준)
[선화렉카투슬] (26.06.15~06.22) : 404,731원 지출 / CPA 404,731원 (전환 1건) / ROAS 28%
[선화렉카슬] (26.04.29~05.05 + 05.11) : 400,891원 지출 / CPA 57,270원 (전환 7건) / ROAS 273%
→ [선화렉카슬]보다 CPA 347,461원 아쉽, ROAS 245% 아쉽
```
"""


def init_db():
    os.makedirs("data", exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS raw_data (
            플랫폼 TEXT,
            캠페인 TEXT,
            광고세트 TEXT,
            소재 TEXT,
            날짜 TEXT,
            광고세트일예산 REAL,
            광고비용 REAL,
            GA_ROAS REAL,
            GA_전환수 REAL,
            GA_전환매출액 REAL,
            GA_객단가 REAL,
            GA_전환율 REAL,
            CPC REAL,
            CPM REAL,
            CTR REAL,
            H_전환매출액 REAL,
            H_구매건수 REAL,
            H_구매전환율 REAL,
            H_ROAS REAL,
            H_객단가 REAL,
            H_구매당비용 REAL,
            PRIMARY KEY (플랫폼, 캠페인, 광고세트, 소재, 날짜)
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS promotions (
            name TEXT PRIMARY KEY,
            start_date TEXT,
            end_date TEXT
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS promo_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TEXT,
            user TEXT,
            action TEXT,
            detail TEXT
        )
    """)
    conn.executemany("INSERT OR IGNORE INTO promotions VALUES (?,?,?)", [
        ("2025 1월 설 프로모션", "2025-01-18", "2025-01-31"),
        ("2025 4월 봄맞이 프로모션", "2025-04-05", "2025-04-13"),
        ("2025 5월 가정의달 프로모션", "2025-05-19", "2025-05-25"),
        ("2025 6월 쿨세일 프로모션", "2025-06-20", "2025-06-30"),
        ("2025 7월 메가 할인전", "2025-07-14", "2025-08-04"),
        ("2025 8월 상반기 결산", "2025-08-24", "2025-09-15"),
        ("2025 10월 추석 프로모션", "2025-10-01", "2025-10-19"),
        ("2025 10월 얼리블프 프로모션", "2025-10-27", "2025-11-10"),
        ("2025 11월 블랙위크 프로모션", "2025-11-24", "2025-12-14"),
        ("2025 12월 프로뉴이어 프로모션", "2025-12-23", "2026-01-11"),
        ("2026 1월 새해 프로모션", "2026-01-29", "2026-02-09"),
        ("2026 2월 설날 프로모션", "2026-02-10", "2026-02-23"),
        ("2026 3월 기록 돌파 프로모션", "2026-03-13", "2026-03-22"),
        ("2026 4월 봄맞이 프로모션", "2026-04-08", "2026-04-26"),
        ("2026 5월 라스트 찬스", "2026-05-18", "2026-05-31"),
        ("2026 6월 바캉스 구조대", "2026-06-24", "2026-07-05"),
    ])
    # 일회성 마이그레이션: 구 이름 → 연도&월 포함 이름
    _renames = [
        ("설 프로모션",              "2025 1월 설 프로모션"),
        ("봄맞이 프로모션 (2025)",   "2025 4월 봄맞이 프로모션"),
        ("가정의달 프로모션",        "2025 5월 가정의달 프로모션"),
        ("쿨세일 프로모션",          "2025 6월 쿨세일 프로모션"),
        ("추석 프로모션",            "2025 10월 추석 프로모션"),
        ("얼리블프 프로모션",        "2025 10월 얼리블프 프로모션"),
        ("블랙위크 프로모션",        "2025 11월 블랙위크 프로모션"),
        ("프로뉴이어 프로모션",      "2025 12월 프로뉴이어 프로모션"),
        ("새해 프로모션",            "2026 1월 새해 프로모션"),
        ("설날 프로모션",            "2026 2월 설날 프로모션"),
        ("기록 돌파 프로모션",       "2026 3월 기록 돌파 프로모션"),
        ("봄맞이 프로모션",          "2026 4월 봄맞이 프로모션"),
        ("라스트 찬스",              "2026 5월 라스트 찬스"),
        ("바캉스 구조대",            "2026 6월 바캉스 구조대"),
    ]
    for old, new in _renames:
        # 구 이름이 남아 있을 때만: seed로 삽입된 새 이름 행을 먼저 제거 후 rename
        conn.execute(
            "DELETE FROM promotions WHERE name=? AND EXISTS (SELECT 1 FROM promotions WHERE name=?)",
            (new, old)
        )
        conn.execute("UPDATE promotions SET name=? WHERE name=?", (new, old))
    conn.commit()
    conn.close()


def delete_campaign(campaign_name):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM raw_data WHERE 캠페인 = ?", (campaign_name,))
    conn.commit()
    conn.close()


def get_data_status():
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("""
            SELECT 캠페인,
                   MIN(날짜) as 시작일,
                   MAX(날짜) as 종료일
            FROM raw_data
            WHERE 날짜 != '' AND 날짜 IS NOT NULL
            GROUP BY 캠페인
            ORDER BY 캠페인
        """, conn)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df


def process_xlsx(file):
    wb = openpyxl.load_workbook(file, read_only=True)
    ws = wb.active
    rows = []
    last_platform = last_campaign = last_adset = last_creative = None

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i < 4:
            continue
        if all(v is None for v in row):
            continue
        if row[4] is None and row[0] is not None and str(row[0]).startswith('이('):
            continue
        if row[4] is None and row[0] is None:
            continue

        platform = row[0] if row[0] else last_platform
        if platform and str(platform).startswith('총('):
            continue
        campaign = row[1] if row[1] else last_campaign
        adset = row[2] if row[2] else last_adset
        creative = row[3] if row[3] else last_creative

        last_platform = platform
        last_campaign = campaign
        last_adset = adset
        last_creative = creative

        if not row[4]:
            continue

        spend = row[6]
        if spend is None or spend == 0:
            continue

        rows.append((
            platform, campaign, adset, creative, str(row[4]),
            row[5], row[6], row[7], row[8], row[9],
            row[10], row[11], row[12], row[13], row[14],
            row[15], row[16], row[17], row[18], row[19], row[20]
        ))

    wb.close()
    return rows


def insert_rows(rows):
    conn = sqlite3.connect(DB_PATH)
    inserted = 0
    skipped = 0
    for row in rows:
        try:
            cursor = conn.execute("""
                INSERT OR IGNORE INTO raw_data VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, row)
            if cursor.rowcount > 0:
                inserted += 1
            else:
                skipped += 1
        except Exception:
            skipped += 1
    conn.commit()
    conn.close()
    return inserted, skipped


def resolve_campaign(prefix):
    for alias, campaign in CAMPAIGN_ALIAS.items():
        if prefix.startswith(alias):
            return campaign
    return None


def get_creative_data(creative_name, prefix=None):
    conn = sqlite3.connect(DB_PATH)
    try:
        campaign = resolve_campaign(prefix) if prefix else None
        if campaign:
            # 해당 캠페인에서 우선 검색
            df = pd.read_sql("""
                SELECT 캠페인, 소재, 날짜, 광고비용, GA_ROAS, GA_전환수, GA_전환매출액, GA_객단가, GA_전환율, CPC, CPM, CTR
                FROM raw_data
                WHERE 소재 LIKE ? AND 캠페인 LIKE ? AND 날짜 != '' AND 날짜 IS NOT NULL
                ORDER BY 소재, 날짜
            """, conn, params=(f"%{creative_name}%", f"%{campaign}%"))
            # 없으면 전체 검색
            if df.empty:
                df = pd.read_sql("""
                    SELECT 캠페인, 소재, 날짜, 광고비용, GA_ROAS, GA_전환수, GA_전환매출액, GA_객단가, GA_전환율, CPC, CPM, CTR
                    FROM raw_data
                    WHERE 소재 LIKE ? AND 날짜 != '' AND 날짜 IS NOT NULL
                    ORDER BY 소재, 날짜
                """, conn, params=(f"%{creative_name}%",))
        else:
            df = pd.read_sql("""
                SELECT 캠페인, 소재, 날짜, 광고비용, GA_ROAS, GA_전환수, GA_전환매출액, GA_객단가, GA_전환율, CPC, CPM, CTR
                FROM raw_data
                WHERE 소재 LIKE ? AND 날짜 != '' AND 날짜 IS NOT NULL
                ORDER BY 소재, 날짜
            """, conn, params=(f"%{creative_name}%",))
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df


def group_creative_names(creative_names, keyword):
    import re
    groups = {}
    for raw in creative_names:
        idx = raw.find(keyword)
        if idx == -1:
            groups.setdefault("(기타)", []).append(raw)
            continue
        suffix = raw[idx + len(keyword):]
        suffix = re.sub(r'(?:_[A-Za-z]{2,}\d+)+$', '', suffix)  # remove _Mhs7_Ahs7
        suffix = re.sub(r'\d+$', '', suffix).strip()             # remove trailing number
        # suffix starting with '/' = slash-separated variant → preserve slash
        group_key = keyword + suffix
        groups.setdefault(group_key, []).append(raw)
    return groups


def get_creative_data_filtered(name_list):
    if not name_list:
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    try:
        placeholders = ','.join('?' * len(name_list))
        df = pd.read_sql(f"""
            SELECT 캠페인, 소재, 날짜, 광고비용, GA_ROAS, GA_전환수, GA_전환매출액, GA_객단가, GA_전환율, CPC, CPM, CTR
            FROM raw_data
            WHERE 소재 IN ({placeholders}) AND 날짜 != '' AND 날짜 IS NOT NULL
            ORDER BY 소재, 날짜
        """, conn, params=name_list)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df


def query_creative_by_period(creative_name, start_date, end_date, exclude_keywords=None):
    conn = sqlite3.connect(DB_PATH)
    try:
        exclude_clause = ""
        params = [f"%{creative_name}%", start_date, end_date]
        if exclude_keywords:
            for kw in exclude_keywords:
                exclude_clause += " AND 소재 NOT LIKE ?"
                params.append(f"%{kw}%")

        df = pd.read_sql(f"""
            SELECT 소재, 날짜, 플랫폼,
                   광고비용,
                   GA_전환수, GA_전환매출액, GA_ROAS, GA_객단가, GA_전환율,
                   CPC, CPM, CTR,
                   H_구매건수, H_전환매출액, H_ROAS, H_객단가, H_구매당비용
            FROM raw_data
            WHERE 소재 LIKE ?
              AND 날짜 >= ? AND 날짜 <= ?
              AND 날짜 != '' AND 날짜 IS NOT NULL
              {exclude_clause}
            ORDER BY 소재, 날짜
        """, conn, params=params)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df


@st.cache_data
def get_campaign_excel(campaign_name):
    import io
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql(
        "SELECT * FROM raw_data WHERE 캠페인 = ? ORDER BY 소재, 날짜",
        conn, params=(campaign_name,)
    )
    conn.close()
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine='openpyxl')
    buf.seek(0)
    return buf.getvalue()


def get_promotions():
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql("SELECT name, start_date, end_date FROM promotions ORDER BY start_date DESC", conn)
    except Exception:
        df = pd.DataFrame(columns=["name", "start_date", "end_date"])
    conn.close()
    return df


def add_promotion(name, start_date, end_date):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("INSERT OR REPLACE INTO promotions VALUES (?,?,?)", (name, start_date, end_date))
    conn.commit()
    conn.close()


def delete_promotion(name):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM promotions WHERE name = ?", (name,))
    conn.commit()
    conn.close()


def log_promo_action(user, action, detail):
    from datetime import datetime as _dt
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO promo_logs (timestamp, user, action, detail) VALUES (?,?,?,?)",
        (_dt.now().strftime("%Y-%m-%d %H:%M:%S"), user or "미입력", action, detail)
    )
    conn.commit()
    conn.close()


def get_promo_logs():
    conn = sqlite3.connect(DB_PATH)
    try:
        df = pd.read_sql(
            "SELECT timestamp, user, action, detail FROM promo_logs ORDER BY id DESC LIMIT 100",
            conn
        )
    except Exception:
        df = pd.DataFrame(columns=["timestamp", "user", "action", "detail"])
    conn.close()
    return df


def get_promo_date_set():
    promos = get_promotions()
    promo_dates = set()
    for _, row in promos.iterrows():
        cur = pd.Timestamp(row['start_date'])
        end_ts = pd.Timestamp(row['end_date'])
        while cur <= end_ts:
            promo_dates.add(cur.strftime('%Y-%m-%d'))
            cur += pd.Timedelta(days=1)
    return promo_dates


def classify_hypothesis_types(hypothesis):
    """가설 텍스트에서 TEST 별 유형을 키워드 매칭으로 분류."""
    import re as _re
    TYPE_NAMES = {
        "유형1": "세트 간 효율 비교",
        "유형2": "인트로 위너",
        "유형3": "콘셉트 그룹 비교",
        "유형4": "효율 복사",
        "유형5": "후킹 분석",
        "유형6": "전환지표 비교",
        "유형7": "연령별 분석",
        "유형8": "전환 여부",
        "분류불가": "분류 불가",
    }
    # 한 줄 / 여러 줄 모두 대응: TEST N. 등장 위치를 기준으로 분리
    parts = _re.split(r'(?=\bTEST\s*\d+\.)', hypothesis, flags=_re.IGNORECASE)
    tests = [p.strip() for p in parts if _re.match(r'TEST\s*\d+\.', p.strip(), _re.IGNORECASE)]
    if not tests:
        return []
    results = []
    for part in tests:
        # 첫 줄만 TEST 질문줄로 사용
        test_line = part.split('\n')[0].strip()
        t = test_line.lower()
        m = _re.match(r'(TEST\s*\d+\.)\s*', test_line, _re.IGNORECASE)
        label = m.group(1) if m else "TEST"
        question = test_line[m.end():].strip() if m else test_line
        if "복사" in t:
            code = "유형4"
        elif "전환이 나올까" in t or "일단 전환" in t:
            code = "유형8"
        elif "연령" in t or "나이대" in t:
            code = "유형7"
        elif "후킹" in t:
            code = "유형5"
        elif any(k in t for k in ["전환이 가장", "전환에 더 유리", "전환 지표", "전환 개선", "전환에 개선"]):
            code = "유형6"
        elif "vs" in t and not _re.search(r'\[[^\]]+\]\s*vs\s*\[', t, _re.IGNORECASE):
            code = "유형3"
        elif "인트로" in t or "위너" in t:
            code = "유형2"
        elif any(k in t for k in ["효율", "우수", "더 좋을까", "성과가 좋"]):
            code = "유형1"
        else:
            code = "분류불가"
        results.append({"label": label, "question": question, "code": code, "name": TYPE_NAMES[code]})
    return results


def correct_creative_names(text, known_names):
    """결론 내 소재명을 실제 DB 소재명과 유사도 비교하여 자동 교정."""
    import difflib
    if not known_names:
        return text, {}
    corrections = {}
    bracket_names = list(dict.fromkeys(_re.findall(r'\[([^\]]+)\]', text)))
    for name in bracket_names:
        if name in known_names:
            continue
        matches = difflib.get_close_matches(name, known_names, n=1, cutoff=0.75)
        if matches:
            corrections[name] = matches[0]
    for wrong, correct in corrections.items():
        text = text.replace(f"[{wrong}]", f"[{correct}]")
    return text, corrections


def compute_ad_manager_requirements(hypothesis, main_label, groups_found, raw_data_cache):
    import re
    h = hypothesis.lower()
    types = []
    age_keywords = ["연령", "1824", "18-24", "2534", "24-34", "3544", "35-44", "중년", "나이"]
    if any(k in h for k in age_keywords) or re.search(r'\d+세', h) or re.search(r'\d+살', h):
        types.append("연령별")
    if any(k in h for k in ["후킹", "ctr", "cpc", "cpm", "플랫폼", "클릭"]):
        types.append("플랫폼별")
    if any(k in h for k in ["성별", "남성", "남자", "여성", "여자"]):
        types.append("성별")
    if not types:
        return []

    main_df = raw_data_cache.get("__main__", pd.DataFrame())
    if main_df.empty:
        return []

    main_daily = main_df.groupby('날짜')['광고비용'].sum().sort_index()
    main_total = main_daily.sum()

    # [소재명]vN 패턴 파싱 — 특정 인트로가 지목된 비교 세트 추출
    intro_refs = {}
    for m in re.finditer(r'\[([^\]]+)\]v(\d+)', hypothesis):
        creative = m.group(1).replace('/', '').strip()
        intro_num = int(m.group(2))
        if creative not in intro_refs:
            intro_refs[creative] = intro_num

    def fmt(s, e):
        sy, sm, sd = s.split("-")
        ey, em, ed = e.split("-")
        s_str = f"{sy[2:]}.{sm}.{sd}"
        e_str = f"{em}.{ed}" if sy == ey else f"{ey[2:]}.{em}.{ed}"
        return f"{s_str}~{e_str}"

    reqs = [{
        "label": main_label or "주인공",
        "period": fmt(main_daily.index.min(), main_daily.index.max()),
        "types": types,
        "note": "",
        "promo_note": "",
    }]

    comp_keys = [k for k in groups_found if k != "__main__" and groups_found.get(k)]
    if comp_keys and main_total > 0:
        promo_dates = get_promo_date_set()
        main_promo_spend = main_daily[main_daily.index.isin(promo_dates)].sum()
        main_is_promo = (main_promo_spend / main_total) > 0.5

        for ck in comp_keys:
            comp_df = raw_data_cache.get(ck, pd.DataFrame())
            if comp_df.empty:
                continue

            # 특정 인트로 지목 시 해당 인트로 데이터만 사용
            display_label = ck
            is_intro_ref = ck in intro_refs
            if is_intro_ref:
                intro_num = intro_refs[ck]
                all_creatives = sorted(comp_df['소재'].unique())
                if intro_num <= len(all_creatives):
                    target_creative = all_creatives[intro_num - 1]
                    comp_df = comp_df[comp_df['소재'] == target_creative]
                    display_label = f"{ck}v{intro_num}"

            comp_daily = comp_df.groupby('날짜')['광고비용'].sum().sort_index()

            matching = {d: v for d, v in comp_daily.items()
                        if (d in promo_dates) == main_is_promo}
            promo_mismatch = len(matching) < 2
            if promo_mismatch:
                matching = dict(comp_daily)

            cumsum = 0
            start_d = end_d = None
            for d in sorted(matching):
                if start_d is None:
                    start_d = d
                cumsum += matching[d]
                end_d = d
                if cumsum >= main_total:
                    break

            note = "(프모 조건 불일치, 동일지출 기준)" if promo_mismatch else "(동일지출 기준)"
            promo_note = ""
            if promo_mismatch:
                main_promo_str = "프모" if main_is_promo else "비프모"
                comp_promo_str = "비프모" if main_is_promo else "프모"
                promo_note = f"단, [{main_label}]은 {main_promo_str} / [{ck}]는 {comp_promo_str}"

            reqs.append({
                "label": display_label,
                "period": fmt(start_d, end_d),
                "types": types,
                "note": note,
                "promo_note": promo_note,
            })

    return reqs


@st.dialog("프로모션 추가")
def add_promotion_dialog():
    import datetime
    name = st.text_input("프로모션명", placeholder="예) 2026 7월 바캉스 구조대")
    col1, col2 = st.columns(2)
    with col1:
        start = st.date_input("시작일", value=datetime.date.today())
    with col2:
        end = st.date_input("종료일", value=datetime.date.today())
    st.divider()
    who = st.text_input("이름", placeholder="홍길동")
    why = st.text_input("추가 이유", placeholder="신규 프로모션 일정 확정")
    can_submit = bool(name and who and why)
    if st.button("추가", type="primary", disabled=not can_submit):
        add_promotion(name, start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))
        log_promo_action(who, "추가", f"이유: {why} | {name} ({start}~{end})")
        st.rerun()


@st.dialog("프로모션 수정")
def edit_promotion_dialog(orig_name, orig_start, orig_end):
    import datetime
    name = st.text_input("프로모션명", value=orig_name)
    col1, col2 = st.columns(2)
    with col1:
        start = st.date_input("시작일", value=datetime.date.fromisoformat(orig_start))
    with col2:
        end = st.date_input("종료일", value=datetime.date.fromisoformat(orig_end))
    st.divider()
    who = st.text_input("이름", placeholder="홍길동")
    why = st.text_input("수정 이유", placeholder="종료일 변경")
    can_submit = bool(who and why)
    if st.button("수정 완료", type="primary", disabled=not can_submit):
        if name != orig_name:
            delete_promotion(orig_name)
        add_promotion(name, start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))
        log_promo_action(who, "수정", f"이유: {why} | {orig_name} → {name} ({start}~{end})")
        st.rerun()


@st.dialog("프로모션 삭제")
def delete_promo_confirm_dialog(promo_name):
    st.warning(f"**{promo_name}** 을(를) 삭제합니다.")
    who = st.text_input("이름", placeholder="홍길동")
    why = st.text_input("삭제 이유", placeholder="잘못 입력한 프로모션")
    can_submit = bool(who and why)
    col1, col2 = st.columns(2)
    with col1:
        if st.button("삭제 확인", type="primary", use_container_width=True, disabled=not can_submit):
            log_promo_action(who, "삭제", f"이유: {why} | {promo_name}")
            delete_promotion(promo_name)
            st.rerun()
    with col2:
        if st.button("취소", use_container_width=True):
            st.rerun()


def build_system_prompt():
    promos = get_promotions()
    lines = "\n".join(
        f"  - {r['start_date'].replace('-', '.')}-{r['end_date'].replace('-', '.')}"
        for _, r in promos.iterrows()
    )
    return SYSTEM_PROMPT.replace("<<PROMO_LIST>>", lines)


@st.dialog("삭제 확인")
def confirm_delete_dialog(campaign_name):
    st.warning(f"**{campaign_name}** 데이터를 전부 삭제합니다.\n\n이 작업은 되돌릴 수 없습니다.")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("삭제", type="primary", use_container_width=True):
            delete_campaign(campaign_name)
            st.rerun()
    with col2:
        if st.button("취소", use_container_width=True):
            st.rerun()


# 앱 시작
st.set_page_config(page_title="콘텐츠 인사이트 분석", layout="wide")
init_db()

st.title("콘텐츠 인사이트 분석")

tab4, tab1, tab5 = st.tabs(["분석", "데이터 현황", "분석 규칙"])

# ── 탭 1: 데이터 현황 (서브탭: 현황 / 업데이트 / 조회) ──────────
with tab1:
    _sub_status, _sub_update, _sub_query = st.tabs(["📊 데이터 현황", "⬆️ 업데이트", "🔍 조회"])

    # ── 서브탭: 현황 ──────────────────────────────────────────────
    with _sub_status:
        st.subheader("데이터 업데이트 현황")
        status = get_data_status()
        if status.empty:
            st.info("등록된 데이터가 없습니다. '업데이트' 탭에서 파일을 업로드해주세요.")
        else:
            for _, row in status.iterrows():
                col1, col2, col3 = st.columns([5, 1, 1])
                with col1:
                    start = row['시작일'].replace('-', '.') if row['시작일'] else '-'
                    end = row['종료일'].replace('-', '.') if row['종료일'] else '-'
                    st.markdown(f"**{row['캠페인']}** : {start} - {end}")
                with col2:
                    st.download_button(
                        "다운로드",
                        data=get_campaign_excel(row['캠페인']),
                        file_name=f"{row['캠페인']}_{start}-{end}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{row['캠페인']}"
                    )
                with col3:
                    if st.button("삭제", key=f"del_{row['캠페인']}"):
                        confirm_delete_dialog(row['캠페인'])

        st.divider()

        col_h, col_btn = st.columns([5, 2])
        with col_h:
            st.subheader("프로모션 현황")
        with col_btn:
            st.write("")
            if st.button("+ 프로모션 추가", use_container_width=True):
                add_promotion_dialog()

        promos = get_promotions()
        if promos.empty:
            st.info("등록된 프로모션이 없습니다.")
        else:
            for _, row in promos.iterrows():
                s = row['start_date'].replace('-', '.')
                e = row['end_date'].replace('-', '.')
                col1, col2, col3 = st.columns([6, 1, 1])
                with col1:
                    st.markdown(f"**{row['name']}** : {s} ~ {e}")
                with col2:
                    if st.button("수정", key=f"edit_promo_{row['name']}"):
                        edit_promotion_dialog(row['name'], row['start_date'], row['end_date'])
                with col3:
                    if st.button("삭제", key=f"del_promo_{row['name']}"):
                        delete_promo_confirm_dialog(row['name'])

        with st.expander("로그 보기"):
            logs = get_promo_logs()
            if logs.empty:
                st.caption("기록된 로그가 없습니다.")
            else:
                logs = logs.rename(columns={"user": "이름", "action": "액션", "detail": "상세", "timestamp": "시각"})
                st.dataframe(logs, use_container_width=True, hide_index=True)

    # ── 서브탭: 업데이트 ──────────────────────────────────────────
    with _sub_update:
        st.subheader("데이터 업데이트")
        st.caption("스냅애즈에서 내보낸 xlsx 파일을 업로드하세요. 0원 행은 자동으로 제거됩니다.")
        uploaded = st.file_uploader(
            "xlsx 파일 선택 (여러 개 가능)",
            type=["xlsx"],
            accept_multiple_files=True
        )
        if uploaded:
            if st.button("업데이트 시작"):
                total_inserted = 0
                for file in uploaded:
                    with st.spinner(f"{file.name} 처리 중..."):
                        rows = process_xlsx(file)
                        inserted, skipped = insert_rows(rows)
                        total_inserted += inserted
                        st.write(f"✓ {file.name}: {inserted:,}건 추가, {skipped:,}건 중복 스킵")
                st.success(f"완료. 총 {total_inserted:,}건 추가됨.")
                st.rerun()

    # ── 서브탭: 조회 ──────────────────────────────────────────────
    with _sub_query:
        st.subheader("데이터 조회")
        st.caption("소재명 일부를 입력하고 기간을 지정하면 일별 데이터를 확인할 수 있습니다.")

        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            search_name = st.text_input("소재명 검색", placeholder="예) 교관듀디슬")
        with col2:
            start_date = st.date_input("시작일")
        with col3:
            end_date = st.date_input("종료일")

        exclude_input = st.text_input(
            "제외 키워드 (쉼표로 구분)",
            placeholder="예) 재세팅, 믹스"
        )

        if st.button("조회") and search_name:
            exclude_keywords = [k.strip() for k in exclude_input.split(",") if k.strip()] if exclude_input else None
            result = query_creative_by_period(
                search_name,
                start_date.strftime("%Y-%m-%d"),
                end_date.strftime("%Y-%m-%d"),
                exclude_keywords=exclude_keywords
            )
            if result.empty:
                st.warning("해당 소재 및 기간의 데이터가 없습니다.")
            else:
                st.success(f"총 {len(result):,}행 조회됨")

                summary = result.groupby("소재").agg(
                    기간시작=("날짜", "min"),
                    기간종료=("날짜", "max"),
                    총광고비=("광고비용", "sum"),
                    GA전환수=("GA_전환수", "sum"),
                    GA전환매출=("GA_전환매출액", "sum"),
                ).reset_index()
                summary["CPA"] = (summary["총광고비"] / summary["GA전환수"]).where(summary["GA전환수"] > 0)
                summary["총광고비"] = summary["총광고비"].map("{:,.0f}원".format)
                summary["GA전환수"] = summary["GA전환수"].map("{:.0f}건".format)
                summary["GA전환매출"] = summary["GA전환매출"].map("{:,.0f}원".format)
                summary["CPA"] = summary["CPA"].map(lambda x: f"{x:,.0f}원" if pd.notna(x) else "-")

                st.markdown("**소재별 요약**")
                st.dataframe(summary, use_container_width=True)

                with st.expander("일별 상세 데이터 보기"):
                    st.dataframe(result, use_container_width=True)

# ── 탭 4: 분석 ────────────────────────────────────────────────
with tab4:
    st.subheader("분석")

    import re as _re

    hypothesis = st.text_area(
        "가설 입력",
        placeholder="예) 교관듀디슬(믹스) TEST 1. 인트로N이 전환 성과가 가장 좋지 않을까?",
        height=120
    )

    creative_name_input = st.text_input(
        "주인공 소재명 입력 (DB에서 라이브 기간 자동 조회)",
        placeholder="예) A왜안휴에이슬"
    )

    if creative_name_input:
        _preview = get_creative_data(creative_name_input)
        if not _preview.empty:
            _s = _preview['날짜'].min().replace('-', '.')
            _e = _preview['날짜'].max().replace('-', '.')
            st.info(f"**{creative_name_input}** 라이브 기간: {_s} - {_e}")
        else:
            st.warning("DB에 데이터 없음. 세팅일 직접 확인 필요.")

    uploaded_csv = None

    # ── 소재 확인 ──────────────────────────────────────────────
    if st.button("소재 확인", disabled=not (hypothesis or creative_name_input)):
        groups_found = {}
        raw_data_cache = {}

        def _collect(cname, key):
            _df = get_creative_data(cname)
            if not _df.empty:
                groups_found[key] = group_creative_names(_df["소재"].unique().tolist(), cname)
                raw_data_cache[key] = _df
            else:
                groups_found[key] = {}

        if creative_name_input:
            _collect(creative_name_input, "__main__")
        for _b in _re.findall(r'\[([^\]]+)\]', hypothesis or ""):
            _cname = _b.replace("/", "").strip()
            if _cname and _cname != creative_name_input.strip() and _cname not in groups_found:
                _collect(_cname, _cname)

        st.session_state["_groups"] = groups_found
        st.session_state["_raw_cache"] = raw_data_cache
        st.session_state.pop("_ad_reqs", None)
        st.session_state.pop("_selection_done", None)
        st.session_state.pop("_not_found_creatives", None)
        st.session_state.pop("_filtered_cache", None)
        st.session_state["_hyp_types"] = classify_hypothesis_types(hypothesis or "")
        st.session_state["_groups_hyp"] = hypothesis
        st.session_state["_groups_main"] = creative_name_input

    # 입력 변경 시 캐시 무효화
    if (st.session_state.get("_groups_hyp") != hypothesis or
            st.session_state.get("_groups_main") != creative_name_input):
        st.session_state.pop("_groups", None)
        st.session_state.pop("_raw_cache", None)
        st.session_state.pop("_ad_reqs", None)
        st.session_state.pop("_selection_done", None)
        st.session_state.pop("_not_found_creatives", None)
        st.session_state.pop("_filtered_cache", None)
        st.session_state.pop("_hyp_types", None)

    # ── 소재 선택 UI ────────────────────────────────────────────
    if "_groups" in st.session_state:
        _hyp_types = st.session_state.get("_hyp_types", [])
        if _hyp_types:
            lines = "\n\n".join(
                f"`{r['code']}` **{r['name']}** — {r['label']} {r['question']}"
                for r in _hyp_types
            )
            st.info(lines)
        for _key, _groups in st.session_state["_groups"].items():
            _label = creative_name_input if _key == "__main__" else _key
            if not _groups:
                st.warning(f"**{_label}**: DB에 데이터 없음")
            elif len(_groups) == 1:
                _n = sum(len(v) for v in _groups.values())
                st.success(f"**{_label}**: {list(_groups.keys())[0]} ({_n}개 소재 자동 포함)")
            else:
                st.multiselect(
                    f"**{_label}** — 포함할 세트 선택 ({len(_groups)}개 발견)",
                    options=list(_groups.keys()),
                    default=list(_groups.keys()),
                    key=f"_ms_{_key}",
                )

        if st.button("선택", key="_btn_select"):
            _gmap_all = st.session_state["_groups"]
            _raw = st.session_state.get("_raw_cache", {})
            _filtered = {}
            for _k, _gmap in _gmap_all.items():
                if not _gmap:
                    continue
                _sel = (st.session_state.get(f"_ms_{_k}", list(_gmap.keys()))
                        if len(_gmap) > 1 else list(_gmap.keys()))
                _names = [n for gk in _sel for n in _gmap.get(gk, [])]
                if _k in _raw and _names:
                    _filtered[_k] = _raw[_k][_raw[_k]["소재"].isin(_names)]
            st.session_state["_ad_reqs"] = compute_ad_manager_requirements(
                hypothesis or "", creative_name_input, _gmap_all, _filtered
            )
            st.session_state["_filtered_cache"] = _filtered
            st.session_state["_selection_done"] = True
            _not_found = [_k for _k, _gmap in _gmap_all.items() if _k != "__main__" and not _gmap]
            st.session_state["_not_found_creatives"] = _not_found

        for _name in st.session_state.get("_not_found_creatives", []):
            st.warning(f"⚠️ 가설 내 [{_name}]이 조회되지 않습니다. 소재명을 확인해주세요.")

    # ── 광고관리자 데이터 안내 + 파일 업로드 ────────────────────
    if st.session_state.get("_selection_done"):
        _ad_reqs = st.session_state.get("_ad_reqs", [])
        if _ad_reqs:
            st.markdown("**광고관리자 데이터 필요 기간**")
            for _req in _ad_reqs:
                _meta = " + ".join(_req["types"]) if _req["types"] else ""
                _note = _req["note"]
                _suffix = "  " + " · ".join(filter(None, [_meta, _note])) if (_meta or _note) else ""
                st.info(f"**{_req['label']}** | {_req['period']}{_suffix}")
                if _req.get("promo_note"):
                    st.caption(_req["promo_note"])
            uploaded_csv = st.file_uploader(
                "광고관리자 파일 업로드 (플랫폼별, 연령별 등 — 선택사항)",
                type=["csv", "xlsx"],
                accept_multiple_files=True
            )

    # ── 분석 시작 ───────────────────────────────────────────────
    if st.session_state.get("_selection_done") and st.button("분석 시작") and hypothesis:
        api_key = os.getenv("ANTHROPIC_API_KEY") or st.secrets.get("ANTHROPIC_API_KEY", "")
        if not api_key or api_key == "여기에_API_키_입력":
            st.error("API 키가 설정되지 않았습니다. Streamlit Cloud → Settings → Secrets에 ANTHROPIC_API_KEY를 추가해주세요.")
        else:
            import json
            client = anthropic.Anthropic(api_key=api_key)

            db_parts = []
            found_names, not_found_names = [], []
            groups_state = st.session_state.get("_groups")

            _filtered_cache = st.session_state.get("_filtered_cache", {})

            def fetch_creative(key, label):
                if key in _filtered_cache:
                    _df = _filtered_cache[key]
                elif groups_state and key in groups_state:
                    _g = groups_state[key]
                    _sel = st.session_state.get(f"_ms_{key}", list(_g.keys()))
                    _names = [n for gk in _sel for n in _g.get(gk, [])]
                    _df = get_creative_data_filtered(_names)
                else:
                    _df = get_creative_data(label)
                if not _df.empty:
                    db_parts.append(f"\n\n[DB raw data — {label}]\n{_df.to_string(index=False)}")
                    found_names.append(label)
                else:
                    not_found_names.append(label)

            with st.spinner("DB 조회 중..."):
                if creative_name_input:
                    fetch_creative("__main__", creative_name_input)

                if groups_state:
                    for _key in groups_state:
                        if _key != "__main__":
                            fetch_creative(_key, _key)
                else:
                    # 소재 확인 안 눌렀을 때 fallback
                    _bracketed = _re.findall(r'\[([^\]]+)\]', hypothesis)
                    _comp_names = [
                        b.replace('/', '').strip() for b in _bracketed
                        if b.replace('/', '').strip() and
                           b.replace('/', '').strip() != creative_name_input.strip()
                    ]
                    if not _comp_names:
                        try:
                            _resp = client.messages.create(
                                model="claude-haiku-4-5-20251001",
                                max_tokens=300,
                                messages=[{"role": "user", "content": f"""아래 가설에서 비교 대상 소재명을 JSON 배열로만 출력해.
슬래시(/)는 제거. 설명 없이 JSON만.
가설: {hypothesis}"""}]
                            )
                            _comp_names = json.loads(_resp.content[0].text.strip())
                            if not isinstance(_comp_names, list):
                                _comp_names = []
                        except Exception:
                            _comp_names = []
                    for _name in _comp_names:
                        if _name not in found_names:
                            fetch_creative(_name, _name)

            if found_names:
                st.success(f"DB 조회 완료: {', '.join(found_names)}")
            if not_found_names:
                st.warning(f"DB에 없는 소재 (해당 TEST 분석 제외): {', '.join(not_found_names)}")

            csv_text = ""
            if uploaded_csv:
                for _f in uploaded_csv:
                    if _f.name.endswith(".xlsx"):
                        _df = pd.read_excel(_f, engine="openpyxl")
                    else:
                        _df = pd.read_csv(_f, encoding="utf-8-sig")
                    csv_text += f"\n\n[{_f.name}]\n{_df.to_string(index=False)}"

            db_data_text = "".join(db_parts)

            # Fix B: 세트명 컨텍스트 — Claude가 헤더에 세트명 그대로 쓰도록 명시
            set_ctx_lines = []
            if creative_name_input:
                set_ctx_lines.append(f"주인공 세트명: [{creative_name_input.strip()}]")
            if groups_state:
                for _key in groups_state:
                    if _key != "__main__":
                        set_ctx_lines.append(f"비교 세트명: [{_key}]")
            set_context = (
                "\n\n[세트명 안내 — 결론 헤더에 아래 세트명 그대로 사용. 개별 인트로명 나열 절대 금지]\n"
                + "\n".join(set_ctx_lines)
            ) if set_ctx_lines else ""

            user_message = f"가설:\n{hypothesis}{set_context}{db_data_text}{csv_text}"

            result = ""
            stop_reason = None
            with st.spinner("분석 중..."):
                with client.messages.stream(
                    model="claude-sonnet-4-6",
                    max_tokens=16000,
                    thinking={"type": "enabled", "budget_tokens": 8000},
                    system=build_system_prompt(),
                    messages=[{"role": "user", "content": user_message}]
                ) as stream:
                    for text in stream.text_stream:
                        result += text
                    stop_reason = stream.get_final_message().stop_reason

            if stop_reason == "max_tokens":
                st.warning("⚠️ 응답이 최대 토큰에 도달해 결론이 잘렸습니다. 가설의 TEST 수를 줄이거나 소재를 단순화해주세요.")

            known_names = []
            _groups_state = st.session_state.get("_groups", {})
            for _k, _gmap in _groups_state.items():
                label = creative_name_input if _k == "__main__" else _k
                known_names.append(label)
            result, _fixes = correct_creative_names(result, known_names)
            if _fixes:
                st.caption(f"소재명 자동 교정 {len(_fixes)}건: " + ", ".join(f"{w} → {c}" for w, c in _fixes.items()))

            st.markdown("### 분석 결론")
            _START = "[RESULT_START]"
            _END = "[RESULT_END]"
            if _START in result:
                _out = result[result.index(_START) + len(_START):]
            else:
                _out = result
            if _END in _out:
                _out = _out[:_out.index(_END)]
            _out = _out.strip()
            # 봄·가을 포함 계절 메모 줄 제거
            import re as _re2
            _lines = _out.split('\n')
            _lines = [
                l for l in _lines
                if not (l.strip().startswith('**') and ('봄' in l or '가을' in l))
            ]
            _out = '\n'.join(_lines).strip()
            if _out:
                st.markdown(f"```\n{_out}\n```")

# ── 탭 5: 분석 규칙 ─────────────────────────────────────────────
with tab5:
    # CSS 주입 (아티팩트 스타일)
    st.markdown("""<style>
    .ax-prereq{background:#f0f4ff;border:1.5px solid #c7d2fe;border-radius:10px;padding:16px 20px;margin-bottom:14px;}
    .ax-prereq-title{font-size:11px;font-weight:700;letter-spacing:0.5px;color:#1e40af;margin-bottom:10px;text-transform:uppercase;}
    .ax-prereq ul{list-style:none;margin:0;padding:0;}
    .ax-prereq ul li{font-size:12.5px;padding:3px 0 3px 16px;position:relative;line-height:1.65;}
    .ax-prereq ul li::before{content:"★";position:absolute;left:0;font-size:9px;color:#1e40af;top:5px;}
    .ax-prereq ul li strong{color:#1e40af;}
    .ax-prereq code{background:#dbeafe;color:#1e40af;padding:1px 5px;border-radius:3px;font-size:11.5px;}
    .ax-prereq ul.ax-sub{margin-top:3px;padding-left:14px;}
    .ax-prereq ul.ax-sub li{font-size:11.5px;color:#4b5563;padding:1px 0 1px 12px;}
    .ax-prereq ul.ax-sub li::before{content:"–";font-size:12px;color:#9ca3af;top:2px;}
    .ax-jkey{display:flex;gap:18px;flex-wrap:wrap;background:white;border:1px solid #e2e8f0;border-radius:8px;padding:11px 18px;margin-bottom:14px;align-items:center;}
    .ax-j{display:inline-flex;align-items:center;gap:6px;font-size:12.5px;color:#4b5563;}
    .ax-sym{font-weight:700;font-size:16px;}
    .ax-sym-o{color:#059669;}.ax-sym-x{color:#dc2626;}.ax-sym-tri{color:#d97706;}.ax-sym-q{color:#6b7280;}
    .ax-pnote{font-size:12px;color:#9ca3af;margin-left:auto;}
    .ax-hyp{background:white;border:1px solid #e2e8f0;border-radius:10px;margin-bottom:12px;overflow:hidden;}
    .ax-hyp.ax-unclass{border-color:#e5e7eb;}
    .ax-hyp-head{padding:12px 20px;background:#fafbfc;border-bottom:1px solid #e2e8f0;display:flex;align-items:center;gap:10px;}
    .ax-hyp.ax-unclass .ax-hyp-head{background:#f9fafb;}
    .ax-hyp-badge{font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;background:#1e40af;color:white;padding:3px 8px;border-radius:3px;white-space:nowrap;flex-shrink:0;}
    .ax-hyp.ax-unclass .ax-hyp-badge{background:#6b7280;}
    .ax-hyp-name{font-size:13.5px;font-weight:600;line-height:1.3;}
    .ax-hyp-trigger{font-size:11.5px;color:#4b5563;font-style:italic;margin-top:2px;}
    .ax-ad-tag{display:inline-flex;align-items:center;font-size:10px;font-weight:600;background:#e0f2fe;color:#0369a1;border-radius:3px;padding:2px 7px;margin-left:8px;vertical-align:middle;}
    .ax-hyp-body{padding:16px 20px;display:grid;grid-template-columns:1fr 1fr;gap:18px;}
    @media(max-width:700px){.ax-hyp-body{grid-template-columns:1fr;}}
    .ax-crit h4,.ax-samp h4{font-size:10px;font-weight:700;letter-spacing:0.6px;text-transform:uppercase;color:#9ca3af;margin-bottom:8px;}
    .ax-crit ul{list-style:none;margin:0;padding:0;}
    .ax-crit ul li{font-size:12.5px;padding:2px 0 2px 14px;position:relative;line-height:1.6;}
    .ax-crit ul li::before{content:"·";position:absolute;left:0;color:#1e40af;font-weight:900;}
    .ax-crit ul li strong{color:#1e40af;}
    .ax-crit ul li.warn{color:#dc2626;font-weight:600;}
    .ax-crit ul li.info{color:#0369a1;font-weight:600;}
    .ax-crit code{background:#dbeafe;color:#1e40af;padding:1px 5px;border-radius:3px;font-size:11px;}
    .ax-code{background:#0d1117;border-radius:7px;padding:13px 15px;overflow-x:auto;}
    .ax-code pre{font-family:'SF Mono','Consolas','Monaco',monospace;font-size:11px;color:#e6edf3;line-height:1.7;white-space:pre;margin:0;}
    .ax-section-title{font-size:15px;font-weight:700;color:#0d1117;margin:20px 0 14px;display:flex;align-items:center;gap:12px;}
    .ax-section-title::after{content:'';flex:1;height:1px;background:#e2e8f0;}
    </style>""", unsafe_allow_html=True)

    st.markdown('<div class="ax-section-title">가설 유형별 분석 기준</div>', unsafe_allow_html=True)

    # ── 대전제 ──
    st.markdown("""
<div class="ax-prereq">
  <div class="ax-prereq-title">모든 가설의 대전제</div>
  <ul>
    <li><strong>TEST 제목</strong>: 가설 텍스트 그대로 사용.</li>
    <li><strong>인트로 번호</strong>: v1~v10까지 존재 가능.</li>
    <li><strong>전환 1건</strong>: 1건짜리 소재가 효율이 <strong>더 좋아 보일 때</strong> → 신뢰도 낮음 (△). <strong>더 나쁠 때</strong> → 신뢰도 높음 (X 확정).</li>
    <li><strong>주인공 합산 기준</strong>: 비교 세트가 특정 인트로(v3 등)를 지목하더라도, 주인공 세트에 인트로 언급이 없으면 주인공은 세트 합산으로 표기. 비교 상대의 단위가 주인공 단위를 바꾸지 않음.</li>
    <li><strong>인트로 정렬</strong>: 인트로 데이터 줄은 반드시 번호 오름차순(v1→v2→v3→...)으로 표기. 효율·지출 순 정렬 금지 (순위 블록 제외).</li>
    <li><strong>마지막 줄</strong>: 중복되지 않는 중요한 사항만 요약. 해당 내용 없으면 기재하지 않음.</li>
    <li><strong>판정 표기</strong>: O/X/△로 답변되는 가설 → O/X/△ 표기. △면 괄호 안에 이유 기재.
      <ul class="ax-sub">
        <li>예: <code>= △ (전환 1건으로 신뢰도 낮음)</code></li>
        <li>O/X/△로 답변 안 되는 가설 → 결과 직접 기재 또는 2~12자 간단 서술. 예: <code>= v3(참여형)</code></li>
      </ul>
    </li>
    <li><strong>금액 표기</strong>: <code>100K</code>, <code>500K</code> 등 알파벳 단위 금지. <code>1.3만원</code>, <code>100만원</code> 형태 사용.</li>
    <li><strong>계절 차이 메모</strong>: 여름(6~8월)/겨울(12~2월) 차이 날 때만 결론 최하단에 메모. 봄·가을 기간 소재는 메모 출력 자체 금지.</li>
    <li><strong>프모 이름·월 언급 금지</strong>: 결론에서 프로모션 명칭·특정 월 출력 금지. "프모" / "비프모"로만 표현.</li>
    <li><strong>프모 조건 fallback</strong>: 비교 세트에서 프모 조건 맞는 날이 <strong>2일 미만</strong>이면 → 프모 조건 버리고 전체 기간에서 동일지출 기준으로 매칭. 결론에 반드시 <code>단, [소재명1]은 프모 / [소재명2]는 비프모</code> 한 줄 추가.</li>
  </ul>
</div>
""", unsafe_allow_html=True)

    # ── 판정 기호 ──
    st.markdown("""
<div class="ax-jkey">
  <span class="ax-j"><span class="ax-sym ax-sym-o">O</span> 가설 충족</span>
  <span class="ax-j"><span class="ax-sym ax-sym-x">X</span> 가설 미충족</span>
  <span class="ax-j"><span class="ax-sym ax-sym-tri">△</span> 조건부 / 방향성만</span>
  <span class="ax-j"><span class="ax-sym ax-sym-q">?</span> 판정 불가</span>
  <span class="ax-pnote">전환 지표 우선순위: CPA &gt; ROAS &gt; CVR</span>
</div>
""", unsafe_allow_html=True)

    # ── 유형 카드 ──
    st.markdown("""
<div class="ax-hyp">
  <div class="ax-hyp-head">
    <span class="ax-hyp-badge">유형 1</span>
    <div><div class="ax-hyp-name">세트 간 효율 비교</div>
    <div class="ax-hyp-trigger">"[비교소재]보다 효율이 더 좋을까?", "효율이 더 우수할까?"</div></div>
  </div>
  <div class="ax-hyp-body">
    <div class="ax-crit"><h4>분석 기준</h4><ul>
      <li>전환 지표(CPA·ROAS·CVR) 중심 비교</li>
      <li><strong>동일지출 + 프모 조건 일치</strong> 기간 기준</li>
      <li>프모 불일치 시 괄호로 간단 언급. 2일 미만 매칭 시 → 프모 조건 버리고 동일지출 기준 매칭 + 결론에 <code>단, [소재명1]은 프모 / [소재명2]는 비프모</code> 추가</li>
      <li>결과는 <strong>주인공 기준</strong>으로 서술 ("아쉽다")</li>
      <li>각 소재 줄: 캠페인 + 소재명 + 기간 + 지출 + CPA (전환 N건)</li>
      <li>효율 아쉽더라도 주인공 광고비·매출이 <strong>둘 다</strong> 100만원 이상 높으면: <code>전환 지표 자체는 아쉽지만 광고비 N만원·매출 M만원 더 우수</code></li>
    </ul></div>
    <div class="ax-samp"><h4>결론 샘플</h4>
    <div class="ax-code"><pre>TEST 1. 왜안해디슬보다 효율이 더 좋을까? = X (아쉽다)

**휴터 [왜안휴에이슬] vs 듀플 [왜안해디슬]
(비프모 &amp; 120만원 내외 동일지출 기준)
[왜안휴에이슬](휴터) (26.06.08~06.23)
  1,187,523원 / CPA 63,480원 (전환 19건)
[왜안해디슬](듀플) (26.06.08~06.15)
  1,204,817원 / CPA 44,236원 (전환 27건)
→ [왜안해디슬]보다 CPA 19,244원 아쉽, ROAS 34% 아쉽</pre></div>
    </div>
  </div>
</div>

<div class="ax-hyp">
  <div class="ax-hyp-head">
    <span class="ax-hyp-badge">유형 2</span>
    <div><div class="ax-hyp-name">인트로 위너</div>
    <div class="ax-hyp-trigger">"위너 인트로는?", "어느 인트로가 효율이 좋을지?"</div></div>
  </div>
  <div class="ax-hyp-body">
    <div class="ax-crit"><h4>판정 순서</h4><ul>
      <li>① 1·2위 지출 차 ≥ 50만원 → 위너 확정</li>
      <li>② 15만원 이상 인트로 중 CPA+ROAS 둘 다 최우수 → 위너 확정</li>
      <li>③ CPA 또는 ROAS 하나만 최우수 → △</li>
      <li>전부 0건 → ? (판정불가)</li>
      <li class="warn">개별 인트로 줄에는 O/X/△ 없음 (지출/CPA 순위 블록이 판정 역할)</li>
      <li class="info">예외: "전환이 가장 좋을까?" → <strong>유형 6</strong> / "후킹이 가장 좋을까?" → <strong>유형 5</strong></li>
    </ul></div>
    <div class="ax-samp"><h4>결론 샘플</h4>
    <div class="ax-code"><pre>TEST 2. 위너 인트로는? = v3(참여형)
선화렉카슬 2026.06.08~06.26 총 1,205,340원 지출

**지출 순위
1위_v3/참여형 (523,412원) → 지출 위너 v3
2위_v1/필요성 (381,220원)
3위_v2/할인율 (300,708원)

**CPA 순위
1위_v3/참여형 (52,341원) → CPA 위너 v3
2위_v1/필요성 (71,432원)
3위_v2/할인율 (전환 없음)</pre></div>
    </div>
  </div>
</div>

<div class="ax-hyp">
  <div class="ax-hyp-head">
    <span class="ax-hyp-badge">유형 3</span>
    <div><div class="ax-hyp-name">콘셉트 그룹 비교 (A vs B)</div>
    <div class="ax-hyp-trigger">"몸매 vs 대세감 중 어느 쪽이 유의미한지?", "X그룹 vs Y그룹?"</div></div>
  </div>
  <div class="ax-hyp-body">
    <div class="ax-crit"><h4>분석 기준</h4><ul>
      <li>인트로를 두 콘셉트 그룹으로 분류 → <strong>그룹 합산 CPA/ROAS</strong> 기준 판정</li>
      <li>제목: "콘셉트A vs 콘셉트B?" 형식 · 인트로 번호: v1(몸매1), v3(대세감1) 등 그룹명+순번</li>
      <li>두 그룹 모두 전환: 합산 CPA·ROAS 둘 다 우수 → 위너 확정 / 지표 분산 → △</li>
      <li>한 그룹만 전환: <code>= ? (판정불가)</code> + 방향성 한 줄 (전환 1건이면 신뢰도 낮음)</li>
      <li>두 그룹 모두 전환 0건: <code>= ? (판정불가)</code></li>
    </ul></div>
    <div class="ax-samp"><h4>결론 샘플</h4>
    <div class="ax-code"><pre>TEST 1. 몸매 vs 대세감? = ? (판정불가)
선화렉카투슬 2026.06.15~06.22 총 404,731원 지출

v1(몸매1) 62,557원 / (전환 0건)
v2(몸매2) 123,831원 / (전환 0건)
v3(대세감1) 79,891원 / (전환 0건)
v4(대세감2) 138,453원 / CPA 13.8만원 (전환 1건) → △</pre></div>
    </div>
  </div>
</div>

<div class="ax-hyp">
  <div class="ax-hyp-head">
    <span class="ax-hyp-badge">유형 4</span>
    <div><div class="ax-hyp-name">효율 복사</div>
    <div class="ax-hyp-trigger">"원본의 X%는 효율이 복사될까?", "효율이 복사되지 않을까?"</div></div>
  </div>
  <div class="ax-hyp-body">
    <div class="ax-crit"><h4>분석 기준</h4><ul>
      <li>판정 기준: CPA·ROAS가 아닌 <strong>지출 또는 매출 절대 규모</strong></li>
      <li><strong>가설에 기준 % 있음</strong>: 주인공 ÷ 원본 ≥ 기준 % → O / 미달 → X. O/X/△ 사용.</li>
      <li><strong>가설에 기준 % 없음</strong>: O/X/△ 사용 금지. N &gt; 15 → <code>N% 복사됨</code> / N ≤ 15 → <code>N%만 아쉽게 복사</code></li>
    </ul></div>
    <div class="ax-samp"><h4>결론 샘플</h4>
    <div class="ax-code"><pre>TEST 1. 원본의 50%는 효율 복사될까? = O
선화렉카슬 2026.05.20~06.10 총 2,541,700원 지출
선화렉카원슬 2026.03.01~03.31 라이브

* vs [선화렉카원슬]
(비프모 &amp; 300만원 내외 동일지출 기준)
[선화렉카원슬] 3,012,400원 지출
→ 지출 84% 복사</pre></div>
    </div>
  </div>
</div>

<div class="ax-hyp">
  <div class="ax-hyp-head">
    <span class="ax-hyp-badge">유형 5</span>
    <div><div class="ax-hyp-name">후킹지표 분석 <span class="ax-ad-tag">광고관리자 필요</span></div>
    <div class="ax-hyp-trigger">"후킹지표가 개선되는지?", "CTR이 더 높을까?", "후킹이 가장 좋을까?"</div></div>
  </div>
  <div class="ax-hyp-body">
    <div class="ax-crit"><h4>분석 기준</h4><ul>
      <li>광고관리자 <strong>instagram 데이터만</strong> 사용</li>
      <li>CTR·CPC·CPM 중심 분석. 오디언스 네트워크 언급 금지</li>
      <li>전체가 후킹 TEST면 "후킹지표:" 라벨 생략, 수치 바로 표기</li>
      <li>출처: <strong>광고관리자_instagram</strong> — 비교 블록 헤더에만 한 번 명시. 개별 인트로 줄에 중복 금지</li>
      <li>기간 매칭 — 인트로 간: 지출 적은 인트로 기준 비슷한 지출 구간 선별</li>
      <li class="warn">지출 순위 블록 없음 (후킹 지표 수치만 표기)</li>
    </ul></div>
    <div class="ax-samp"><h4>결론 샘플</h4>
    <div class="ax-code"><pre>TEST 3. [선화렉카슬]보다 후킹지표 개선되는지? = X

**[선화렉카투슬] vs [선화렉카슬]
(비프모 &amp; 40만원 내외 지출 기준, 광고관리자_instagram)
[선화렉카투슬] (26.06.15~06.22) : 404,731원 / CTR 0.99% / CPC 5,008원
[선화렉카슬] (26.04.27~05.04) : 456,682원 / CTR 1.28% / CPC 2,047원
→ [선화렉카슬]보다 CTR 0.29% 아쉽, CPC 2,961원 아쉽</pre></div>
    </div>
  </div>
</div>

<div class="ax-hyp">
  <div class="ax-hyp-head">
    <span class="ax-hyp-badge">유형 6</span>
    <div><div class="ax-hyp-name">전환지표 비교 분석</div>
    <div class="ax-hyp-trigger">"전환에 더 유리한지?", "전환 지표는 어떤 게 제일 좋은지?", "전환이 가장 좋을까?"</div></div>
  </div>
  <div class="ax-hyp-body">
    <div class="ax-crit"><h4>분석 기준</h4><ul>
      <li>CPA·ROAS 둘 다 우수 → O / 하나만 → △ / 전부 0건 → ? (판정불가)</li>
      <li><strong>인트로 간 비교 — 지출 슬라이싱:</strong> 지출 최소 인트로 = 기준 지출. 나머지 인트로는 누적 지출이 기준 지출에 도달하는 시점까지 슬라이싱</li>
      <li>슬라이싱으로 날짜가 달라지므로 <strong>각 인트로 줄에 날짜 개별 표기</strong> (헤더 공통 날짜 방식 금지)</li>
      <li>세트 간 비교: 동일 프모 여부 조건 &amp; 비슷한 지출 구간 매칭</li>
      <li>인트로 3개 이상이면 <code>v1 vs v2 vs v3</code> 형태 헤더</li>
      <li class="warn">지출 순위 블록 없음 (전환 지표 수치만 표기)</li>
    </ul></div>
    <div class="ax-samp"><h4>결론 샘플 (인트로 3-way)</h4>
    <div class="ax-code"><pre>TEST 2. 어떤 멘트의 인트로가 전환이 가장 좋을까? = △ (v3_cpa위너)

**v1(유형) vs v2(유형) vs v3(유형) (비프모 &amp; 9만원 내외 지출 기준)
v1(유형) (26.05.25~05.28) : 89,XXX원 지출 / (전환 0건)
v2(유형) (26.05.25~06.04) : 89,368원 지출 / (전환 0건)
v3(유형) (26.05.25~05.26) : 89,XXX원 지출 / CPA N원 (전환 N건) / ROAS N%
→ v3 유일 전환. 전체 모수 적어 신뢰도 낮음 → △</pre></div>
    </div>
  </div>
</div>

<div class="ax-hyp">
  <div class="ax-hyp-head">
    <span class="ax-hyp-badge">유형 7</span>
    <div><div class="ax-hyp-name">연령별 반응 분석 <span class="ax-ad-tag">광고관리자 필요</span></div>
    <div class="ax-hyp-trigger">"어떤 연령대에게 더 반응이 있는지?", "25-34세에게 잘 먹힐까?"</div></div>
  </div>
  <div class="ax-hyp-body">
    <div class="ax-crit"><h4>분석 기준</h4><ul>
      <li>광고관리자 연령별 데이터 업로드 필요</li>
      <li>소재별 한 줄 서술형: 지출 집중 연령 + 전환 성과</li>
      <li>수치 비중(%) 나열 대신 문장 형태로 표기</li>
    </ul></div>
    <div class="ax-samp"><h4>결론 샘플</h4>
    <div class="ax-code"><pre>TEST 2. 어떤 연령대에게 반응이 있는지?
= 25-34 위주로 전환 발생
선화렉카슬 2026.06.08~06.26 총 985,708원 지출

(선화렉카슬/26.06.08~06.26, 광고관리자)
v1(대세감): 25-34에 지출·전환 집중 (CPA 52,341원)
v2(몸매): 18-24에 지출됐으나 전환 저조,
          35-44 소량 지출 전환 없음
v3(참여형): 25-34 지출 많으나 전환 모수 적음</pre></div>
    </div>
  </div>
</div>

<div class="ax-hyp">
  <div class="ax-hyp-head">
    <span class="ax-hyp-badge">유형 8</span>
    <div><div class="ax-hyp-name">전환 여부 (실험 소재)</div>
    <div class="ax-hyp-trigger">"일단 전환이 나올까?", "전환이 가능한지?"</div></div>
  </div>
  <div class="ax-hyp-body">
    <div class="ax-crit"><h4>판정 기준</h4><ul>
      <li>세트 단위 아닌 <strong>인트로별</strong> 분석</li>
      <li>CPA ≤ 10만원 → O</li>
      <li>10만원 &lt; CPA ≤ 20만원 → △ (살짝 아쉽)</li>
      <li>CPA &gt; 20만원 → △ (매우 아쉽)</li>
      <li>전환 0건 → X</li>
      <li>인트로 중 하나라도 O면 전체 판정 = O</li>
    </ul></div>
    <div class="ax-samp"><h4>결론 샘플</h4>
    <div class="ax-code"><pre>TEST 1. 일단 전환이 나올까? = △

선화렉카슬 2026.06.08~06.26 총 985,708원 지출

v1(대세감) 341,220원 / CPA 6.8만원 (전환 5건) → O
v2(몸매) 412,830원 / CPA 13.8만원 (전환 3건) → △ (살짝 아쉽)
v3(참여형) 231,658원 / (전환 0건) → X</pre></div>
    </div>
  </div>
</div>

<div class="ax-hyp ax-unclass">
  <div class="ax-hyp-head">
    <span class="ax-hyp-badge">분류 불가</span>
    <div><div class="ax-hyp-name">분류 불가 가설</div>
    <div class="ax-hyp-trigger">유형1~8 어느 기준에도 해당하지 않는 가설</div></div>
  </div>
  <div class="ax-hyp-body">
    <div class="ax-crit"><h4>분석 기준</h4><ul>
      <li>결론 제목에 <code>(분류 불가)</code> 표기</li>
      <li><strong>가설 텍스트 내 기준·비교 대상·판정 조건</strong>을 직접 해석하여 분석</li>
      <li>O/X/△: 가설에 명시적 기준 있으면 사용, 없으면 수치 서술로 대체</li>
    </ul></div>
    <div class="ax-samp"><h4>결론 샘플</h4>
    <div class="ax-code"><pre>TEST 3. 지출 대비 구매가 많은 인트로는? (분류 불가)
= v2 (구매효율 최우수)

선화렉카슬 2026.06.08~06.26 총 985,708원 지출

v1 341,220원 / 전환 5건 → CPA 6.8만원
v2 412,830원 / 전환 9건 → CPA 4.6만원
v3 231,658원 / 전환 0건</pre></div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

    # ── 추가 규칙 섹션 (SYSTEM_PROMPT 파싱) ──
    import re as _re_sp
    _sp_raw = SYSTEM_PROMPT
    _sp_sections = {}
    _sp_parts = _re_sp.split(r'\n(?=## )', _sp_raw)
    for _part in _sp_parts:
        _part = _part.strip()
        if _part.startswith('## '):
            _title_end = _part.index('\n') if '\n' in _part else len(_part)
            _sec_title = _part[3:_title_end].strip()
            _sec_body = _part[_title_end:].strip()
            _sp_sections[_sec_title] = _sec_body

    st.markdown('<div class="ax-section-title">세부 규칙 원문</div>', unsafe_allow_html=True)
    _other_sections = [
        ("⚖️ 세트 간 효율 판정 기준", "세트 간 효율 판정 기준"),
        ("📅 기간 매칭 원칙", "기간 매칭 원칙"),
        ("📊 세트 간 비교 블록 포맷", "세트 간 비교 블록 포맷"),
        ("🎬 인트로별 분석 포맷", "인트로별 분석 포맷"),
        ("📋 연령별 분석 포맷", "연령별 분석 포맷"),
        ("✅ 정답 예시", "정답 예시"),
    ]
    for _label, _key_prefix in _other_sections:
        _matched = next((v for k, v in _sp_sections.items() if k.startswith(_key_prefix)), None)
        if _matched:
            with st.expander(_label):
                st.markdown(_matched)

    st.divider()
    st.subheader("전체 규칙 원문")
    st.caption("app.py의 SYSTEM_PROMPT 전문. 앱 재시작 시 항상 최신 규칙이 반영됩니다.")
    st.markdown(f"```\n{SYSTEM_PROMPT}\n```")
