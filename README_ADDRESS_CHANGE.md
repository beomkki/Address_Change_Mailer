# Address Change Mail Merge

국가별로 상표를 그룹핑하여 주소변경 안내 메일을 자동 생성하는 도구입니다.

## 개요

100개의 상표가 있어도 같은 국가면 **1개의 메일**로 통합하여 발송할 수 있습니다. Word 템플릿의 상표 테이블에 자동으로 행이 추가됩니다.

### 기존 Mail Merge와의 차이점

| 구분 | 기존 (Filing/Search) | 신규 (Address Change) |
|------|---------------------|----------------------|
| **메일 개수** | 1행 = 1메일 | 1국가 = 1메일 |
| **예시** | 100개 상표 = 100개 메일 | 100개 상표 (같은 국가) = 1개 메일 |
| **테이블** | 고정 내용 | 동적으로 행 추가 |
| **사용 케이스** | 출원/검색 안내 | 대량 주소변경 |

## 필요 파일

### 1. List of Marks.xlsx (상표 리스트)

**필수 컬럼:**
- `Country Code`: 국가 코드 (예: OA, VN, ZA)
- `Mark`: 상표명
- `Class`: 상표 분류
- `Appl. Date`: 출원일
- `Appl. No.`: 출원번호
- `Reg. Date`: 등록일 (선택)
- `Reg. No.`: 등록번호 (선택)

**선택 컬럼 (수신인 정보):**
- `수신` 또는 `To`: 수신인 이메일 주소
- `참조` 또는 `CC`: 참조 이메일 주소
- `국가명칭` 또는 `Country Name`: 국가명

**예시:**

| Country Code | Mark | Class | Appl. Date | Appl. No. | 수신 | 참조 | 국가명칭 |
|--------------|------|-------|------------|-----------|------|------|----------|
| OA | DORCO | 21 | 2025-06-13 | 3202502194 | agent@example.com | cc@example.com | OAPI |
| OA | DORCO | 8 | 2012-01-23 | 3201200185 | agent@example.com | cc@example.com | OAPI |
| VN | DORCO | 21 | 2024-01-15 | VN123456 | vietnam@example.com | cc@example.com | Vietnam |

### 2. 메일링 리스트.xlsx (대리인 매핑)

`List of Marks.xlsx`에 수신인 정보가 없을 때 사용됩니다.

**Sheet1 구조:**
- 컬럼 1: 국가코드
- 컬럼 3: 국가명칭
- 컬럼 4: 발신인
- 컬럼 5: 수신
- 컬럼 6: 참조

**Sheet2 구조:**
- 컬럼 1: 국가코드
- 컬럼 2: 국가명칭
- 컬럼 3: 발신인
- 컬럼 4: 수신
- 컬럼 5: 참조

### 3. Address_Change_Mail_Sample.docx (템플릿)

Word 템플릿의 **첫 번째 테이블**에 상표 데이터가 자동으로 채워집니다.

테이블 구조:
```
| Mark | Class | Appl. Date | Appl. No. | Reg. Date | Reg. No. |
|------|-------|------------|-----------|-----------|----------|
| (자동 채워짐)                                                    |
```

## 수신인 정보 우선순위

스크립트는 다음 순서로 수신인 정보를 찾습니다:

1. **우선순위 1**: `List of Marks.xlsx`의 `수신`/`참조`/`국가명칭` 컬럼
2. **우선순위 2**: `메일링 리스트.xlsx`에서 국가코드로 조회

### 권장 방법

**방법 A: List of Marks.xlsx에 수신인 정보 포함 (권장)**
- 장점: 하나의 파일로 관리, 국가별로 다른 대리인 지정 가능
- 단점: 컬럼이 더 많아짐

**방법 B: 메일링 리스트.xlsx 사용**
- 장점: 상표 리스트가 간결함
- 단점: 두 파일 관리 필요, 국가당 1개 대리인만 가능

**방법 C: 혼용**
- 일부 국가는 List of Marks에, 일부는 메일링 리스트에 정의
- 스크립트가 자동으로 우선순위에 따라 선택

## 사용법

### 방법 1: GUI 버전 (권장 - 초보자용)

```bash
python address_change_mail_gui.py
```

GUI 창이 열리면:
1. 각 파일 경로가 자동으로 입력됨 (기본값)
2. 필요시 **찾기** 버튼으로 파일 경로 변경
3. **메일 생성** 버튼 클릭
4. 완료 메시지 확인

**기본값:**
- 상표 리스트: `List of Marks.xlsx`
- 메일링 리스트: `메일링 리스트.xlsx`
- 템플릿: `Address_Change_Mail_Sample.docx`
- 출력: `output-address-change/`

### 방법 2: CLI 버전 (고급 사용자용)

기본 실행:

```bash
python generate_address_change_mail.py
```

커스텀 파일 지정:

```bash
python generate_address_change_mail.py \
  --marks-excel "My Trademark List.xlsx" \
  --mailing-list "My Recipients.xlsx" \
  --template "My Template.docx" \
  --output-dir "my-output"
```

### 옵션

| 옵션 | 기본값 | 설명 |
|------|--------|------|
| `--marks-excel` | `List of Marks.xlsx` | 상표 리스트 파일 |
| `--mailing-list` | `메일링 리스트.xlsx` | 대리인 매핑 파일 |
| `--template` | `Address_Change_Mail_Sample.docx` | 메일 템플릿 |
| `--output-dir` | `output-address-change` | MSG 파일 저장 디렉터리 |

## 실행 결과

### 콘솔 출력 예시

```
Loading trademark data...
Found 30 countries with 140 total marks

Loading recipient mapping...
Found 31 recipient mappings

Processing OA: 6 marks
  To: agent@example.com
  CC: cc@example.com
  ✓ Saved: OA_OAPI_AddressChange.msg

Processing VN: 20 marks
  To: vietnam@example.com
  CC: cc@example.com
  ✓ Saved: VN_Vietnam_AddressChange.msg

완료: 30개의 MSG 파일을 output-address-change에 생성했습니다.
```

### 생성된 MSG 파일

각 국가당 1개씩 생성됩니다:

- `OA_OAPI_AddressChange.msg` (6개 상표 포함)
- `VN_Vietnam_AddressChange.msg` (20개 상표 포함)
- `ZA_South_Africa_AddressChange.msg` (12개 상표 포함)
- ...

**메일 제목 형식:**
```
Address Change - Vietnam (VN) - 20 marks
```

**메일 본문:**
- Word 템플릿 내용 그대로
- 첫 번째 테이블에 해당 국가의 모든 상표 정보 자동 삽입

## 데이터 준비 가이드

### List of Marks.xlsx 준비하기

1. 기존 상표 리스트를 준비합니다
2. 다음 컬럼을 추가합니다:
   - `수신`: 국가별 대리인 이메일 (같은 국가는 같은 값)
   - `참조`: 참조 이메일 (같은 국가는 같은 값)
   - `국가명칭`: 국가 이름 (같은 국가는 같은 값)

**팁: Excel에서 빠르게 채우기**
1. 국가별로 정렬
2. 첫 행에 수신/참조 입력
3. 아래로 드래그하여 자동 채우기

### 메일링 리스트.xlsx 업데이트하기

1. `메일링 리스트.xlsx` 열기
2. Sheet1 또는 Sheet2에 누락된 국가 추가
3. 국가코드, 국가명, 수신, 참조 입력

## 문제 해결

### "No recipient info found" 경고

```
Warning: No recipient info found (neither in marks nor mailing list), skipping...
```

**원인:** 해당 국가의 수신인 정보가 없습니다.

**해결:**
- `List of Marks.xlsx`에 `수신` 컬럼 추가하거나
- `메일링 리스트.xlsx`에 해당 국가 추가

### 일부 국가만 MSG 생성됨

**원인:** 수신인 정보가 있는 국가만 생성됩니다.

**확인:**
- 콘솔 출력에서 "skipping..." 메시지 확인
- 해당 국가의 수신인 정보 추가

### 테이블이 비어있음

**원인:** 템플릿의 테이블 구조가 잘못되었거나 데이터 컬럼명 불일치

**확인:**
- 템플릿의 첫 번째 테이블 헤더가 다음과 일치하는지 확인:
  `Mark | Class | Appl. Date | Appl. No. | Reg. Date | Reg. No.`
- Excel의 컬럼명이 정확히 일치하는지 확인

## 비교: 기존 vs 신규

### 기존 Filing/Search Mail Merge
```python
python generate_mail_merge.py --excel Filing_Merge.xlsx --template Filing.docx
```
- 용도: 출원/검색 안내 개별 메일
- 1행 → 1메일
- 100행 → 100개 MSG 파일

### 신규 Address Change Mail Merge
```python
python generate_address_change_mail.py
```
- 용도: 주소변경 안내 그룹 메일
- 1국가 → 1메일
- 100개 상표 (10개 국가) → 10개 MSG 파일

## 개발 배경

**문제:**
- 100개 상표의 주소변경을 진행할 때
- 기존 방식: 동일 대리인에게 100개 메일 발송
- 대리인 입장에서 번거로움

**해결:**
- 국가별로 그룹핑하여 1개 메일로 통합
- 메일 본문에 상표 목록 테이블 자동 생성
- 대리인은 1개 메일로 모든 상표 확인 가능

## 라이선스

기존 Mail_Merge 프로젝트와 동일

## 작성자

특허법인 성암 - Claude Code
