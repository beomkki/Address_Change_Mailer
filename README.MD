# Mail Merge - Outlook 메일 자동 생성 도구

Excel 데이터와 Word 템플릿을 활용하여 Outlook 메일 초안(.msg 파일)을 자동으로 생성하는 Python 스크립트입니다.

## 📋 프로젝트 개요

해외 대리인들에게 국가별로 맞춤화된 메일을 대량으로 발송해야 할 때, 수작업으로 제목과 수신인을 일일이 입력하는 번거로움을 해결하기 위한 도구입니다.

### 주요 기능
- Excel 데이터를 기반으로 Word 템플릿의 병합 필드를 자동 치환
- 제목(Subject)도 병합 필드를 활용하여 건별로 자동 생성
- 수신(To), 참조(CC) 자동 설정
- **파일 자동 첨부 지원** (Excel에 첨부파일 경로 지정)
- Outlook에서 바로 열 수 있는 .msg 파일 생성
- 결재 및 검토를 위한 파일 형태로 저장

## 🚀 시작하기

### 1. 필요 패키지 설치

```bash
pip install docx2msg openpyxl python-docx pywin32
```

**주의**: `docx2msg`는 Windows에서 Outlook이 설치되어 있어야 작동합니다.

### 2. 파일 준비

프로젝트 폴더에 다음 파일들을 준비합니다:

- **Excel 파일** (예: `Filing_Merge.xlsx`, `Search_Merge.xlsx`)
  - 첫 번째 행: 컬럼명 (병합 필드명과 일치해야 함)
  - 이후 행: 각 메일에 해당하는 데이터
  - 필수 컬럼: `수신`, `참조` (변경 가능)

- **Word 템플릿** (예: `Filing.docx`, `Search.docx`)
  - 병합 필드는 `«필드명»` 형식으로 작성 (Word의 병합 필드 삽입 기능 사용)
  - 예: `«관리번호»`, `«국가명칭»`, `«수신»`, `«참조»` 등

### 3. 실행

#### 방법 1: GUI 버전 (권장 - 초보자용)

```bash
python mail_merge_gui.py
```

GUI 창이 열리면:
1. **메일 유형** 선택 (출원/검색)
2. 필요시 **찾기** 버튼으로 파일 경로 변경
3. 제목 템플릿 확인/수정 (기본값이 자동으로 입력됨)
4. **메일 생성** 버튼 클릭

#### 방법 2: CLI 버전 (고급 사용자용)

기본 설정으로 실행:
```bash
python generate_mail_merge.py
```

옵션을 지정하여 실행:
```bash
python generate_mail_merge.py --excel "Search_Merge.xlsx" --template "Search.docx" --output-dir "output-msg"
```

#### 방법 3: EXE 파일 (Python 설치 없이 사용)

EXE 파일을 더블클릭하면 GUI가 실행됩니다. Python이 설치되어 있지 않아도 사용 가능합니다.

## 📖 상세 가이드

### 1. Excel 파일 작성하기

Excel 파일의 첫 번째 행은 **컬럼명(필드명)**이며, 이 이름이 Word 템플릿의 병합 필드와 정확히 일치해야 합니다.

**예시: Filing_Merge.xlsx**

| 관리번호 | 국가코드 | 국가명칭 | 수신 | 참조 | 출원인 | 출원번호 | 첨부파일 |
|---------|---------|---------|------|------|--------|----------|----------|
| OT-253784 | EC | Ecuador | agent@example.com | manager@example.com | ABC Company | 2024-001 | files/doc1.pdf |
| OT-253785 | BR | Brazil | brasil@example.com | | XYZ Corp | 2024-002 | files/doc2.pdf;files/doc3.pdf |
| OT-253786 | MX | Mexico | mexico@example.com | cc@example.com | Test Inc | 2024-003 | |

**중요 포인트:**
- **첫 번째 행**: 컬럼명 (한글/영문 모두 가능)
- **두 번째 행부터**: 실제 데이터
- **필수 컬럼**: `수신`, `참조` (또는 `--to-field`, `--cc-field`로 지정한 컬럼명)
- **선택 컬럼**: `첨부파일` (메일에 자동으로 첨부할 파일 경로)
- **빈 셀**: 빈 칸으로 둬도 됩니다 (자동으로 빈 문자열로 처리)

#### 첨부파일 사용하기

Excel에 `첨부파일` 컬럼을 추가하여 메일에 자동으로 파일을 첨부할 수 있습니다.

**사용 방법:**
- **단일 파일**: `files/document.pdf`
- **여러 파일**: 세미콜론(`;`)으로 구분 - `files/doc1.pdf;files/doc2.pdf;files/doc3.pdf`
- **절대 경로**: `C:\Documents\file.pdf`
- **상대 경로**: 프로젝트 폴더 기준 - `attachments/sample.docx`

**예시:**
```
files/trademark_certificate.pdf;files/power_of_attorney.pdf
```

**주의사항:**
- 파일이 존재하지 않으면 경고 메시지가 표시되지만 메일 생성은 계속됩니다
- 상대 경로는 프로젝트 폴더 (`generate_mail_merge.py`가 있는 위치) 기준입니다
- 첨부파일 컬럼이 없거나 비어있으면 첨부파일 없이 메일이 생성됩니다

### 2. Word 템플릿에 병합 필드 삽입하기

Word 문서에서 병합 필드를 삽입하는 방법:

#### 방법 1: Word 메일 머지 기능 사용 (권장)

1. Word에서 템플릿 파일 열기 (예: `Filing.docx`)
2. 상단 메뉴: **메일링** 탭 클릭
3. **필드 삽입** 그룹에서 **병합 필드 삽입** 클릭
4. 드롭다운에서 원하는 필드명 선택 (예: `관리번호`)
5. 문서에 `«관리번호»` 형태로 삽입됨

**주의**: Word의 메일 머지 기능을 처음 사용할 때는 먼저 Excel 파일을 연결해야 드롭다운에 필드가 나타납니다:
- **메일링** → **받는 사람 선택** → **기존 목록 사용** → Excel 파일 선택

#### 방법 2: 직접 입력 (간편하지만 주의 필요)

병합 필드 기호 `«»`를 직접 입력할 수도 있습니다:

```
«필드명»
```

**⚠️ 중요: 기호 주의!**

병합 필드는 **꺾쇠 괄호 `«»`** 를 사용합니다. **부등호 `<<>>`가 아닙니다!**

- ✅ **올바른 기호**: `«관리번호»` (길러멧 기호, guillemet)
- ❌ **잘못된 기호**: `<<관리번호>>` (부등호 두 개)
- ❌ **잘못된 기호**: `<관리번호>` (부등호 한 개)

**기호 입력 방법:**
- `«` : Alt + 0171 (숫자 키패드 사용)
- `»` : Alt + 0187 (숫자 키패드 사용)

또는 여기서 복사해서 사용: `«»`

**팁**: 헷갈리면 Word의 메일 머지 기능(방법 1)을 사용하는 것을 권장합니다.

#### 병합 필드 예시

**메일 본문 템플릿:**

```
Dear «수신»,

We are pleased to inform you about the trademark application in «국가명칭».

Application Details:
- Reference Number: «관리번호»
- Country Code: «국가코드»
- Applicant: «출원인»
- Application Number: «출원번호»

Best regards,
```

### 3. 제목(Subject) 설정하기

이 도구의 **핵심 장점**은 메일 제목도 병합 필드를 사용해 자동으로 생성할 수 있다는 것입니다!

#### 기본 제목 템플릿

코드에 내장된 기본 제목 템플릿:
```
(«관리번호»«국가코드») New trademark application(s) in «국가명칭»
```

**결과 예시:**
- 데이터: `관리번호=OT-253784`, `국가코드=EC`, `국가명칭=Ecuador`
- 생성된 제목: `(OT-253784EC) New trademark application(s) in Ecuador`

#### 커스텀 제목 템플릿 사용

`--subject-template` 옵션으로 제목을 변경할 수 있습니다:

**예시 1: 간단한 제목**
```bash
python generate_mail_merge.py --subject-template "«관리번호» - «국가명칭» Trademark Filing"
```
→ 결과: `OT-253784 - Ecuador Trademark Filing`

**예시 2: 한글 제목**
```bash
python generate_mail_merge.py --subject-template "[«국가코드»] «출원인» 상표출원 안내"
```
→ 결과: `[EC] ABC Company 상표출원 안내`

**예시 3: 날짜 포함 (Excel에 날짜 컬럼이 있을 경우)**
```bash
python generate_mail_merge.py --subject-template "«출원일» - «국가명칭» Application"
```

#### 제목 템플릿 규칙

- 병합 필드는 `«필드명»` 형식 사용
- Excel 파일에 있는 컬럼명만 사용 가능
- 한글, 영문, 숫자, 특수문자 모두 사용 가능
- 제목은 .msg 파일명으로도 사용됨 (특수문자는 자동으로 `_`로 변환)

### 4. 실행 전 체크리스트

모든 준비가 끝났다면 실행 전에 확인하세요:

- [ ] Excel 파일의 첫 번째 행에 컬럼명이 정확히 입력되어 있는가?
- [ ] Word 템플릿의 병합 필드(`«필드명»`)가 Excel 컬럼명과 정확히 일치하는가?
- [ ] `수신` 컬럼에 이메일 주소가 제대로 입력되어 있는가?
- [ ] Excel과 Word 파일이 같은 폴더에 있는가?
- [ ] 필요한 패키지가 모두 설치되어 있는가? (`pip install docx2msg openpyxl python-docx pywin32`)

### 5. 실행 및 확인

```bash
python generate_mail_merge.py
```

성공하면 다음과 같이 출력됩니다:

```
Saved MSG: output-msg/(OT-253784EC)_New_trademark_application_s__in_Ecuador.msg
Saved MSG: output-msg/(OT-253785BR)_New_trademark_application_s__in_Brazil.msg
Saved MSG: output-msg/(OT-253786MX)_New_trademark_application_s__in_Mexico.msg
완료: 3개의 MSG 파일을 output-msg에 생성했습니다.
```

이제 `output-msg` 폴더의 .msg 파일을 Outlook에서 열어 확인하세요!

## 📝 사용 방법

### 기본 사용법

1. Excel 파일에 발송할 메일 목록 작성
2. Word 템플릿에 메일 본문 및 병합 필드 작성
3. 스크립트 실행
4. `output-msg` 폴더에 생성된 .msg 파일들을 Outlook에서 열어 확인
5. 필요시 수정 후 발송

### 커맨드 라인 옵션

| 옵션 | 기본값 | 설명 |
|------|--------|------|
| `--excel` | `Filing_Merge.xlsx` | Excel 데이터 파일 경로 |
| `--template` | `Filing.docx` | Word 템플릿 파일 경로 |
| `--output-dir` | `output-msg` | 생성된 MSG 파일 저장 디렉터리 |
| `--to-field` | `수신` | Excel에서 수신인(To)에 해당하는 컬럼명 |
| `--cc-field` | `참조` | Excel에서 참조(CC)에 해당하는 컬럼명 |
| `--subject-template` | (기본 템플릿) | 제목 템플릿 (병합 필드 사용 가능) |
| `--attachment-field` | `첨부파일` | Excel에서 첨부파일 경로에 해당하는 컬럼명 (세미콜론으로 구분) |

### 예제

**Filing 메일 생성:**
```bash
python generate_mail_merge.py --excel "Filing_Merge.xlsx" --template "Filing.docx"
```

**Search 메일 생성:**
```bash
python generate_mail_merge.py --excel "Search_Merge.xlsx" --template "Search.docx"
```

**커스텀 제목 템플릿 사용:**
```bash
python generate_mail_merge.py --subject-template "(«관리번호») 상표출원 관련 - «국가명칭»"
```

## 🔧 작동 원리

1. **Excel 데이터 로드**: `openpyxl`을 사용하여 Excel 파일을 읽고 각 행을 딕셔너리로 변환
2. **병합 필드 수집**: Word 템플릿에서 `«필드명»` 패턴을 찾아 필요한 필드 목록 추출
3. **템플릿 변환**: Word의 병합 필드(`«필드명»`)를 Jinja2 형식(`{{ 필드명 }}`)으로 변환
4. **제목 생성**: 제목 템플릿의 병합 필드를 데이터로 치환
5. **MSG 파일 생성**: `docx2msg`를 사용하여 렌더링된 본문과 제목, 수신인 정보가 포함된 .msg 파일 생성
6. **파일 저장**: 제목 기반으로 파일명을 생성하여 출력 디렉터리에 저장

## 📁 프로젝트 구조

```
Mail_Merge/
├── generate_mail_merge.py      # CLI 스크립트
├── mail_merge_gui.py            # GUI 스크립트
├── Filing.docx                  # Filing 메일 템플릿
├── Search.docx                  # Search 메일 템플릿
├── Filing_Merge.xlsx            # Filing 병합 데이터
├── Search_Merge.xlsx            # Search 병합 데이터
├── 메일링 리스트.xlsx            # 대리인 이메일 주소 목록
├── output-msg/                  # 생성된 .msg 파일들 (자동 생성)
├── AGENTS.md                    # 개발 가이드라인
├── helpme.txt                   # 프로젝트 배경 및 요구사항
└── README.MD                    # 이 문서
```

## 💡 주요 특징

### 1. 제목 자동 생성
Word의 메일 머지는 제목에 병합 필드를 사용할 수 없지만, 이 스크립트는 제목도 자동으로 생성합니다.

기본 제목 템플릿:
```
(«관리번호»«국가코드») New trademark application(s) in «국가명칭»
```

### 2. 수신인/참조 자동 설정
Excel의 `수신`, `참조` 컬럼 데이터를 읽어 메일의 To, CC 필드에 자동 설정합니다.

### 3. 파일 자동 첨부
Excel의 `첨부파일` 컬럼에 파일 경로를 지정하면 메일에 자동으로 첨부됩니다.
- 단일/다중 파일 지원 (세미콜론으로 구분)
- 절대 경로 및 상대 경로 모두 지원
- 파일 미존재 시 경고 후 계속 진행

### 4. 안전한 파일명 생성
제목을 기반으로 파일 시스템에 안전한 파일명을 자동 생성합니다.
- 특수문자는 언더스코어(_)로 치환
- 파일명이 비어있을 경우 `message_01`, `message_02` 등으로 대체

### 5. HTML 엔티티 처리
데이터의 특수문자(&, <, >)를 자동으로 이스케이프하여 HTML 본문이 깨지지 않도록 합니다.

### 6. 병합 필드 기호 자동 정규화
`«»`와 `<<>>` 두 가지 기호를 모두 지원합니다. 내부적으로 자동 변환되므로 어떤 기호를 사용해도 동일하게 작동합니다.

## 🔨 EXE 파일 생성하기 (배포용)

Python이 설치되지 않은 팀원들에게 배포할 때는 EXE 파일을 만들 수 있습니다.

### 1. PyInstaller 설치

```bash
pip install pyinstaller
```

### 2. GUI EXE 빌드 (권장)

```bash
pyinstaller --onefile --windowed mail_merge_gui.py
```

- `--onefile`: 단일 EXE 파일로 생성
- `--windowed`: 콘솔 창 숨김 (GUI만 표시)
- 생성 위치: `dist/mail_merge_gui.exe`

### 3. CLI EXE 빌드 (선택사항)

```bash
pyinstaller --onefile --console generate_mail_merge.py
```

- `--console`: 콘솔 창 표시 (출력 메시지 확인 가능)
- 생성 위치: `dist/generate_mail_merge.exe`

### 4. 배포

`dist` 폴더의 EXE 파일과 함께 다음 파일들을 배포하세요:
- `Filing.docx`, `Search.docx` (템플릿)
- `Filing_Merge.xlsx`, `Search_Merge.xlsx` (데이터)
- `메일링 리스트.xlsx` (대리인 이메일)
- `README.MD` (사용 설명서)

**주의**: EXE 파일은 약 20-30MB 크기입니다 (Python 런타임 포함).

## ⚠️ 주의사항

1. **Windows + Outlook 필수**: 이 도구는 Windows에서 Outlook이 설치되어 있어야 작동합니다.
2. **파일 인코딩**: 한글이 포함된 경우 모든 파일을 UTF-8로 저장하세요.
3. **병합 필드 일치**: Excel의 컬럼명과 Word 템플릿의 병합 필드명이 정확히 일치해야 합니다.
4. **데이터 검증**: 생성 후 최소 1-2개의 샘플 .msg 파일을 Outlook에서 열어 확인하세요.
5. **발송인 설정**: .msg 파일을 열면 현재 Outlook에 로그인된 계정이 자동으로 발송인으로 설정됩니다.

## 🐛 문제 해결

### `docx2msg를 찾을 수 없습니다` 오류
```bash
pip install docx2msg
```

### 생성된 MSG 파일이 열리지 않음
- Outlook이 제대로 설치되어 있는지 확인
- Windows 환경인지 확인
- pywin32가 올바르게 설치되었는지 확인

### 병합 필드가 치환되지 않음
- Excel 컬럼명과 Word 병합 필드명이 일치하는지 확인
- **중요**: `«»` 기호를 사용했는지 확인! `<<>>` (부등호)가 아닙니다!
  - ✅ 올바른 예: `«관리번호»`
  - ❌ 잘못된 예: `<<관리번호>>`
- 병합 필드가 `«필드명»` 형식으로 되어 있는지 확인 (Word의 병합 필드 삽입 기능 사용 권장)

### 한글이 깨짐
- 모든 파일(Python, Excel, Word)을 UTF-8 인코딩으로 저장

## 📧 워크플로우

1. **데이터 준비**: Excel에 발송할 메일 목록 작성
2. **스크립트 실행**: 메일 초안 자동 생성
3. **검토**: 생성된 .msg 파일들을 Outlook에서 열어 확인
4. **결재**: 필요시 .msg 파일을 결재 시스템에 첨부
5. **발송**: Outlook에서 메일 확인 후 발송

## 🔄 버전 관리

- 커밋 메시지는 명령형으로 작성 (예: `Add Outlook save helper`, `Fix relative path handling`)
- 한 번에 하나의 논리적 변경사항만 커밋
- 템플릿 변경 시 최소 1개 샘플로 테스트 확인

## 📚 참고

- [docx2msg 문서](https://github.com/Nirma/docx2msg)
- [python-docx 문서](https://python-docx.readthedocs.io/)
- [openpyxl 문서](https://openpyxl.readthedocs.io/)
