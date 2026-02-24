# Excel Regression

회귀테스트 양식 처리를 위한 자동화된 Excel 파일 도구입니다.

## 사용 방법

### Step 1: GitHub에서 다운로드

1. 저장소에 접속 https://github.com/reomoon/excel_regression
2. "Code" 버튼 클릭
3. "Download ZIP" 선택
4. ZIP 파일을 원하는 위치에 압축 해제
5. Bug Tracking System 두레이에서 결함목록 다운로드

### 단계 2: 프로그램 실행

1. 압축 해제한 폴더 오픈
2. `start.bat` 파일 더블클릭

### 단계 3: Excel 파일 선택

1. 파일 선택 대화창이 나타나면 Excel 파일(.xlsx) 선택
2. 프로그램이 자동으로 파일 처리
3. 처리 결과가 `output` 폴더에 저장됩니다

### 단계 4: 결과 확인

1. `output` 폴더가 자동으로 열립니다
2. 처리된 파일 확인: `reg_output.xlsx`

## 기술 사항

### Python 경로 자동 감지

- 배치 파일이 Python 설치 위치를 자동으로 찾으므로 PATH 환경변수 설정이 필요하지 않습니다. Python 3를 검색하여 전체 실행 경로를 사용합니다.
- Python파일을 찾을 수 없을 경우 아래 문제 해결을 확인하여 해결 합니다.

### 인코딩 처리

- 배치 파일은 영문만 사용하여 GitHub에서 다운로드했을 때 인코딩 문제 방지
- UTF-8 인코딩 오류를 제거하기 위해 모든 한글 문자 제거
- 다양한 시스템에서 호환성 보장

### 필수 패키지

- openpyxl (첫 실행 시 자동 설치)

### 폴더 구조

```
project/
├── start.bat          (이 파일을 실행하세요)
├── run.py             (메인 스크립트)
├── requirements.txt   (패키지 의존성)
├── core/
│   └── regression_excel.py
├── output/            (결과 저장 폴더)
└── upload/            (입력 파일 저장)
```

## 시스템 요구사항

- Python 3.7 이상 (컴퓨터에 설치되어 있어야 함)
- Windows 운영체제
- 최소 50MB 디스크 용량

## 문제 해결

### "Python을 찾을 수 없음" 오류

다음과 같이 진행하세요:

1. Python을 https://www.python.org에서 설치 (3.9 이상)
2. 설치 중 "Add Python to PATH" 반드시 체크
3. 컴퓨터 재 시작
4. `start.bat` 다시 실행

### 파일 처리 실패

- Excel 파일이 .xlsx 형식인지 확인
- 다른 프로그램에서 파일이 열려있지 않은지 확인
- 프로그램 폴더에 쓰기 권한이 있는지 확인

### Output 폴더가 자동으로 열리지 않음

- `output` 폴더가 생성되었는지 확인
- 수동으로 프로그램 폴더 내 `output` 폴더를 열어 처리된 파일 확인
