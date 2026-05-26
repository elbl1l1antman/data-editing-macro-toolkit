# 데이터 편집 매크로 툴킷

Excel 데이터 편집과 간단한 분석 작업을 위한 VBA 매크로 툴킷입니다.

## 프로젝트 구조

- `src/`: VBA 원본 소스입니다. 기능 수정은 여기서 먼저 합니다.
- `dev/`: 개발/검토용 `.xlsm` 통합문서입니다.
- `release/`: 버전별 배포 산출물입니다.
- `scripts/`: 빌드와 퍼블리시 자동화 스크립트입니다.

## 현재 버전

`0.9.2`

버전의 기준은 `src/MacroToolkit.bas`의 `TOOLKIT_VERSION`입니다.

## 배포 정책

앞으로 실제 설치/배포 파일은 `.xlam`을 기준으로 합니다.

- `.xlam`: Excel 추가 기능 설치용 파일입니다. 사용자는 이 파일을 설치합니다.
- `.xlsm`: 개발/검토용 통합문서입니다. 설명 시트와 소스 미리보기를 확인할 때 사용합니다.
- `.bas`: Git에서 관리하는 VBA 원본 소스입니다.

## 현재 기능

- 텍스트 형식 숫자를 실제 숫자값으로 변환
- 현재 시트의 실제 데이터 영역만 대상으로 텍스트 숫자 일괄 변환
- 혼재된 한국 전화번호 형식 정규화
- 전화번호 정리 시 원본 오른쪽에 복사본 열을 생성한 뒤 복사본만 정규화
- 전화번호/사업자등록번호의 하이픈 제거 복사본 열 생성
- 기업체명/사업체명 정리: 공백, 회사형태 표기, '의 ' 앞부분, 알파벳 한글 읽기 규칙을 선택 적용
- 구주소/도로명주소/우편번호 변환표를 이용한 도로명주소 변환
- 모든 워크시트의 화면 위치를 A1 기준 최상단으로 정렬
- 모든 워크시트의 선택 셀을 A1로 변경
- 숫자 변환과 전화번호 정리를 UserForm 기반 범위 선택 GUI로 실행
- 범위 선택 시 Excel 시트에서 Ctrl+Shift+방향키, Shift+방향키, 마우스 드래그로 선택 후 실행 가능
- 이상치 탐지 및 대체 결과 시트 생성
- 이상치 확인 실행 시 범위 검증 후 기초통계량표를 먼저 생성
- 표 전체 선택 시 숫자형 열 자동 탐지 후 선택 목록 제공
- Q1/Q3 기준값 열, 원본 검토복사 열, 조건부 서식 자동 생성
- SPSS 내보내기 통계표에서 선택한 가로/세로 배너 블록 삭제
- SPSS 목차 링크 기반 표/배너 탐지 후 원본 시트 복사본에서 삭제하고 목차 재생성
- SPSS 표/배너 재구성 목록에서 Shift+클릭 연속 선택 및 Ctrl+클릭 개별 선택 지원
- SPSS 표/배너 삭제 결과에서 표와 표 사이 공백 행을 1개로 자동 정리
- 목차가 없는 SPSS 집계표 시트에서 표 제목을 탐지해 목차 링크 시트 생성
- SPSS BOT/TOP/MID 계열 세로형 배너 삭제 기능 분리
- 0.9.0 리본 UI 정리: 중복 버튼 축소, 작업 중심 라벨, 기능 연상 아이콘/설명 개선
- 0.9.0 UserForm 개선: 버튼/목록 폭 확대, 긴 한글 버튼과 표/배너명이 잘리지 않도록 배치 조정

## 개발용 통합문서 확인

`dev/DataEditingMacroToolkit_dev.xlsm`을 열어 다음 시트를 확인합니다.

- `README`: 구현 위치와 설치 안내
- `매크로 목록`: 사용 가능한 매크로와 기능 설명
- `VBA 소스 미리보기`: `src/MacroToolkit.bas` 내용을 통합문서 안에서 확인

## 개발용 통합문서 빌드

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build_macro_workbook.ps1
```

## 릴리스 퍼블리시

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish_release.ps1
```

`TOOLKIT_VERSION`을 읽어 다음 파일을 생성합니다.

- `release/v0.9.2/DataEditingMacroToolkit_v0.9.2.xlam`
- `release/v0.9.2/DataEditingMacroToolkit_v0.9.2.xlsm`
- `release/v0.9.2/MacroToolkit.bas`
- `release/v0.9.2/manifest.json`
- `release/latest/DataEditingMacroToolkit_latest.xlam`
- `release/latest/DataEditingMacroToolkit_latest.xlsm`

## Excel 추가 기능 설치

Excel에서 다음 경로로 `.xlam` 파일을 설치합니다.

파일 > 옵션 > 추가 기능 > Excel 추가 기능 > 이동 > 찾아보기


설치 후 Excel 상단 리본의 `보기` 탭 옆에 `데이터 툴킷` 탭이 추가됩니다. 기능은 이 탭의 버튼으로 실행합니다.

- `값 정리`: 숫자 변환, 시트 숫자 변환, 전화번호 정리, 전화번호 숫자만, 기업명 정리, 주소 변환
- `사업자등록`: 인증키 설정, 번호 숫자만, 3열 병합, 상태조회
- `SPSS 표 정리`: 목차 생성, 표/배너 재구성, 세로 배너 삭제
- `이상치 탐지`: 범위 선택, Q1/Q3 검토열 생성, 설정
- `정보`: 모든 시트 A1 선택, 버전 정보
설치 대상 파일:

`release/latest/DataEditingMacroToolkit_latest.xlam`

## Excel VBA 접근 설정

자동으로 `.bas`를 통합문서에 삽입하려면 Excel에서 다음 옵션이 필요합니다.

파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정 > VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음

이 옵션이 꺼져 있으면 `.xlsm`에는 실행 가능한 VBA 모듈이 없을 수 있습니다. 이 경우 `src/MacroToolkit.bas`를 VBA 편집기에서 직접 가져오면 됩니다.



