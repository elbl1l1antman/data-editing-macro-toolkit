# 데이터 편집 매크로 툴킷

Excel 데이터 편집과 간단한 분석 작업을 위한 VBA 매크로 툴킷입니다.

## 프로젝트 구조

- `src/`: VBA 원본 소스입니다. 기능 수정은 여기서 먼저 합니다.
- `dev/`: 개발/검토용 `.xlsm` 통합문서입니다.
- `release/`: 버전별 배포 산출물입니다.
- `scripts/`: 빌드와 퍼블리시 자동화 스크립트입니다.

## 현재 버전

`0.6.1`

버전의 기준은 `src/MacroToolkit.bas`의 `TOOLKIT_VERSION`입니다.

## 배포 정책

앞으로 실제 설치/배포 파일은 `.xlam`을 기준으로 합니다.

- `.xlam`: Excel 추가 기능 설치용 파일입니다. 사용자는 이 파일을 설치합니다.
- `.xlsm`: 개발/검토용 통합문서입니다. 설명 시트와 소스 미리보기를 확인할 때 사용합니다.
- `.bas`: Git에서 관리하는 VBA 원본 소스입니다.

## 현재 기능

- 텍스트 형식 숫자를 실제 숫자값으로 변환
- 혼재된 한국 전화번호 형식 정규화
- 이상치 탐지 및 대체 결과 시트 생성`r`n- 이상치 탐지 실행 시 팝업으로 데이터 범위, 분석 변수, 기준, 대체 방식 지정`r`n- SPSS 내보내기 통계표에서 선택한 가로/세로 배너 블록 삭제

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

- `release/v0.6.1/DataEditingMacroToolkit_v0.6.1.xlam`
- `release/v0.6.1/DataEditingMacroToolkit_v0.6.1.xlsm`
- `release/v0.3.0/MacroToolkit.bas`
- `release/v0.3.0/manifest.json`
- `release/latest/DataEditingMacroToolkit_latest.xlam`
- `release/latest/DataEditingMacroToolkit_latest.xlsm`

## Excel 추가 기능 설치

Excel에서 다음 경로로 `.xlam` 파일을 설치합니다.

파일 > 옵션 > 추가 기능 > Excel 추가 기능 > 이동 > 찾아보기


설치 후 Excel 상단 리본의 `보기` 탭 옆에 `데이터 툴킷` 탭이 추가됩니다. 기능은 이 탭의 버튼으로 실행합니다.

- `숫자 변환`: 선택 영역, 현재 시트, 범위 선택
- `전화번호`: 선택 영역, 범위 선택
- `SPSS 표 정리`: 선택 배너 삭제`r`n- `이상치 탐지`: 설정, 실행
- `정보`: 버전 정보
설치 대상 파일:

`release/latest/DataEditingMacroToolkit_latest.xlam`

## Excel VBA 접근 설정

자동으로 `.bas`를 통합문서에 삽입하려면 Excel에서 다음 옵션이 필요합니다.

파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정 > VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음

이 옵션이 꺼져 있으면 `.xlsm`에는 실행 가능한 VBA 모듈이 없을 수 있습니다. 이 경우 `src/MacroToolkit.bas`를 VBA 편집기에서 직접 가져오면 됩니다.



