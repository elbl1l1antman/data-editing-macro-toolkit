# 개발 지침

## 한글 문자열 인코딩 검증

VBA, PowerShell, XML, 테스트 래퍼 코드에 한글 문자열을 추가하거나 수정할 때는 다음 절차를 반드시 수행한다.

- VBA로 직접 주입되는 테스트 문자열은 가능한 한 `ChrW(&H....)` 조합으로 작성한다.
- 배포 전 `publish_release.ps1`를 실행해 CP949 가져오기 경로에서 컴파일 가능한지 확인한다.
- 이전에 확인된 깨짐 패턴인 `怨듭떇`, `寃`, `吏`, `�` 같은 문자열을 최신 소스, 테스트, 릴리스 `MacroToolkit.bas`에서 검색한다.
- 사용자에게 표시되는 한글이 포함된 리본 XML, UserForm 생성 코드, 테스트용 VBA 래퍼는 소스 파일뿐 아니라 생성된 `release/v*/MacroToolkit.bas`와 `release/latest/MacroToolkit.bas`에서도 확인한다.
- 한글 기대값을 테스트해야 할 때는 테스트 자체가 인코딩 오류를 만들지 않도록 한글 직접 리터럴 대신 `ChrW` 헬퍼 함수를 우선 사용한다.

## 릴리스 전 확인

- 기능 변경 시 `TOOLKIT_VERSION`, `README.md`, `docs/USER_MANUAL.md`를 함께 갱신한다.
- 현재 버전의 `docs/releases/v버전.md`를 작성하며 `주요 변경`, `사용 방법`, `수정된 문제`, `주의사항`을 반드시 포함한다.
- 패치노트는 내부 함수명 나열보다 사용자가 무엇을 할 수 있고 무엇을 주의해야 하는지 중심으로 작성한다.
- 배포 스크립트가 패치노트를 `release/v버전/PATCH_NOTES.md`와 `release/latest/PATCH_NOTES.md`에 복사했는지 확인한다.
- 새 버튼 또는 메뉴명은 리본, 통합실행 메뉴, 사용자 메뉴얼의 표현을 일치시킨다.
- API 키가 필요한 기능은 최신본과 버전별 릴리스 모두 키가 남지 않는지 manifest의 `latestMayContainLocalCredentials`, `versionedArtifactsSanitized` 값을 확인한다.
- XLSM과 XLAM은 동일하게 조립된 VBA 모듈과 UserForm에서 생성하고, 퍼블리시 중 `Toolkit_CompileSmokeTest`를 통과해야 한다.
- 커밋 전 `git diff --check`와 관련 테스트 스크립트를 실행한다.

## 개발 문서 작성 원칙

- 기능 개발 전 문서는 기본적으로 하나의 통합 개발 명세서로 작성한다.
- 통합 명세서 안에 요구사항, UX 흐름, 구현 파일과 단계, 테스트 및 배포 기준을 함께 기록한다.
- 사용자가 별도 문서를 요구하거나 독립 검토가 필요한 고위험 변경이 아니면 설계서와 구현계획서를 중복 생성하지 않는다.
