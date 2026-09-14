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
- 새 버튼 또는 메뉴명은 리본, 통합실행 메뉴, 사용자 메뉴얼의 표현을 일치시킨다.
- API 키가 필요한 기능은 버전별 릴리스 산출물에 키가 남지 않는지 manifest의 `versionedArtifactsSanitized` 값을 확인한다.
- 커밋 전 `git diff --check`와 관련 테스트 스크립트를 실행한다.
