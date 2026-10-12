# Windows 배포 안내

이 앱은 PyInstaller `onedir` 배포 방식이다. 배포 대상 PC에는 Python을 설치할 필요가 없다.

> **2026-09-09 갱신**: 앱 범위가 처음(원본/위하고 파일 2개 선택 → 1회 변환)보다
> 크게 늘어서, 지금은 사업장 등록 DB(SQLite, `data/registry.db` — exe 폴더
> 안에 실행 중 자동 생성됨) + 배치 큐 + 위하고 양식 폴더 일괄 동기화가
> 포함된다. 자세한 기능 변천사는 `CLAUDE.md` 참고.
>
> **알려진 제약 — 위하고 폴더 경로가 하드코딩돼 있음**: "위하고 양식 폴더
> 업데이트" 버튼이 스캔하는 기본 폴더는 `core/wehago_sync.py`의
> `DEFAULT_FOLDER = Path(r"C:\Users\<사용자>\Desktop\작업파일\Wehago")`로
> 이 PC/이 계정 기준 절대경로다. 이 사용자 본인의 PC에서만 쓰는 걸 전제로
> 지금까지 문제가 안 됐지만, 다른 PC나 다른 Windows 계정에 배포하면 이
> 버튼이 "폴더 없음" 오류로 바로 막힌다 — 그런 경우엔 폴더 선택 다이얼로그로
> 바꾸는 작업이 먼저 필요하다.

## 배포본 만들기

1. **64비트 Windows** PC에서 이 프로젝트 폴더를 연다.
2. [Python](https://www.python.org/downloads/windows/) 3.10 이상(64비트)을 설치한다. 설치 화면에서 **pip**와 **Add python.exe to PATH**를 선택한다.
3. `build_windows.bat`을 더블클릭한다.
4. 성공하면 `dist\PayrollAutomation_Maumson&Care24-windows.zip`이 생성되고, 같은 ZIP 파일이 바탕화면에도 자동 복사된다.

빌드 스크립트는 가상환경 생성, 의존성 설치, PyInstaller 빌드, ZIP 생성을 순서대로 수행한다. 의존성 설치 시 인터넷 연결이 필요하다.

## 배포 및 실행

1. 생성된 `PayrollAutomation_Maumson&Care24-windows.zip` 하나만 전달한다.
2. 받는 사람은 ZIP 파일을 로컬 폴더에 **모두 압축 해제**한다.
3. `PayrollAutomation_Maumson&Care24\PayrollAutomation_Maumson&Care24.exe`를 실행한다.

`PayrollAutomation_Maumson&Care24.exe` 파일만 따로 옮기면 실행되지 않는다. 반드시 압축을 푼 `PayrollAutomation_Maumson&Care24` 폴더 전체를 함께 유지해야 한다.

## 참고 사항

- PyInstaller는 운영체제 간 교차 컴파일을 지원하지 않으므로, Windows 배포본은 Windows에서 빌드한다.
- 처음 실행할 때 Windows SmartScreen 경고가 나타날 수 있다. 코드 서명 인증서가 없는 일반 배포본의 정상적인 동작이다.
- 현재 앱의 입력 형식은 `.xlsx`와 `.xlsm`이다. `.hwp` 입력은 아직 처리 흐름에 연결되지 않았다.
