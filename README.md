# 급여대장 자동 변환 · 위하고 급여입력 마감 자동화

업체마다 양식이 다른 원본 급여대장(엑셀)을 위하고(Wehago T) 엑셀 불러오기 양식으로 바꾸는 PySide6 앱과,
변환 결과를 위하고에 올린 뒤 원본과 사원별 실지급액이 0원 차이로 맞을 때만 급여를 마감하는 Playwright 스크립트입니다.
구조와 위하고 화면 관찰 기록은 [CLAUDE.md](CLAUDE.md)를 보세요.

## 준비

- Windows, Python 3.12, Chrome(또는 Edge), 위하고 인쇄 모듈(WehagoPrint, 영수증 PDF 저장용)
- `pip install -r requirements.txt`, `python -m playwright install chromium`, 영수증 PDF 저장에는 `pywinauto`
- 비공개 자료는 저장소에 없습니다: `data/`(사업장 등록 DB, 설정·위하고 로그인, 학습 규칙), 업체 목록 엑셀,
  원본 급여대장 폴더, `tests/fixtures/`(실제 급여대장 샘플), `assets/fonts/`(NanumGothic).
- 코드 주석과 테스트의 `업체001`, `사원001`, `sample1`, `100xxxxxxx` 등은 실제 이름·번호를 바꾼 자리표시자입니다.
- 경로 상수(`run_one_company.py`의 `FOLDER`/`LIST`, `core/settings.py` 기본 폴더 등)는 환경에 맞게 고쳐서 씁니다.

## 실행

```
python main.py                                                    # 변환 앱
python run_one_company.py "<원본 파일명 일부>"                      # 위하고 상태 읽기만
python run_one_company.py "<원본 파일명 일부>" --apply --close      # 복사·임시등록·업로드·대조·마감
python run_midway_receipt.py "<업체 목록 업체명>"                  # 중도퇴사자 원천징수영수증 PDF
python -m pytest tests -q
```
