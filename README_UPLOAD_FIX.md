# Streamlit 업로드 오류 수정 패치

## 수정 원인
Streamlit `file_uploader`로 받은 엑셀은 `uploaded.getvalue()`에서 bytes로 반환됩니다.
기존 `app.py`는 이 bytes를 `pd.ExcelFile()`에 그대로 넘겨서 Streamlit Cloud에서 아래 오류가 발생했습니다.

`TypeError: Expected file path name or file-like object`

## 적용 방법
1. 이 ZIP 안의 `app.py`를 GitHub 저장소 루트의 기존 `app.py`와 교체합니다.
2. Commit changes를 누릅니다.
3. Streamlit Cloud에서 Reboot app 또는 Rerun을 실행합니다.
4. 다시 엑셀을 업로드하거나, `dashboard_data/latest_sourcing_rank.xlsx` 자동 파일을 확인합니다.
