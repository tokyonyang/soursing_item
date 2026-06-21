# Streamlit safe app patch

GitHub 저장소 루트의 기존 app.py를 이 파일로 교체하세요.

주요 수정:
- 업로드 엑셀 bytes를 BytesIO로 처리
- dashboard_data/latest_sourcing_rank.xlsx가 없어도 앱이 죽지 않음
- 오류 발생 시 Streamlit 화면에 traceback 표시
- 일부 Streamlit 버전에서 문제가 될 수 있는 차트/column_config 의존성 최소화
