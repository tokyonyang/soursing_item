# Telegram 요약 이미지 + Streamlit 링크 업그레이드 패치

이 패치는 GitHub Actions 실행 후 텔레그램으로 다음을 자동 전송합니다.

1. 이번 주 소싱 TOP 10 요약 이미지
2. Gemini 또는 기본 텍스트 브리핑
3. Streamlit 대시보드 링크

## 적용 파일

GitHub 저장소 루트에 아래 파일을 덮어쓰기/추가하세요.

```text
notify_report.py
requirements.txt
.github/workflows/scheduled-sourcing-rank.yml
```

Streamlit 대시보드를 아직 안 넣었다면 아래 파일도 함께 추가하세요.

```text
app.py
update_dashboard_data.py
dashboard_data/.gitkeep
.streamlit/config.toml
```

## GitHub Secrets

텔레그램 알림에 필요한 값입니다.

```text
TELEGRAM_BOT_TOKEN
TELEGRAM_CHAT_ID
```

기존 네이버 API 값도 그대로 필요합니다.

```text
NAVER_CLIENT_ID
NAVER_CLIENT_SECRET
NAVER_SEARCHAD_API_KEY
NAVER_SEARCHAD_SECRET_KEY
NAVER_SEARCHAD_CUSTOMER_ID
```

Gemini 요약은 선택입니다.

```text
GEMINI_API_KEY
```

## GitHub Variables

Streamlit 대시보드 URL은 Secret이 아니라 Variable로 넣는 것을 추천합니다.

```text
Settings
→ Secrets and variables
→ Actions
→ Variables
→ New repository variable
```

이름:

```text
DASHBOARD_URL
```

값:

```text
https://본인-streamlit-app-url.streamlit.app
```

선택 변수:

```text
GEMINI_MODEL
```

예:

```text
gemini-3.5-flash
```

## 작동 방식

```text
GitHub Actions 실행
→ sourcing_ranker.py가 상품 키워드 랭킹 생성
→ update_dashboard_data.py가 latest_sourcing_rank.xlsx 갱신
→ notify_report.py가 TOP 10 요약 이미지 생성
→ 텔레그램 sendPhoto + sendMessage 발송
→ Artifacts에 xlsx/csv/png 보관
```

## 테스트 방법

GitHub에서:

```text
Actions
→ scheduled-sourcing-rank
→ Run workflow
```

성공하면 텔레그램으로 다음 2개가 옵니다.

```text
1. 소싱 TOP 10 이미지
2. 소싱 브리핑 텍스트 + 대시보드 링크
```

## 한글 이미지가 깨질 때

워크플로에 아래 단계가 포함되어 있어야 합니다.

```yaml
- name: Install Korean fonts for Telegram image
  run: |
    sudo apt-get update
    sudo apt-get install -y fonts-noto-cjk
```

이 단계가 없으면 이미지 안의 한글이 네모로 보일 수 있습니다.

## 텔레그램 메시지만 받고 이미지는 끄고 싶을 때

워크플로의 Notify report env에서:

```yaml
SEND_TELEGRAM_IMAGE: "false"
```

로 바꾸거나 실행 명령에 `--no-image`를 붙이면 됩니다.

```yaml
run: python notify_report.py --top 15 --no-image
```
