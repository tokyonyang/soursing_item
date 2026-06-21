# GitHub Actions 자동 스케줄링 설정 가이드

이 패키지는 네이버 해외직구/구매대행 소싱 후보를 자동으로 랭킹화하고,
GitHub Actions에서 주 1회/월 1회 실행한 뒤 결과를 알림으로 보내는 구성입니다.

## 1. GitHub 저장소 만들기

1. GitHub에서 새 저장소를 만듭니다.
2. 이 폴더의 모든 파일을 저장소에 업로드합니다.
3. `브랜드&키워드 300개.xlsx` 파일도 저장소 루트에 같이 올립니다.
   - 파일명이 다르면 `.github/workflows/scheduled-sourcing-rank.yml`의 `--input` 값을 바꾸세요.

## 2. 필수 Secrets 등록

GitHub 저장소 > Settings > Secrets and variables > Actions > New repository secret

필수:

- `NAVER_CLIENT_ID`
- `NAVER_CLIENT_SECRET`

검색량 자동 갱신까지 하려면 추가:

- `NAVER_SEARCHAD_API_KEY`
- `NAVER_SEARCHAD_SECRET_KEY`
- `NAVER_SEARCHAD_CUSTOMER_ID`

Gemini 요약을 쓰려면 추가:

- `GEMINI_API_KEY`

## 3. 알림 채널 Secrets

### 텔레그램

- `TELEGRAM_BOT_TOKEN`
- `TELEGRAM_CHAT_ID`

### Slack / Discord / Make / Zapier Webhook

- `REPORT_WEBHOOK_URL`

### 이메일 SMTP

- `SMTP_HOST`
- `SMTP_PORT`
- `SMTP_USER`
- `SMTP_PASSWORD`
- `REPORT_FROM_EMAIL`
- `REPORT_TO_EMAIL`

Gmail을 쓰는 경우 일반 비밀번호가 아니라 앱 비밀번호를 사용하세요.

## 4. 실행 주기 변경

`.github/workflows/scheduled-sourcing-rank.yml`에서 아래 부분을 바꾸면 됩니다.

```yaml
on:
  schedule:
    - cron: "0 9 * * 1"
      timezone: "Asia/Seoul"
    - cron: "0 9 1 * *"
      timezone: "Asia/Seoul"
```

현재 설정:

- 매주 월요일 오전 9시 KST
- 매월 1일 오전 9시 KST

## 5. 수동 실행

GitHub 저장소 > Actions > `scheduled-sourcing-rank` > Run workflow

## 6. 결과 확인

- Actions 실행 화면의 Summary에서 TOP 요약 확인
- Artifacts에서 `sourcing-rank-report` 다운로드
- 텔레그램/웹훅/이메일 설정 시 자동 알림 수신

## 7. 주의

- API 키는 절대 코드나 엑셀에 직접 적지 마세요.
- 공개 저장소에 올릴 경우, 입력 엑셀 안에 민감한 소싱처 URL/원가가 들어있지 않은지 확인하세요.
- 자동 점수는 후보 선별용입니다. 등록 전에는 상표권, KC/전파법, 식품·화장품·의약외품 규정, 정품 증빙을 별도로 확인하세요.
