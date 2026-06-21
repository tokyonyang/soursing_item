# Streamlit 소싱 대시보드 패치

이 패치는 기존 네이버 해외직구 소싱 자동화 결과를 웹 화면에서 볼 수 있게 해주는 Streamlit 대시보드입니다.

## 추가되는 파일

```text
app.py
update_dashboard_data.py
dashboard_data/.gitkeep
.streamlit/config.toml
requirements.txt
github_workflow_examples/scheduled-sourcing-rank-with-dashboard-commit.yml
DASHBOARD_README.md
```

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

실행 후 브라우저가 열리면, 직접 `sourcing_rank_YYYYMMDD.xlsx` 파일을 업로드해서 확인할 수 있습니다.

## GitHub + Streamlit Cloud 운영 구조

추천 구조는 아래와 같습니다.

```text
GitHub Actions
→ sourcing_ranker.py 실행
→ reports/sourcing_rank_YYYYMMDD.xlsx 생성
→ update_dashboard_data.py 실행
→ dashboard_data/latest_sourcing_rank.xlsx 갱신
→ GitHub repo에 자동 커밋
→ Streamlit Cloud가 최신 대시보드 표시
```

## 기존 워크플로에 추가할 단계

기존 `.github/workflows/scheduled-sourcing-rank.yml`에서 `permissions`를 아래처럼 바꿔주세요.

```yaml
permissions:
  contents: write
```

그리고 `Upload Excel report` 전에 아래 단계를 추가하세요.

```yaml
      - name: Prepare dashboard data
        if: success()
        run: python update_dashboard_data.py

      - name: Commit dashboard data
        if: success()
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "41898282+github-actions[bot]@users.noreply.github.com"
          git add dashboard_data/
          git commit -m "Update dashboard data" || echo "No dashboard data changes"
          git push
```

저장소가 Public이면 소싱 결과가 공개될 수 있으니, 가능하면 Private 저장소로 바꾸는 것을 권장합니다.

## Streamlit Cloud 배포

1. Streamlit Community Cloud 접속
2. GitHub 저장소 연결
3. Main file path에 `app.py` 입력
4. Deploy 클릭

## 화면 구성

- TOP 리스트
- 점수 TOP 차트
- 검색량 대비 경쟁 산점도
- 상위 후보 상세 검토
- 스마트스토어 상품명/태그 초안
- API 오류/긴 키워드 품질 경고
- CSV 다운로드
