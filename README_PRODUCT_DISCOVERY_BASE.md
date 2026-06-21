# 네이버 해외직구 소싱 후보 자동 랭킹

## 무엇을 자동화하나
입력 엑셀의 브랜드/키워드 후보를 기준으로 매주 또는 매월 다음 값을 조회합니다.

- 네이버 검색광고 키워드도구 월간검색수
- 네이버 쇼핑검색 API 전체 상품수
- `exclude=cbshop` 기준 국내/일반 상품수
- 역산 해외직구/구매대행 상품수
- 상위 쇼핑 결과 가격 샘플
- 전회 실행 대비 검색량 상승률
- 최종 소싱점수와 추천등급

## 설치

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env
```

`.env` 파일에 네이버 API 키를 넣으세요.

## 실행

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --top 80
```

검색광고 API 키가 아직 없으면 입력 파일의 검색량으로 임시 실행할 수 있습니다.

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --top 80 --no-searchad
```

테스트로 10개만 조회:

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --limit 10
```

## 출력

`reports/sourcing_rank_YYYYMMDD.xlsx`

시트 구성:

- `요약`
- `TOP_80`
- `전체랭킹`

## 스케줄링 예시

### macOS/Linux cron: 매주 월요일 오전 9시

```cron
0 9 * * 1 cd /path/to/naver_sourcing_automation && /path/to/.venv/bin/python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --top 80 >> logs/run.log 2>&1
```

### Windows 작업 스케줄러

동작 프로그램:
```text
C:\path\to\.venv\Scripts\python.exe
```

인수:
```text
C:\path\to\naver_sourcing_automation\sourcing_ranker.py --input "C:\path\to\브랜드&키워드 300개.xlsx" --top 80
```

## 점수 기준

최종 점수는 아래 항목을 가중합합니다.

- 검색량
- 전회 대비 상승률
- 검색량 대비 해외직구 상품수 기회
- 경쟁도
- 객단가 적정성
- 시즌성
- 리스크 감점: 식품/화장품/전자/KC/정품증빙/사이즈반품 등

등록 전에는 반드시 상표권, 정품 구매처 증빙, KC/전파법, 식품·화장품·의약외품 수입 규정, 네이버 상품명 SEO를 별도로 확인하세요.
