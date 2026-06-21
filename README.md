# 네이버 해외직구 소싱 후보 자동 랭킹

## 무엇을 자동화하나
입력 엑셀의 브랜드/키워드 후보를 기준으로 매주 또는 매월 다음 값을 조회합니다.

- 기존 300개 브랜드/키워드 후보 재점수화
- 네이버 검색광고 키워드도구 연관키워드 기반 신규 후보 자동 발굴
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

신규 연관키워드 발굴까지 포함해 실행:

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --top 100 --discover-related --discovery-seed-limit 40 --related-per-brand 3 --min-discovered-volume 3000 --max-discovered 80
```

테스트로 10개만 조회:

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --limit 10 --discover-related --discovery-seed-limit 5 --related-per-brand 2 --max-discovered 10
```

### 신규 키워드 발굴 옵션

| 옵션 | 기본값 | 설명 |
|---|---:|---|
| `--discover-related` | OFF | 검색광고 연관키워드로 신규 후보를 추가 |
| `--discovery-seed-limit` | 40 | 신규 발굴에 사용할 상위 기존 후보 수 |
| `--related-per-brand` | 5 | 브랜드/시드별 추가할 연관키워드 개수 |
| `--min-discovered-volume` | 3000 | 신규 후보 최소 월간검색수 |
| `--max-discovered` | 100 | 한 번에 추가할 신규 후보 최대 개수 |
| `--include-risky-discovery` | OFF | 식품/의약/성인 등 리스크 키워드도 포함 |

## 출력

`reports/sourcing_rank_YYYYMMDD.xlsx`

시트 구성:

- `요약`
- `TOP_100` 또는 실행 시 지정한 TOP 시트
- `신규발굴_TOP`
- `전체랭킹`

신규 발굴 후보가 있으면 별도 CSV도 저장됩니다.

```text
reports/discovered_keywords_YYYYMMDD.csv
```

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


## 운영 팁

처음에는 아래처럼 작게 돌려서 정상 여부를 확인하세요.

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --limit 20 --top 20 --discover-related --discovery-seed-limit 5 --related-per-brand 2 --max-discovered 10
```

정상 작동하면 GitHub Actions에서는 기본값처럼 `discovery-seed-limit 40`, `max-discovered 80` 정도로 운영하는 것을 권장합니다. 신규 후보를 너무 많이 늘리면 네이버 API 호출 수와 실행 시간이 급격히 늘어납니다.
