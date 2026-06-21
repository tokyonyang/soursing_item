# 신규 키워드 자동 발굴 패치

## 핵심 변경

이번 패치는 기존 `브랜드&키워드 300개.xlsx` 안의 후보만 재정렬하던 구조에서,
네이버 검색광고 키워드도구의 연관키워드를 활용해 신규 후보를 자동 추가하는 구조로 확장합니다.

## 추가 기능

- `--discover-related`: 신규 연관키워드 자동 발굴
- `--discovery-seed-limit`: 신규 발굴에 사용할 기존 후보 개수 제한
- `--related-per-brand`: 시드 후보당 추가할 연관키워드 수
- `--min-discovered-volume`: 신규 후보 최소 월간검색수
- `--max-discovered`: 한 번에 추가할 신규 후보 최대 수
- `--include-risky-discovery`: 리스크 키워드까지 포함

## 추천 테스트 명령

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --limit 20 --top 20 --discover-related --discovery-seed-limit 5 --related-per-brand 2 --max-discovered 10
```

## 추천 운영 명령

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --top 100 --discover-related --discovery-seed-limit 40 --related-per-brand 3 --min-discovered-volume 3000 --max-discovered 80
```

## 결과 확인

엑셀 리포트에 아래 시트가 추가됩니다.

- `신규발굴_TOP`

CSV로도 별도 저장됩니다.

- `reports/discovered_keywords_YYYYMMDD.csv`

## 주의

신규 키워드 자동 발굴은 검색광고 API 키가 있어야 작동합니다.
검색광고 API 키가 없으면 기존 300개 후보 재점수화만 수행됩니다.
