# 상품 단위 소싱 후보 발굴 패치

기존 버전은 `브랜드&키워드 300개.xlsx`의 브랜드/키워드를 기준으로 랭킹을 만들었습니다.
이번 패치는 브랜드 단위가 아니라 **브랜드 안의 실제 제품/모델/품목 키워드**를 발굴해 랭킹화합니다.

예:
- 나이키 → 나이키 에어맥스, 나이키 에어포스, 나이키 운동화
- 아식스 → 아식스 젤 카야노, 아식스 젤 1130
- WPC → WPC 양산, WPC 우양산, WPC 산리오
- 롱샴 → 롱샴 르 플리아쥬, 롱샴 파우치

## 핵심 옵션

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" \
  --top 100 \
  --discover-products \
  --products-only \
  --product-seed-limit 60 \
  --product-keywords-per-brand 8 \
  --min-product-volume 1000 \
  --max-product-candidates 200
```

## 테스트용

```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" \
  --limit 20 \
  --top 30 \
  --discover-products \
  --products-only \
  --product-seed-limit 10 \
  --product-keywords-per-brand 5 \
  --min-product-volume 500 \
  --max-product-candidates 50
```

## 결과물

- `reports/sourcing_rank_YYYYMMDD.xlsx`
  - `상품키워드_TOP`: 제품/모델 단위 추천 리스트
  - `전체랭킹`: 전체 점수표
  - `요약`: 후보 개수/점수 요약
- `reports/product_keywords_YYYYMMDD.csv`
  - 상품 단위 발굴 후보 원본 CSV

## 필수 Secrets

상품 단위 발굴은 검색광고 연관키워드/검색량을 써야 하므로 아래 3개가 필요합니다.

- `NAVER_SEARCHAD_API_KEY`
- `NAVER_SEARCHAD_SECRET_KEY`
- `NAVER_SEARCHAD_CUSTOMER_ID`

쇼핑 상품수/가격 조회에는 기존처럼 아래 2개가 필요합니다.

- `NAVER_CLIENT_ID`
- `NAVER_CLIENT_SECRET`

## 운영 팁

처음에는 `--product-seed-limit 20`, `--max-product-candidates 50` 정도로 작게 돌려서 결과를 확인하세요.
정상적으로 상품명/모델명 단위가 나오면 운영값을 60~100개 시드, 200~300개 후보로 늘리면 됩니다.
