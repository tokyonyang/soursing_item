# sourcing_ranker.py 긴 키워드/414 오류 수정 패치

이 패치는 `sourcing_rank_20260621.xlsx` 검수 결과 발견된 문제를 고친 버전입니다.

## 수정 내용
1. `키워드&시리즈` 줄바꿈 보존
   - 기존 버전은 줄바꿈을 공백으로 바꿔서 전체 키워드 묶음이 API query로 들어갔습니다.
   - 그 결과 Shopping API 414 Request-URI Too Large, SearchAd 400 hintKeywords 오류가 발생했습니다.

2. 대표 키워드 1개만 API 조회
   - 줄바꿈/쉼표 기준 첫 번째 키워드만 사용합니다.
   - 너무 길면 브랜드명으로 자동 축소합니다.

3. 엑셀 헤더 줄바꿈 인식
   - `네이버\n상품수`, `네이버 해외\n상품수` 같은 헤더를 정상 매칭합니다.

4. API 실패 시 입력 엑셀의 상품수/해외상품수를 fallback으로 사용
   - API 오류가 나도 점수표가 0으로 무너지는 문제를 줄였습니다.

## 적용 방법
GitHub 저장소의 `sourcing_ranker.py` 파일을 이 패치 버전으로 교체한 뒤 커밋하세요.

테스트 실행:
```bash
python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --limit 20 --top 20 --no-searchad
```

GitHub Actions 테스트용:
```yaml
run: python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --limit 20 --top 20 --no-searchad
```

정상 확인 후:
```yaml
run: python sourcing_ranker.py --input "브랜드&키워드 300개.xlsx" --top 80
```
