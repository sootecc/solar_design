# 태양광 가설계 자동화 시스템

Excel VBA 태양광현장분석프로그램 Ver15를 Python + Streamlit으로 이식한 웹 기반 가설계 툴.

## 기능

| 탭 | 기능 |
|----|------|
| 배열 N | 배열별 모듈/구조 설정 (최대 30배열) |
| 계산결과 | 구조물 치수·기둥·면적·발전량 요약 |
| 자재표 | 모듈·기초·기둥·보 등 자재 수량 산출 |
| 전기설계 | 스트링 설계 + 단선결선도 이미지 |
| 배치도 | 카카오 위성지도 + 배열 평면 오버레이 |
| 지역일사량 | 74개 지역 일조시간/일사량 DB (기상청 30년) |
| 수익분석 | 20/30년 IRR·투자회수기간·현금흐름 |

## 설치 및 실행

```bash
# 의존성 설치 및 실행 (Windows)
run.bat

# 또는 직접 실행
pip install -r requirements.txt
python -m streamlit run app.py
```

접속: http://localhost:8501

## 카카오 API 키 설정 (배치도 주소검색 필요)

1. [카카오 개발자 콘솔](https://developers.kakao.com) → 앱 생성 → REST API 키 복사
2. `.streamlit/secrets.toml` 파일 생성:

```toml
kakao_api_key = "여기에_REST_API_키_입력"
```

> ⚠️ `secrets.toml`은 `.gitignore`에 포함되어 있어 커밋되지 않습니다.

## 파일 구조

```
solar_design/
├── app.py                  # Streamlit 메인 앱
├── run.bat                 # 실행 스크립트
├── requirements.txt
├── core/
│   ├── models.py           # 데이터 모델 (ArrayConfig, ModuleSpec 등)
│   ├── calculator.py       # 구조·자재 계산 (SolarCalculator)
│   ├── electrical.py       # 전기설계 (ElectricalDesigner, 단선결선도)
│   ├── revenue.py          # 수익분석 (RevenueCalculator, IRR)
│   └── layout.py           # 배치도 (KakaoMap, LayoutDrawer)
├── data/
│   ├── modules.py          # 모듈 DB 41종
│   ├── inverters.py        # 인버터 DB 28종
│   └── irradiance.py       # 지역 일사량 DB 74지역
└── utils/
    └── output.py           # Excel 내보내기
```

## 다음 작업 (TODO)

### 1순위 - 발전량 계산 정확도 개선
현재 발전량은 `용량 × 일조시간 × 365 × PR` 단순 공식으로 PR이 모든 손실을 뭉뚱그림.
아래 항목이 **발전량에 미반영** 상태:

- [ ] **경사각(설치각도) → 경사면 일사량 보정**
  - 수평 일사량 → 경사면 일사량 변환 (Liu-Jordan 모델 또는 단순 cos 보정)
  - VBA 원본: Module1.bas / Module11.bas 참조
- [ ] **방위각 보정** — 정남(180°) 기준 편차에 따른 발전량 손실
- [ ] **인삼밭형 보정** — 양면 수광 or 반사광 추가 계수
- [ ] **배열간 음영 손실** — 앞 배열이 뒷 배열에 드리우는 그림자 시간

### 2순위 - 배치도 표시 개선
- [ ] 위성지도 위 배열 사각형이 제대로 표시되지 않는 문제 수정
  - 스케일(m/px) 검증 및 보정
  - 배열 최소 표시 크기 보장
  - 레이블 가독성 개선

### 3순위 - PDF 출력
- [ ] 계산결과 + 자재표 + 배치도 + 단선결선도 통합 보고서

### 4순위 - SMP/REC 자동 업데이트
- [ ] 한전/REC거래시장 최신 단가 자동 반영

---

## 원본 VBA → Python 매핑

| VBA 모듈 | Python |
|---------|--------|
| Module1.bas (자재산출) | core/calculator.py |
| Module4.bas (수익분석) | core/revenue.py |
| Module6.bas (단선결선도) | core/electrical.py |
| Module11.bas (배치도) | core/layout.py |
| Module2.bas (지역검색) | data/irradiance.py |
