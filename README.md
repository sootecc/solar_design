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

## 원본 VBA → Python 매핑

| VBA 모듈 | Python |
|---------|--------|
| Module1.bas (자재산출) | core/calculator.py |
| Module4.bas (수익분석) | core/revenue.py |
| Module6.bas (단선결선도) | core/electrical.py |
| Module11.bas (배치도) | core/layout.py |
| Module2.bas (지역검색) | data/irradiance.py |
