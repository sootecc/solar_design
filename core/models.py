"""
태양광 가설계 데이터 모델
VBA 원본: Module1.bas Public 변수 및 자료입력 시트 구조 기반
"""
from dataclasses import dataclass, field
from typing import Optional

# 설치형태 설명 (VBA 원본 분기 로직 기반)
INSTALLATION_TYPES = {
    1:  "평지형 - 소형기초 (경사설치)",
    2:  "평지형 - 소형기초 (수평설치)",
    3:  "건물옥상형 - 소형기초 (경사설치)",
    4:  "건물옥상형 - 소형기초 (수평설치)",
    5:  "건물지붕형 - 기초없음 (경사형)",
    6:  "건물지붕형 - 기초없음 (평지붕)",
    7:  "수면부유식 - 소형기초",
    8:  "지상형 - 대형기초 (독립기둥)",
    9:  "지상형 - 대형기초 (공유기둥)",
    10: "지상형 - 대형기초 (수직설치)",
    11: "인삼밭형 - 소형기초 A",
    12: "인삼밭형 - 소형기초 B",
    13: "인삼밭형 - 소형기초 C",
    14: "지상형 - 대형기초 (특수형)",
}

# 설치 방향
DIRECTIONS = ["남향", "남동향", "동향", "남서향", "서향", "북향", "북동향", "북서향"]


@dataclass
class ModuleSpec:
    """태양광 모듈 사양"""
    index: int = 0
    품명: str = ""
    용량_W: float = 640
    가로_mm: int = 1134
    세로_mm: int = 2465
    두께_mm: int = 30
    Voc_V: float = 56.61
    Vmax_V: float = 46.79
    무게_kg: float = 34.7
    효율_pct: float = 22.9
    연간효율감소_pct: float = 0.4


@dataclass
class InverterSpec:
    """인버터 사양"""
    index: int = 0
    품명: str = ""
    용량_kW: float = 3.7
    단상삼상: str = "단상"
    DC입력전압: str = "70-450"
    MPPT범위: str = "185-400"
    효율_pct: float = 96.0
    무게_kg: float = 8.6
    MPPT수: int = 1
    MPPT1스트링수: int = 1
    MPPT2스트링수: Optional[int] = None


@dataclass
class ArrayConfig:
    """
    배열 구성 파라미터
    VBA 원본: 자료입력 시트 행별 입력값 (Cells(행, 3 + 배열위치*2))
    """
    # 배열 기본 구성
    배열의가로: int = 4          # 가로 방향 모듈 수
    배열의세로: int = 2          # 세로 방향 모듈 수
    앞쪽깔기삭제수량: int = 0     # 앞쪽 제거 모듈 수
    뒤쪽깔기삭제수량: int = 0     # 뒤쪽 제거 모듈 수
    배열중앙추가장수: int = 0     # 중앙 추가 모듈 수

    # 설치 조건
    설치각도: float = 15.0       # 경사각 (도)
    시작높이_m: float = 0.5      # 최하단 높이 (m)
    설치방향: str = "남향"       # 방위각 방향
    설치방향각도: float = 0.0    # 방위각 (도, 남향=0)
    설치형태: int = 3            # 설치형태 코드 (1~14)

    # 이격거리
    모듈판가로이격거리_mm: int = 20   # 모듈 가로 간격
    모듈판세로이격거리_mm: int = 20   # 모듈 세로 간격

    # 모듈 정보 (계산에 사용)
    모듈_index: int = 0

    @property
    def 실제모듈수량(self) -> int:
        """실제 설치 모듈판 수량 계산 (VBA: 자재산출메인 로직)"""
        return (self.배열의가로 * self.배열의세로
                - self.앞쪽깔기삭제수량
                - self.뒤쪽깔기삭제수량
                + self.배열중앙추가장수)


@dataclass
class ProjectInfo:
    """프로젝트 기본 정보"""
    현장명: str = ""
    주소: str = ""
    설계자: str = ""
    설계일자: str = ""
    총배열수: int = 1
    인버터_index: int = 0
    인버터수량: int = 1
    추가인버터_index: Optional[int] = None
    추가인버터수량: int = 0


@dataclass
class StructureDefaults:
    """
    구조물 기본 설정값
    VBA 원본: 도면제작설정값 시트
    """
    # 부속류 무게 (kW당)
    부속류_kW당무게_kg: float = 2.5

    # 소형기초 (설치형태 1,2,3,4,7,11,13)
    소형기초_가로_mm: int = 300
    소형기초_세로_mm: int = 300
    소형기초_높이_mm: int = 100
    소형기초_무게_kg: float = 21.6

    # 베이스판
    베이스판_가로_mm: int = 220
    베이스판_세로_mm: int = 220
    베이스판_두께_mm: int = 9
    베이스판_무게_kg: float = 3.42

    # 대형기초 (설치형태 8,9,10,14)
    대형기초_가로_mm: int = 500
    대형기초_세로_mm: int = 500
    대형기초_높이_mm: int = 500
    대형기초_무게_kg: float = 300.0

    # 기둥 (C형강)
    기둥_C형강_규격: str = "C형강:100×50×20×2.1T"
    기둥_단위무게_kg_m: float = 3.6

    # 퍼거라 보
    주보_규격: str = "각관:100×100×2.1T"
    주보_단위무게_kg_m: float = 6.9

    # 연결재
    연결재_규격: str = "C형강:75×45×2.1T"
    연결재_단위무게_kg_m: float = 2.9

    # 최소/최대 이격거리
    최소가로이격거리_mm: int = 200
    최대가로이격거리_mm: int = 1000
    최소세로이격거리_mm: int = 500
    최대세로이격거리_mm: int = 1200


@dataclass
class DesignResult:
    """가설계 계산 결과"""
    # 배열 정보
    배열번호: int = 1
    모듈수량: int = 0
    총용량_kW: float = 0.0
    구조물가로_m: float = 0.0
    구조물세로_m: float = 0.0
    구조물높이_m: float = 0.0
    설치면적_m2: float = 0.0

    # 기둥 정보
    가로기둥수: int = 0
    세로기둥수: int = 0
    총기둥수: int = 0

    # 자재 수량 목록 [(자재명, 규격, 수량, 단위, 단위무게, 총무게)]
    자재목록: list = field(default_factory=list)

    # 총 무게
    총무게_kg: float = 0.0

    # 전기 설계
    스트링직렬수: int = 0
    스트링병렬수: int = 0
    인버터수량: int = 1
