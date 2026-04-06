"""
태양광 가설계 핵심 계산 엔진
VBA 원본: Module1.bas 자재산출메인(), 평면배치도그리기_자재산출신(), 측면도_자재산출신() 등
"""
import math
from .models import ArrayConfig, ModuleSpec, InverterSpec, StructureDefaults, DesignResult


class SolarCalculator:
    """태양광 가설계 계산기"""

    def __init__(self, defaults: StructureDefaults = None):
        self.defaults = defaults or StructureDefaults()

    # -----------------------------------------------------------------------
    # 배열 기하학 계산
    # -----------------------------------------------------------------------

    def calc_structure_width(self, config: ArrayConfig, module: ModuleSpec) -> float:
        """
        구조물 가로 길이 계산 (m)
        VBA: 구조물가로길이 = (모듈가로 × 가로수 + 이격거리 × (가로수-1)) / 1000
        """
        width_mm = (module.가로_mm * config.배열의가로
                    + config.모듈판가로이격거리_mm * (config.배열의가로 - 1))
        return width_mm / 1000.0

    def calc_module_projected_length(self, config: ArrayConfig, module: ModuleSpec) -> float:
        """
        모듈 세로 수평 투영 길이 계산 (mm)
        경사각을 반영한 수평 투영
        VBA: 모듈판세로 = 모듈판세로길이 × Cos(경사각)
        """
        angle_rad = math.radians(config.설치각도)
        return module.세로_mm * math.cos(angle_rad)

    def calc_structure_depth(self, config: ArrayConfig, module: ModuleSpec) -> float:
        """
        구조물 세로(앞뒤) 길이 계산 (m) - 수평 투영 기준
        VBA: 구조물세로길이 = 투영세로 × 세로수 + 이격거리 × (세로수-1)
        """
        projected_mm = self.calc_module_projected_length(config, module)
        depth_mm = (projected_mm * config.배열의세로
                    + config.모듈판세로이격거리_mm * (config.배열의세로 - 1))
        return depth_mm / 1000.0

    def calc_structure_height(self, config: ArrayConfig, module: ModuleSpec) -> float:
        """
        구조물 높이 (m) - 경사각과 시작높이 고려
        뒤쪽 최고 높이 기준
        """
        angle_rad = math.radians(config.설치각도)
        module_height_m = module.세로_mm * math.sin(angle_rad) / 1000.0
        return config.시작높이_m + module_height_m * config.배열의세로

    def calc_installation_area(self, config: ArrayConfig, module: ModuleSpec) -> float:
        """설치 면적 계산 (m²) - 구조물 투영면적"""
        w = self.calc_structure_width(config, module)
        d = self.calc_structure_depth(config, module)
        return w * d

    # -----------------------------------------------------------------------
    # 기둥 수량 계산 (평면배치도그리기_자재산출신 로직)
    # -----------------------------------------------------------------------

    def calc_column_count(self, config: ArrayConfig) -> tuple[int, int, int]:
        """
        기둥 수량 계산
        VBA: 평면배치도그리기_자재산출신() 로직
        Returns: (가로기둥수, 세로기둥수, 총기둥수)
        """
        # 기본: 가로는 배열 가로 + 1, 세로는 배열 세로 + 1
        # (각 모듈 사이와 양끝에 기둥)
        cols = config.배열의가로 + 1
        rows = config.배열의세로 + 1

        # 설치형태별 보정
        if config.설치형태 in (5, 6):
            # 지붕 고정형: 기둥 없음 (클램프 고정)
            return 0, 0, 0

        total = cols * rows
        return cols, rows, total

    # -----------------------------------------------------------------------
    # 전기 설계 계산
    # -----------------------------------------------------------------------

    def calc_string_design(
        self,
        module: ModuleSpec,
        inverter: InverterSpec,
        module_count: int
    ) -> dict:
        """
        스트링 설계 계산
        VBA: Module6.bas 스트링 설계 로직
        """
        # MPPT 전압 범위 파싱
        try:
            mppt_min, mppt_max = [float(v) for v in inverter.MPPT범위.split("-")]
        except Exception:
            mppt_min, mppt_max = 200.0, 800.0

        # 직렬 모듈 수 결정 (Vmax 기준으로 MPPT 범위 내)
        if module.Vmax_V <= 0:
            직렬수 = 1
        else:
            직렬수_최대 = int(mppt_max / module.Vmax_V)
            직렬수_최소 = max(1, int(math.ceil(mppt_min / module.Vmax_V)))
            직렬수 = min(직렬수_최대, max(직렬수_최소, 1))

        # 병렬 수 결정 (총 모듈수 기준)
        병렬수 = math.ceil(module_count / 직렬수) if 직렬수 > 0 else module_count

        # 인버터 용량과 배열 용량 비교
        배열용량_kW = round(module_count * module.용량_W / 1000, 2)
        인버터적합여부 = (inverter.용량_kW * 0.8 <= 배열용량_kW <= inverter.용량_kW * 1.2)

        return {
            "직렬모듈수": 직렬수,
            "병렬스트링수": 병렬수,
            "총모듈수": module_count,
            "배열용량_kW": 배열용량_kW,
            "인버터용량_kW": inverter.용량_kW,
            "인버터적합여부": 인버터적합여부,
            "스트링전압_V": round(직렬수 * module.Vmax_V, 1),
        }

    # -----------------------------------------------------------------------
    # 자재 산출 (자재산출메인 로직)
    # -----------------------------------------------------------------------

    def calc_materials(
        self,
        config: ArrayConfig,
        module: ModuleSpec,
        inverter: InverterSpec,
        inverter_count: int = 1,
        array_index: int = 1,
        total_arrays: int = 1,
        project_capacity_kW: float = None,
    ) -> DesignResult:
        """
        자재 산출 메인 계산
        VBA 원본: 자재산출메인() 함수
        """
        d = self.defaults
        result = DesignResult(배열번호=array_index)

        # 1. 모듈 수량
        result.모듈수량 = config.실제모듈수량
        result.총용량_kW = round(result.모듈수량 * module.용량_W / 1000.0, 2)

        # 2. 구조물 치수
        result.구조물가로_m = round(self.calc_structure_width(config, module), 2)
        result.구조물세로_m = round(self.calc_structure_depth(config, module), 2)
        result.구조물높이_m = round(self.calc_structure_height(config, module), 2)
        result.설치면적_m2 = round(self.calc_installation_area(config, module), 2)

        # 3. 기둥 수량
        result.가로기둥수, result.세로기둥수, result.총기둥수 = self.calc_column_count(config)

        # 4. 자재 목록 작성
        materials = []

        # [전체 현장 공통 자재 - 배열1에서만 출력]
        if array_index == 1:
            # 인버터
            materials.append({
                "자재명": "인버터",
                "규격": inverter.품명,
                "단위": "대",
                "수량": inverter_count,
                "단위무게_kg": inverter.무게_kg,
                "총무게_kg": round(inverter_count * inverter.무게_kg, 2),
            })

            # 분전함 등 기타 부속류
            cap_kW = project_capacity_kW or result.총용량_kW
            부속류무게 = round(cap_kW * d.부속류_kW당무게_kg, 2)
            materials.append({
                "자재명": "분전함등기타자재류",
                "규격": "볼트, 너트 등 기타부속류",
                "단위": "식",
                "수량": 1,
                "단위무게_kg": 부속류무게,
                "총무게_kg": 부속류무게,
            })

        # [배열별 자재]
        # 태양전지모듈
        모듈총무게 = round(result.모듈수량 * module.무게_kg, 2)
        materials.append({
            "자재명": "태양전지모듈",
            "규격": f"{module.품명} ({module.용량_W}W)",
            "단위": "장",
            "수량": result.모듈수량,
            "단위무게_kg": module.무게_kg,
            "총무게_kg": 모듈총무게,
        })

        # 설치형태별 기초 자재
        if config.설치형태 in (1, 2, 3, 4, 7, 11, 12, 13):
            # 소형 기초
            if result.총기둥수 > 0:
                materials.append({
                    "자재명": "소형보조기초",
                    "규격": f"{d.소형기초_가로_mm}×{d.소형기초_세로_mm}×{d.소형기초_높이_mm}mm 콘크리트",
                    "단위": "EA",
                    "수량": result.총기둥수,
                    "단위무게_kg": d.소형기초_무게_kg,
                    "총무게_kg": round(result.총기둥수 * d.소형기초_무게_kg, 2),
                })
                materials.append({
                    "자재명": "베이스판",
                    "규격": f"{d.베이스판_가로_mm}×{d.베이스판_세로_mm}×{d.베이스판_두께_mm}mm 알루미늄철판",
                    "단위": "EA",
                    "수량": result.총기둥수,
                    "단위무게_kg": d.베이스판_무게_kg,
                    "총무게_kg": round(result.총기둥수 * d.베이스판_무게_kg, 2),
                })

        elif config.설치형태 in (8, 9, 10, 14):
            # 대형 기초
            if result.총기둥수 > 0:
                materials.append({
                    "자재명": "대형기초",
                    "규격": f"{d.대형기초_가로_mm}×{d.대형기초_세로_mm}×{d.대형기초_높이_mm}mm 콘크리트",
                    "단위": "EA",
                    "수량": result.총기둥수,
                    "단위무게_kg": d.대형기초_무게_kg,
                    "총무게_kg": round(result.총기둥수 * d.대형기초_무게_kg, 2),
                })
                materials.append({
                    "자재명": "베이스판",
                    "규격": f"{d.베이스판_가로_mm}×{d.베이스판_세로_mm}×{d.베이스판_두께_mm}mm 알루미늄철판",
                    "단위": "EA",
                    "수량": result.총기둥수,
                    "단위무게_kg": d.베이스판_무게_kg,
                    "총무게_kg": round(result.총기둥수 * d.베이스판_무게_kg, 2),
                })

        # 기둥 자재 (C형강 기둥)
        if result.총기둥수 > 0 and config.설치형태 not in (5, 6):
            기둥길이_m = result.구조물높이_m + 0.3  # 기초 매립 여유
            기둥총길이 = round(result.총기둥수 * 기둥길이_m, 2)
            materials.append({
                "자재명": "기둥(C형강)",
                "규격": d.기둥_C형강_규격,
                "단위": "m",
                "수량": 기둥총길이,
                "단위무게_kg": d.기둥_단위무게_kg_m,
                "총무게_kg": round(기둥총길이 * d.기둥_단위무게_kg_m, 2),
            })

        # 주보 (가로 방향)
        if result.가로기둥수 > 0 and config.설치형태 not in (5, 6):
            주보총길이 = round(result.세로기둥수 * result.구조물가로_m * config.배열의세로, 2)
            materials.append({
                "자재명": "주보(각관)",
                "규격": d.주보_규격,
                "단위": "m",
                "수량": 주보총길이,
                "단위무게_kg": d.주보_단위무게_kg_m,
                "총무게_kg": round(주보총길이 * d.주보_단위무게_kg_m, 2),
            })

        # 연결재/보강재
        if result.가로기둥수 > 0 and config.설치형태 not in (5, 6):
            연결재길이 = round(result.가로기둥수 * result.구조물세로_m, 2)
            materials.append({
                "자재명": "연결재(C형강)",
                "규격": d.연결재_규격,
                "단위": "m",
                "수량": 연결재길이,
                "단위무게_kg": d.연결재_단위무게_kg_m,
                "총무게_kg": round(연결재길이 * d.연결재_단위무게_kg_m, 2),
            })

        result.자재목록 = materials
        result.총무게_kg = round(sum(m["총무게_kg"] for m in materials), 2)

        return result

    # -----------------------------------------------------------------------
    # 발전량 추정
    # -----------------------------------------------------------------------

    def estimate_generation(
        self,
        capacity_kW: float,
        location_irradiance_kWh_m2_day: float = 3.5,
        system_efficiency: float = 0.85,
        performance_ratio: float = 0.80,
    ) -> dict:
        """
        연간 발전량 추정
        발전량(kWh/년) = 설비용량(kW) × 일조시간(h/일) × 365 × PR
        """
        daily_gen = capacity_kW * location_irradiance_kWh_m2_day * performance_ratio
        annual_gen = daily_gen * 365

        return {
            "설비용량_kW": capacity_kW,
            "일평균일조시간_h": location_irradiance_kWh_m2_day,
            "성능계수_PR": performance_ratio,
            "일발전량_kWh": round(daily_gen, 1),
            "월발전량_kWh": round(annual_gen / 12, 0),
            "연발전량_kWh": round(annual_gen, 0),
        }
