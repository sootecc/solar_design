"""
태양광 발전소 수익 분석 계산기
VBA 원본: Module4.bas 장기수익분석으로(), 20년수익분석 시트 로직
"""
import math
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class RevenueInput:
    """수익 분석 입력 파라미터"""
    # 설비 정보
    설비용량_kW: float = 0.0            # 총 설비 용량 (kW)
    일평균발전시간_h: float = 3.6        # 일평균 발전시간 (h/일)
    환경효율_pct: float = 95.78         # 환경 효율 (%, 음영·온도·먼지 손실 반영)
    인버터효율_pct: float = 97.5        # 인버터 변환 효율 (%)
    모듈1년차효율_pct: float = 98.5     # 1년차 모듈 효율 (%)
    모듈연간감소율_pct: float = 0.33    # 연간 효율 감소율 (%/년)
    가중치: float = 1.0                 # REC 가중치

    # 전력 판매 단가
    SMP_원kWh: float = 147.84          # SMP (원/kWh)
    REC_원kWh: float = 70.22           # REC 단가 (원/kREC)
    판매구분: str = "일반판매"          # 일반판매 / 상계거래

    # 투자비
    총투자비_원: float = 0.0           # 총 투자비 (원)
    자기자본_원: float = 0.0           # 자기 자본 (원)
    타인자본_원: float = 0.0           # 대출금 (원)
    대출이자율_pct: float = 3.5        # 대출 이자율 (%/년)
    대출기간_년: int = 15              # 대출 상환 기간
    거치기간_년: int = 0               # 거치 기간 (이자만 납부, 원금 유예)

    # 운영 비용
    연간운영비_원: float = 1_500_000   # 연간 유지관리비 (원)
    연간보험비_원: float = 0           # 연간 보험비 (원)

    # 분석 기간
    분석기간_년: int = 20              # 분석 기간 (년)

    # 세금
    소득세율_pct: float = 0.0          # 소득세율 (%) - 0이면 세금 미적용


@dataclass
class YearlyResult:
    """연도별 수익 분석 결과"""
    년차: int
    발전량_kWh: float
    연수입_원: float
    연운영비_원: float
    대출원리금_원: float
    순수입_원: float
    누적순수입_원: float
    IRR: float = 0.0


@dataclass
class RevenueResult:
    """수익 분석 최종 결과"""
    연도별: list = field(default_factory=list)
    총발전량_kWh: float = 0.0
    총수입_원: float = 0.0
    총운영비_원: float = 0.0
    총순수입_원: float = 0.0
    투자회수기간_년: float = 0.0
    내부수익률_pct: float = 0.0
    연평균수익률_pct: float = 0.0
    kW당투자비_원: float = 0.0
    입력값: RevenueInput = None


class RevenueCalculator:
    """20년/30년 수익 분석 계산기"""

    def calc_annual_generation(
        self,
        inp: RevenueInput,
        year: int,
    ) -> float:
        """
        연간 발전량 계산 (kWh)
        VBA: 발전량 = 용량 × 발전시간 × 365 × 환경효율 × 인버터효율 × 모듈효율
        """
        # 모듈 효율 계산 (1년차 이후 연간 감소)
        if year == 1:
            module_eff = inp.모듈1년차효율_pct / 100
        else:
            module_eff = (inp.모듈1년차효율_pct / 100) * ((1 - inp.모듈연간감소율_pct / 100) ** (year - 1))

        gen = (inp.설비용량_kW
               * inp.일평균발전시간_h
               * 365
               * (inp.환경효율_pct / 100)
               * (inp.인버터효율_pct / 100)
               * module_eff)
        return round(gen, 2)

    def calc_annual_revenue(self, inp: RevenueInput, generation_kWh: float) -> float:
        """
        연간 수입 계산 (원)
        일반판매: 발전량 × (SMP + REC × 가중치)
        상계거래: 발전량 × 한전 전력판매단가
        VBA: 20년수익분석 시트 연수입 컬럼
        """
        if inp.판매구분 == "상계거래":
            # 상계거래: SMP만 적용 (한전 요금 기준)
            return round(generation_kWh * inp.SMP_원kWh)
        else:
            # 일반판매 (RPS): SMP + REC
            return round(generation_kWh * (inp.SMP_원kWh + inp.REC_원kWh * inp.가중치))

    def calc_loan_payment(self, inp: RevenueInput, year: int) -> float:
        """
        연간 대출 원리금 계산
        - 거치기간: 이자만 납부 (원금 × 이자율)
        - 상환기간: 원리금균등상환 = 잔여대출금 × r / (1-(1+r)^-n)
        """
        if inp.타인자본_원 <= 0 or inp.대출기간_년 <= 0:
            return 0.0
        total_years = inp.거치기간_년 + inp.대출기간_년
        if year > total_years:
            return 0.0

        r = inp.대출이자율_pct / 100

        # 거치기간: 이자만
        if year <= inp.거치기간_년:
            return round(inp.타인자본_원 * r)

        # 상환기간: 원리금균등상환
        n = inp.대출기간_년
        if r == 0:
            return round(inp.타인자본_원 / n)
        payment = inp.타인자본_원 * r / (1 - (1 + r) ** (-n))
        return round(payment)

    def calc_irr(self, cashflows: list) -> float:
        """
        내부수익률(IRR) 계산
        뉴턴-랩슨 방법 사용
        cashflows[0] = 초기투자비(음수), cashflows[1:] = 연간순수입
        """
        if not cashflows or cashflows[0] >= 0:
            return 0.0

        rate = 0.1
        for _ in range(100):
            npv = sum(cf / (1 + rate) ** i for i, cf in enumerate(cashflows))
            dnpv = sum(-i * cf / (1 + rate) ** (i + 1) for i, cf in enumerate(cashflows))
            if abs(dnpv) < 1e-10:
                break
            rate_new = rate - npv / dnpv
            if abs(rate_new - rate) < 1e-8:
                rate = rate_new
                break
            rate = rate_new

        return round(rate * 100, 4) if -1 < rate < 10 else 0.0

    def calc_payback_period(self, yearly_results: list, initial_investment: float) -> float:
        """투자 회수 기간 계산 (년)"""
        cumulative = 0.0
        for r in yearly_results:
            cumulative += r.순수입_원
            if cumulative >= initial_investment:
                # 선형 보간
                prev = cumulative - r.순수입_원
                fraction = (initial_investment - prev) / r.순수입_원
                return round(r.년차 - 1 + fraction, 1)
        return float(len(yearly_results))  # 회수 못함

    def analyze(self, inp: RevenueInput) -> RevenueResult:
        """
        전체 수익 분석 실행
        VBA 원본: 20년수익분석 시트의 행별 계산 로직
        """
        result = RevenueResult(입력값=inp)

        # kW당 투자비
        if inp.설비용량_kW > 0:
            result.kW당투자비_원 = round(inp.총투자비_원 / inp.설비용량_kW)

        # 현금흐름 (IRR용): 초기 투자비 음수
        cashflows = [-inp.자기자본_원]

        yearly_results = []
        누적순수입 = 0.0

        for year in range(1, inp.분석기간_년 + 1):
            # 발전량
            gen = self.calc_annual_generation(inp, year)

            # 수입
            revenue = self.calc_annual_revenue(inp, gen)

            # 비용
            운영비 = inp.연간운영비_원 + inp.연간보험비_원
            대출원리금 = self.calc_loan_payment(inp, year)
            총비용 = 운영비 + 대출원리금

            # 소득세 (설정 시)
            과세소득 = max(0, revenue - 총비용)
            세금 = round(과세소득 * inp.소득세율_pct / 100)

            순수입 = revenue - 총비용 - 세금
            누적순수입 += 순수입

            yr = YearlyResult(
                년차=year,
                발전량_kWh=gen,
                연수입_원=revenue,
                연운영비_원=운영비,
                대출원리금_원=대출원리금,
                순수입_원=순수입,
                누적순수입_원=round(누적순수입),
            )
            yearly_results.append(yr)
            cashflows.append(순수입)

        # IRR 계산
        irr = self.calc_irr(cashflows)
        for yr in yearly_results:
            yr.IRR = irr

        result.연도별 = yearly_results
        result.총발전량_kWh = round(sum(r.발전량_kWh for r in yearly_results))
        result.총수입_원 = round(sum(r.연수입_원 for r in yearly_results))
        result.총운영비_원 = round(sum(r.연운영비_원 + r.대출원리금_원 for r in yearly_results))
        result.총순수입_원 = round(sum(r.순수입_원 for r in yearly_results))
        result.내부수익률_pct = irr
        result.투자회수기간_년 = self.calc_payback_period(yearly_results, inp.자기자본_원)
        result.연평균수익률_pct = round(irr, 2)

        return result
