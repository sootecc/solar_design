"""
태양광 단선결선도 설계 계산기
VBA 원본: Module6.bas 단선결선도작성(), 스트링 설계 로직
"""
import io
import math
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class StringDesign:
    """스트링 설계 결과"""
    직렬모듈수: int = 0          # 직렬 연결 모듈 수
    병렬스트링수: int = 0        # 병렬 스트링 수
    MPPT당스트링수: int = 0      # MPPT 채널당 스트링 수
    총모듈수: int = 0
    개방전압_V: float = 0.0      # 스트링 개방전압 (직렬×Voc)
    동작전압_V: float = 0.0      # 스트링 동작전압 (직렬×Vmax)
    스트링전류_A: float = 0.0    # 스트링 전류
    배열용량_kW: float = 0.0
    인버터적합: bool = True
    경고메시지: list = field(default_factory=list)


@dataclass
class CircuitElement:
    """단선결선도 구성 요소"""
    종류: str           # 'module', 'junction', 'inverter', 'meter', 'grid', 'string_cable'
    레이블: str
    위치_x: int = 0
    위치_y: int = 0


@dataclass
class SingleLineDiagram:
    """단선결선도 데이터"""
    스트링목록: list = field(default_factory=list)   # StringDesign list
    인버터수: int = 1
    인버터품명: str = ""
    인버터용량_kW: float = 0.0
    공급방식: str = "단상"       # 단상 / 삼상
    접속함수: int = 1           # 접속함 수량
    텍스트도면: str = ""        # ASCII 텍스트 도면
    요약: dict = field(default_factory=dict)


class ElectricalDesigner:
    """전기 설계 계산기"""

    def design_string(
        self,
        모듈용량_W: float,
        모듈Voc_V: float,
        모듈Vmax_V: float,
        모듈Isc_A: float = 10.0,
        총모듈수: int = 8,
        인버터용량_kW: float = 3.7,
        MPPT범위: str = "185-400",
        MPPT수: int = 1,
        MPPT당스트링수: int = 1,
        단상삼상: str = "단상",
    ) -> StringDesign:
        """
        스트링 설계
        VBA 원본: Module6.bas 스트링직렬수 결정 로직
        """
        result = StringDesign(총모듈수=총모듈수)

        # MPPT 전압 범위 파싱
        try:
            parts = MPPT범위.replace("V", "").split("-")
            mppt_min = float(parts[0])
            mppt_max = float(parts[1])
        except Exception:
            mppt_min, mppt_max = 200.0, 800.0

        # 직렬 모듈 수 결정
        # 조건: MPPT범위_최소 ≤ 직렬수 × Vmax ≤ MPPT범위_최대
        if 모듈Vmax_V <= 0:
            직렬수 = 1
        else:
            직렬수_최대 = int(mppt_max / 모듈Vmax_V)
            직렬수_최소 = max(1, math.ceil(mppt_min / 모듈Vmax_V))
            직렬수 = min(직렬수_최대, max(직렬수_최소, 1))

        # 병렬 스트링 수 (인버터 용량 기반)
        스트링용량_kW = 직렬수 * 모듈용량_W / 1000
        if 스트링용량_kW > 0:
            병렬수 = max(1, round(총모듈수 / 직렬수))
        else:
            병렬수 = 1

        # MPPT당 스트링 수
        if MPPT수 > 0:
            MPPT당 = math.ceil(병렬수 / MPPT수)
        else:
            MPPT당 = 병렬수

        # 전압/전류 계산
        개방전압 = round(직렬수 * 모듈Voc_V, 1)
        동작전압 = round(직렬수 * 모듈Vmax_V, 1)
        스트링전류 = round(모듈Isc_A, 1)

        # 배열 용량
        배열용량 = round(총모듈수 * 모듈용량_W / 1000, 2)

        # 적합성 검토
        경고 = []
        if 개방전압 > 1000:
            경고.append(f"개방전압 {개방전압}V가 시스템 최대전압(1000V) 초과")
        if 동작전압 > mppt_max:
            경고.append(f"동작전압 {동작전압}V가 MPPT최대({mppt_max}V) 초과")
        if 동작전압 < mppt_min:
            경고.append(f"동작전압 {동작전압}V가 MPPT최소({mppt_min}V) 미달")

        인버터적합 = 배열용량 <= 인버터용량_kW * 1.2 and len(경고) == 0

        result.직렬모듈수 = 직렬수
        result.병렬스트링수 = 병렬수
        result.MPPT당스트링수 = MPPT당
        result.개방전압_V = 개방전압
        result.동작전압_V = 동작전압
        result.스트링전류_A = 스트링전류
        result.배열용량_kW = 배열용량
        result.인버터적합 = 인버터적합
        result.경고메시지 = 경고

        return result

    def generate_diagram(
        self,
        string_design: StringDesign,
        인버터품명: str,
        인버터용량_kW: float,
        단상삼상: str,
        인버터수: int = 1,
        모듈품명: str = "",
        모듈용량_W: float = 0,
    ) -> SingleLineDiagram:
        """
        단선결선도 생성
        VBA 원본: Module6.bas Shape 기반 도면 → 텍스트 기반으로 재구현
        """
        diagram = SingleLineDiagram(
            인버터수=인버터수,
            인버터품명=인버터품명,
            인버터용량_kW=인버터용량_kW,
            공급방식=단상삼상,
        )

        # 접속함 수량 결정 (병렬스트링이 4개 초과 시 접속함 필요)
        if string_design.병렬스트링수 > 4:
            접속함수 = math.ceil(string_design.병렬스트링수 / 4)
        else:
            접속함수 = 1
        diagram.접속함수 = 접속함수

        # ── 텍스트 단선결선도 생성 ──────────────────────
        lines = []
        sep = "─" * 60

        lines.append("=" * 60)
        lines.append("   단 선 결 선 도 (Single Line Diagram)")
        lines.append("=" * 60)

        # 모듈/스트링 표시
        lines.append("")
        lines.append("[ 태양전지 어레이 ]")
        lines.append(f"  모듈: {모듈품명} ({모듈용량_W}W)")
        lines.append(f"  직렬: {string_design.직렬모듈수}직 × 병렬: {string_design.병렬스트링수}병")
        lines.append(f"  총 모듈: {string_design.총모듈수}장 / 용량: {string_design.배열용량_kW}kW")
        lines.append(f"  스트링 개방전압: {string_design.개방전압_V}V")
        lines.append(f"  스트링 동작전압: {string_design.동작전압_V}V")

        # 스트링별 표현
        lines.append("")
        for i in range(string_design.병렬스트링수):
            modules = " → ".join(["[M]"] * min(string_design.직렬모듈수, 6))
            if string_design.직렬모듈수 > 6:
                modules += f" → ... ({string_design.직렬모듈수}직)"
            lines.append(f"  STR {i+1:02d}: {modules}")

        lines.append("")
        lines.append("         │  (DC)")
        lines.append("         ↓")

        # 접속함
        if 접속함수 > 0:
            lines.append("")
            for j in range(접속함수):
                lines.append(f"  ┌─────────────────────────┐")
                lines.append(f"  │   접속함(DC) #{j+1}          │")
                strs_in_box = min(4, string_design.병렬스트링수 - j * 4)
                lines.append(f"  │   입력: {strs_in_box}스트링           │")
                lines.append(f"  │   역전류방지 다이오드    │")
                lines.append(f"  └─────────────────────────┘")
            lines.append("         │  (DC합산)")
            lines.append("         ↓")

        # 인버터
        lines.append("")
        for k in range(인버터수):
            lines.append(f"  ┌─────────────────────────────┐")
            lines.append(f"  │   인버터 #{k+1}                  │")
            lines.append(f"  │   {인버터품명[:25]:<25}│")
            lines.append(f"  │   용량: {인버터용량_kW}kW / {단상삼상}      │")
            lines.append(f"  │   DC→AC 변환                │")
            lines.append(f"  └─────────────────────────────┘")

        lines.append("         │  (AC)")
        lines.append("         ↓")

        # 분전반
        lines.append("")
        lines.append("  ┌───────────────────┐")
        lines.append("  │   분전반(AC)       │")
        lines.append("  │   누전차단기        │")
        lines.append("  │   전력량계(계량기)  │")
        lines.append("  └───────────────────┘")
        lines.append("         │")
        lines.append("         ↓")

        # 계통 연결
        if 단상삼상 == "단상":
            lines.append("  ═══════════════════════")
            lines.append("   한전 배전선로 (단상 220V)")
        else:
            lines.append("  ═══════════════════════")
            lines.append("   한전 배전선로 (삼상 380V)")
        lines.append("  ═══════════════════════")

        lines.append("")
        lines.append(sep)

        # 요약 정보
        diagram.요약 = {
            "총모듈수": string_design.총모듈수,
            "배열용량_kW": string_design.배열용량_kW,
            "직렬모듈수": string_design.직렬모듈수,
            "병렬스트링수": string_design.병렬스트링수,
            "개방전압_V": string_design.개방전압_V,
            "동작전압_V": string_design.동작전압_V,
            "접속함수": 접속함수,
            "인버터수": 인버터수,
            "인버터용량_kW": 인버터용량_kW,
            "공급방식": 단상삼상,
            "인버터적합": string_design.인버터적합,
        }

        diagram.텍스트도면 = "\n".join(lines)
        return diagram

    def generate_diagram_image(
        self,
        string_design: StringDesign,
        인버터품명: str,
        인버터용량_kW: float,
        단상삼상: str,
        인버터수: int = 1,
        모듈품명: str = "",
        모듈용량_W: float = 0,
    ) -> bytes:
        """
        matplotlib 기반 단선결선도 PNG 이미지 생성
        반환: PNG bytes (st.image() 에 바로 전달 가능)
        """
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import matplotlib.patches as mpatches
        from matplotlib.patches import FancyBboxPatch
        import matplotlib.font_manager as fm

        # 한글 폰트 자동 설정 (우선순위: 맑은고딕 > NanumGothic > Gulim)
        _kr_candidates = ["Malgun Gothic", "NanumGothic", "Gulim", "Dotum", "Batang"]
        _kr_font = next(
            (f.name for c in _kr_candidates
             for f in fm.fontManager.ttflist if f.name == c),
            None,
        )
        if _kr_font:
            plt.rcParams["font.family"] = _kr_font
        plt.rcParams["axes.unicode_minus"] = False

        병렬수 = string_design.병렬스트링수
        직렬수 = string_design.직렬모듈수
        접속함수 = math.ceil(병렬수 / 4) if 병렬수 > 4 else 1

        # ── 캔버스 크기: 스트링 수에 비례 ──────────────
        fig_h = max(10, 4 + 병렬수 * 0.55 + 인버터수 * 1.2)
        fig, ax = plt.subplots(figsize=(10, fig_h))
        ax.set_xlim(0, 10)
        ax.set_ylim(0, fig_h)
        ax.axis("off")
        ax.set_facecolor("#f8f9fa")
        fig.patch.set_facecolor("#f8f9fa")

        # 색상
        C_PANEL  = "#dce8f8"   # 박스 배경
        C_BORDER = "#2c5f8a"   # 박스 테두리
        C_WIRE   = "#2c5f8a"   # 배선
        C_STR    = "#e8f4e8"   # 스트링 배경
        C_INV    = "#fff3cd"   # 인버터 배경
        C_GRID   = "#f8d7da"   # 계통 배경
        C_MOD    = "#c8e6c9"   # 모듈 셀 색
        FONT_KR  = {"fontfamily": "sans-serif", "fontsize": 8}

        def box(x, y, w, h, color=C_PANEL, border=C_BORDER, lw=1.5, radius=0.15):
            rect = FancyBboxPatch(
                (x, y), w, h,
                boxstyle=f"round,pad=0,rounding_size={radius}",
                facecolor=color, edgecolor=border, linewidth=lw,
            )
            ax.add_patch(rect)

        def wire(x1, y1, x2, y2):
            ax.plot([x1, x2], [y1, y2], color=C_WIRE, lw=1.5, solid_capstyle="round")

        def arrow(x, y, dy=-0.4):
            ax.annotate(
                "", xy=(x, y + dy), xytext=(x, y),
                arrowprops=dict(arrowstyle="->", color=C_WIRE, lw=1.5),
            )

        def label(x, y, txt, ha="center", va="center", size=8, bold=False, color="black"):
            w = "bold" if bold else "normal"
            ax.text(x, y, txt, ha=ha, va=va, fontsize=size, fontweight=w, color=color)

        # ─── 레이아웃 y 좌표 (아래→위 순서) ─────────────
        y_cur = fig_h - 0.5   # 최상단 여백

        # ── 제목 ────────────────────────────────────────
        label(5, y_cur, "단 선 결 선 도  (Single Line Diagram)", size=12, bold=True)
        y_cur -= 0.8

        # ── 어레이 영역 ─────────────────────────────────
        str_area_h = 병렬수 * 0.55 + 0.7
        str_top = y_cur
        str_bot = y_cur - str_area_h
        box(0.3, str_bot, 9.4, str_area_h, color="#eaf4ea", border="#388e3c", lw=2)
        label(5, str_top - 0.25, f"태양전지 어레이  ({모듈품명}  {모듈용량_W}W)", size=9, bold=True, color="#1b5e20")

        # 스트링 개별 표시
        str_y_positions = []
        for i in range(병렬수):
            sy = str_top - 0.6 - i * 0.55
            str_y_positions.append(sy)
            # 모듈 셀 (최대 6개 표시)
            n_show = min(직렬수, 6)
            cell_w = 0.5
            start_x = 0.6
            for m in range(n_show):
                cx = start_x + m * (cell_w + 0.05)
                box(cx, sy - 0.18, cell_w, 0.36, color=C_MOD, border="#388e3c", lw=0.8, radius=0.04)
                if m == 0:
                    label(cx + cell_w / 2, sy, "M", size=6, color="#1b5e20")
                else:
                    label(cx + cell_w / 2, sy, "M", size=6, color="#1b5e20")
            # 생략 표시
            end_x = start_x + n_show * (cell_w + 0.05)
            if 직렬수 > 6:
                label(end_x + 0.1, sy, f"···({직렬수}직)", ha="left", size=7, color="#555")
            # 스트링 레이블
            label(0.5, sy, f"S{i+1:02d}", ha="right", size=7, color="#333")
            # 오른쪽 집선 수평선
            wire(end_x + (0.35 if 직렬수 > 6 else 0), sy, 9.0, sy)

        # 오른쪽 수직 집선
        wire(9.0, str_y_positions[0], 9.0, str_y_positions[-1])

        # 집선 → 아래 화살표
        mid_y_str = (str_y_positions[0] + str_y_positions[-1]) / 2
        wire(9.0, mid_y_str, 5.0, mid_y_str)
        wire(5.0, mid_y_str, 5.0, str_bot)
        y_cur = str_bot
        arrow(5.0, y_cur, dy=-0.35)
        label(5.3, y_cur - 0.15, "DC", size=7, color="#555")
        y_cur -= 0.5

        # ── 접속함 ──────────────────────────────────────
        jbox_h = 0.9 * 접속함수 + 0.3
        box(2.5, y_cur - jbox_h, 5.0, jbox_h, color=C_PANEL, border=C_BORDER, lw=1.5)
        label(5.0, y_cur - 0.3, "접속함 (DC Junction Box)", size=9, bold=True, color=C_BORDER)
        for j in range(접속함수):
            jy = y_cur - 0.7 - j * 0.9
            n_str = min(4, 병렬수 - j * 4)
            label(5.0, jy, f"#{j+1}  입력 {n_str}스트링  |  역전류방지 다이오드", size=8, color="#333")
        y_cur -= jbox_h
        arrow(5.0, y_cur, dy=-0.35)
        label(5.3, y_cur - 0.15, "DC 합산", size=7, color="#555")
        y_cur -= 0.5

        # ── 인버터 ──────────────────────────────────────
        inv_h = 1.0 * 인버터수 + 0.5
        box(1.8, y_cur - inv_h, 6.4, inv_h, color=C_INV, border="#e65100", lw=2)
        label(5.0, y_cur - 0.28, "인 버 터  (Inverter)", size=9, bold=True, color="#bf360c")
        for k in range(인버터수):
            iy = y_cur - 0.7 - k * 1.0
            short_name = 인버터품명[:26] if len(인버터품명) > 26 else 인버터품명
            label(5.0, iy, f"#{k+1}  {short_name}  {인버터용량_kW}kW  {단상삼상}  (DC→AC)", size=8, color="#333")
        y_cur -= inv_h
        arrow(5.0, y_cur, dy=-0.35)
        label(5.3, y_cur - 0.15, "AC", size=7, color="#555")
        y_cur -= 0.5

        # ── 분전반 ──────────────────────────────────────
        dist_h = 1.0
        box(2.8, y_cur - dist_h, 4.4, dist_h, color=C_PANEL, border=C_BORDER, lw=1.5)
        label(5.0, y_cur - 0.3, "분전반  (누전차단기 · 전력량계)", size=9, bold=True, color=C_BORDER)
        label(5.0, y_cur - 0.7, "계량기 → 한전 연계", size=8, color="#555")
        y_cur -= dist_h
        arrow(5.0, y_cur, dy=-0.35)
        y_cur -= 0.5

        # ── 계통 ────────────────────────────────────────
        grid_h = 0.7
        volt = "단상 220V" if 단상삼상 == "단상" else "삼상 380V"
        box(1.5, y_cur - grid_h, 7.0, grid_h, color=C_GRID, border="#c62828", lw=2)
        label(5.0, y_cur - grid_h / 2, f"한전 배전선로  ({volt})", size=10, bold=True, color="#b71c1c")
        y_cur -= grid_h

        # ── 우측 요약 박스 ───────────────────────────────
        sx = 0.05
        sy_top = fig_h - 1.5
        summary_lines = [
            ("총 모듈", f"{string_design.총모듈수}장"),
            ("배열 용량", f"{string_design.배열용량_kW:.2f} kW"),
            ("직렬×병렬", f"{직렬수}직 × {병렬수}병"),
            ("개방전압", f"{string_design.개방전압_V} V"),
            ("동작전압", f"{string_design.동작전압_V} V"),
            ("접속함", f"{접속함수}개"),
            ("인버터", f"{인버터수}대"),
        ]
        box(sx, sy_top - len(summary_lines) * 0.42 - 0.2, 1.35, len(summary_lines) * 0.42 + 0.4,
            color="#fffde7", border="#f9a825", lw=1.2)
        label(sx + 0.675, sy_top - 0.05, "설계 요약", size=7, bold=True, color="#e65100")
        for li, (k, v) in enumerate(summary_lines):
            label(sx + 0.1, sy_top - 0.35 - li * 0.42, k, ha="left", size=6.5, color="#555")
            label(sx + 1.25, sy_top - 0.35 - li * 0.42, v, ha="right", size=6.5, bold=True, color="#222")

        # 경고 표시
        if string_design.경고메시지:
            warn_y = 0.6
            for wi, wmsg in enumerate(string_design.경고메시지):
                label(5.0, warn_y - wi * 0.35, f"⚠ {wmsg}", size=8, color="#c62828", bold=True)

        plt.tight_layout(pad=0.3)
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=120, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        return buf.read()
