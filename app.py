"""
태양광 가설계 자동화 시스템 v2
원본: 태양광현장분석프로그램(20260324)-Ver15-0-0_배포용.xlsm VBA 이식
기능: 가설계, 자재산출, 전기설계(단선결선도), 지역일사량, 20년 수익분석
"""
import streamlit as st
import pandas as pd
import math
from datetime import date

from data.modules import MODULE_DATA, get_module_names
from data.inverters import INVERTER_DATA, get_inverter_names
from data.irradiance import SUNSHINE_DATA, search_region, get_sunshine, get_irradiance_kwh
from core.models import (
    ArrayConfig, ModuleSpec, InverterSpec,
    StructureDefaults, ProjectInfo, INSTALLATION_TYPES, DIRECTIONS
)
from core.calculator import SolarCalculator
from core.revenue import RevenueCalculator, RevenueInput
from core.electrical import ElectricalDesigner
from core.layout import KakaoMap, LayoutDrawer, ArrayPosition, ARRAY_COLORS
from utils.output import export_to_excel, results_to_dataframe

# ──────────────────────────────────────────────────────────
st.set_page_config(
    page_title="태양광 가설계 시스템",
    page_icon="☀️",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown("""
<style>
.main-title{font-size:1.8rem;font-weight:700;color:#1F4E79}
.sec{font-size:1rem;font-weight:600;color:#2E75B6;border-bottom:2px solid #2E75B6;padding-bottom:3px;margin-top:.8rem}
.card{background:#EBF4FF;border-radius:8px;padding:10px 14px;margin:6px 0;font-size:.9rem}
.ok{background:#D4EDDA;border-left:4px solid #28A745;padding:8px 12px;border-radius:4px}
.warn{background:#FFF3CD;border-left:4px solid #FFC107;padding:8px 12px;border-radius:4px}
.err{background:#F8D7DA;border-left:4px solid #DC3545;padding:8px 12px;border-radius:4px}
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────
# 헬퍼
# ──────────────────────────────────────────────────────────
def make_module_spec(idx):
    m = MODULE_DATA[idx]
    return ModuleSpec(index=idx, 품명=m[1], 용량_W=m[2], 가로_mm=m[3], 세로_mm=m[4],
                      두께_mm=m[5], Voc_V=m[6], Vmax_V=m[7], 무게_kg=m[8],
                      효율_pct=m[9], 연간효율감소_pct=m[11])

def make_inverter_spec(idx):
    inv = INVERTER_DATA[idx]
    return InverterSpec(index=idx, 품명=inv[0], 용량_kW=inv[1], 단상삼상=inv[2],
                        DC입력전압=inv[3], MPPT범위=inv[4],
                        효율_pct=inv[6], 무게_kg=inv[7],
                        MPPT수=inv[8], MPPT1스트링수=inv[9], MPPT2스트링수=inv[10])

def fmt_won(v):
    return f"{v:,.0f} 원"

# ──────────────────────────────────────────────────────────
# 사이드바
# ──────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ☀️ 프로젝트 정보")
    현장명 = st.text_input("현장명", placeholder="예) ○○ 태양광발전소")
    주소   = st.text_input("주소",   placeholder="예) 전남 나주시 ...")
    설계자 = st.text_input("설계자")
    설계일 = st.date_input("설계일", value=date.today())

    st.markdown("---")
    st.markdown("### 배열/인버터")
    총배열수 = st.number_input("총 배열 수", 1, 30, 1)
    inv_names = get_inverter_names()
    inv_idx = st.selectbox("인버터 모델", range(len(inv_names)),
                           format_func=lambda i: inv_names[i])
    인버터수량 = st.number_input("인버터 수량", 1, 50, 1)

    st.markdown("---")
    st.markdown("### 지역 일사량")
    지역검색어 = st.text_input("지역 검색", placeholder="예) 나주, 서울, 충주")
    지역목록 = []
    선택지역 = None
    if 지역검색어:
        지역목록 = search_region(지역검색어)
    if not 지역목록:
        지역목록 = sorted(SUNSHINE_DATA.keys())
    선택지역 = st.selectbox("지역 선택", 지역목록)

    sun_data = get_sunshine(선택지역) if 선택지역 else None
    일평균일조시간 = sun_data["일평균_h"] if sun_data else 3.5
    st.caption(f"일평균 일조시간: {일평균일조시간:.2f} h/일")

    pr = st.slider("성능계수 PR", 0.70, 0.95, 0.80, 0.01)

    st.markdown("---")
    st.markdown("### 수익 분석 기본값")
    SMP단가  = st.number_input("SMP 단가 (원/kWh)", 50.0, 300.0, 147.84, 0.01)
    REC단가  = st.number_input("REC 단가 (원/REC)", 10.0, 200.0, 70.22, 0.01)
    REC가중치 = st.number_input("REC 가중치", 0.5, 5.0, 1.0, 0.1)
    판매구분  = st.selectbox("전력 판매 구분", ["일반판매", "상계거래"])


# ──────────────────────────────────────────────────────────
# 배열 설정 탭
# ──────────────────────────────────────────────────────────
tab_labels = [f"배열 {i+1}" for i in range(총배열수)]
tab_labels += ["계산결과", "자재표", "전기설계", "배치도", "지역일사량", "수익분석"]
tabs = st.tabs(tab_labels)

array_configs = []
module_specs  = []
calc = SolarCalculator()

for arr_idx in range(총배열수):
    with tabs[arr_idx]:
        st.markdown(f'<div class="sec">배열 {arr_idx+1} 구성</div>', unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)

        with c1:
            st.markdown("**모듈**")
            mod_names = get_module_names()
            mod_idx = st.selectbox("모듈 선택", range(len(mod_names)),
                                   format_func=lambda i: mod_names[i],
                                   key=f"mod_{arr_idx}")
            m = MODULE_DATA[mod_idx]
            st.caption(f"{m[3]}×{m[4]}mm / {m[8]}kg / Voc {m[6]}V")

            st.markdown("**배열 구성**")
            가로수 = st.number_input("가로 수",  1, 50, 4, key=f"c_{arr_idx}")
            세로수 = st.number_input("세로 수",  1, 20, 2, key=f"r_{arr_idx}")
            앞삭제 = st.number_input("앞쪽 삭제", 0, 가로수*세로수, 0, key=f"df_{arr_idx}")
            뒤삭제 = st.number_input("뒤쪽 삭제", 0, 가로수*세로수, 0, key=f"db_{arr_idx}")
            추가   = st.number_input("중앙 추가", 0, 50, 0, key=f"ac_{arr_idx}")

        with c2:
            st.markdown("**설치 조건**")
            형태_opt = list(INSTALLATION_TYPES.items())
            형태_idx = st.selectbox("설치형태", range(len(형태_opt)),
                                    format_func=lambda i: f"{형태_opt[i][0]}. {형태_opt[i][1]}",
                                    index=2, key=f"tp_{arr_idx}")
            형태코드 = 형태_opt[형태_idx][0]
            각도  = st.slider("경사각 (°)", 0, 90, 15, key=f"ang_{arr_idx}")
            높이  = st.number_input("최하단 높이 (m)", 0.0, 10.0, 0.5, 0.1, key=f"ht_{arr_idx}")
            방향  = st.selectbox("설치 방향", DIRECTIONS, key=f"dir_{arr_idx}")
            방위각 = st.slider("방위각 오프셋 (°)", -45, 45, 0, key=f"az_{arr_idx}")

        with c3:
            st.markdown("**이격거리**")
            가로이격 = st.number_input("가로 이격 (mm)", 0, 500, 20, key=f"gh_{arr_idx}")
            세로이격 = st.number_input("세로 이격 (mm)", 0, 500, 20, key=f"gv_{arr_idx}")

            cfg = ArrayConfig(
                배열의가로=가로수, 배열의세로=세로수,
                앞쪽깔기삭제수량=앞삭제, 뒤쪽깔기삭제수량=뒤삭제,
                배열중앙추가장수=추가, 설치각도=float(각도),
                시작높이_m=높이, 설치방향=방향,
                설치방향각도=float(방위각), 설치형태=형태코드,
                모듈판가로이격거리_mm=가로이격, 모듈판세로이격거리_mm=세로이격,
            )
            mod_spec = make_module_spec(mod_idx)
            w = calc.calc_structure_width(cfg, mod_spec)
            d = calc.calc_structure_depth(cfg, mod_spec)
            h = calc.calc_structure_height(cfg, mod_spec)
            area = calc.calc_installation_area(cfg, mod_spec)
            _, _, total_col = calc.calc_column_count(cfg)

            실제수량 = cfg.실제모듈수량
            용량 = round(실제수량 * m[2] / 1000, 2)
            st.markdown(f"""
<div class="card">
<b>모듈 수량:</b> {실제수량}장 ({가로수}×{세로수})<br>
<b>배열 용량:</b> {용량} kW<br>
<b>구조물:</b> {w:.2f}m × {d:.2f}m × {h:.2f}m<br>
<b>설치 면적:</b> {area:.2f} m²<br>
<b>기둥 수:</b> {total_col}개
</div>""", unsafe_allow_html=True)

        array_configs.append(cfg)
        module_specs.append(mod_spec)


# ──────────────────────────────────────────────────────────
# 공통 계산
# ──────────────────────────────────────────────────────────
inverter_spec = make_inverter_spec(inv_idx)
총용량_kW = sum(c.실제모듈수량 * MODULE_DATA[c.모듈_index][2] / 1000 for c in array_configs)

all_results = []
for i, (cfg, ms) in enumerate(zip(array_configs, module_specs)):
    r = calc.calc_materials(cfg, ms, inverter_spec,
                            inverter_count=인버터수량,
                            array_index=i+1, total_arrays=총배열수,
                            project_capacity_kW=총용량_kW)
    all_results.append(r)

project_info = ProjectInfo(현장명=현장명, 주소=주소, 설계자=설계자,
                           설계일자=str(설계일), 총배열수=총배열수,
                           인버터_index=inv_idx, 인버터수량=인버터수량)

# 발전량 추정
gen_info = calc.estimate_generation(총용량_kW, 일평균일조시간, pr)

TAB_OFF = 총배열수


# ──────────────────────────────────────────────────────────
# 탭: 계산결과
# ──────────────────────────────────────────────────────────
with tabs[TAB_OFF + 0]:
    st.markdown('<div class="sec">가설계 요약</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 설비 용량", f"{총용량_kW:.2f} kW")
    c2.metric("총 모듈 수", f"{sum(r.모듈수량 for r in all_results):,} 장")
    c3.metric("총 설비 중량", f"{sum(r.총무게_kg for r in all_results):,.1f} kg")
    c4.metric("총 설치 면적", f"{sum(r.설치면적_m2 for r in all_results):.1f} m²")

    st.markdown('<div class="sec">배열별 요약</div>', unsafe_allow_html=True)
    df_sum = pd.DataFrame([{
        "배열": f"배열 {r.배열번호}", "모듈(장)": r.모듈수량,
        "용량(kW)": r.총용량_kW, "폭(m)": r.구조물가로_m,
        "길이(m)": r.구조물세로_m, "높이(m)": r.구조물높이_m,
        "면적(m²)": r.설치면적_m2, "기둥(개)": r.총기둥수,
        "중량(kg)": r.총무게_kg,
    } for r in all_results])
    st.dataframe(df_sum, use_container_width=True, hide_index=True)

    st.markdown('<div class="sec">발전량 추정</div>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    c1.metric("일 발전량",  f"{gen_info['일발전량_kWh']:,.1f} kWh")
    c2.metric("월 발전량",  f"{gen_info['월발전량_kWh']:,.0f} kWh")
    c3.metric("연간 발전량", f"{gen_info['연발전량_kWh']:,.0f} kWh")


# ──────────────────────────────────────────────────────────
# 탭: 자재표
# ──────────────────────────────────────────────────────────
with tabs[TAB_OFF + 1]:
    st.markdown('<div class="sec">발전설비 중량계산표</div>', unsafe_allow_html=True)
    df = results_to_dataframe(all_results)
    if df is not None:
        st.dataframe(df, use_container_width=True, hide_index=True)
        st.markdown(f"**총 발전설비 중량: {df['총무게(kg)'].sum():,.1f} kg**")
        excel_bytes = export_to_excel(all_results, project_info, gen_info)
        if excel_bytes:
            st.download_button("📥 Excel 다운로드", excel_bytes,
                               f"태양광가설계_{현장명 or '현장'}_{설계일}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ──────────────────────────────────────────────────────────
# 탭: 전기설계 (단선결선도)
# ──────────────────────────────────────────────────────────
with tabs[TAB_OFF + 2]:
    st.markdown('<div class="sec">전기 설계 / 단선결선도</div>', unsafe_allow_html=True)

    ed = ElectricalDesigner()
    ms0 = module_specs[0]
    inv = inverter_spec

    # 스트링 설계
    string_result = ed.design_string(
        모듈용량_W=ms0.용량_W,
        모듈Voc_V=ms0.Voc_V,
        모듈Vmax_V=ms0.Vmax_V,
        총모듈수=all_results[0].모듈수량,
        인버터용량_kW=inv.용량_kW,
        MPPT범위=inv.MPPT범위,
        MPPT수=inv.MPPT수,
        단상삼상=inv.단상삼상,
    )

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**스트링 설계 결과**")
        cls = "ok" if string_result.인버터적합 else "warn"
        msg = "인버터 용량 적합" if string_result.인버터적합 else "⚠ 인버터 용량 검토 필요"
        st.markdown(f"""<div class="{cls}">
<b>직렬 모듈 수:</b> {string_result.직렬모듈수} 개<br>
<b>병렬 스트링 수:</b> {string_result.병렬스트링수} 개<br>
<b>개방전압(Voc):</b> {string_result.개방전압_V} V<br>
<b>동작전압(Vmax):</b> {string_result.동작전압_V} V<br>
<b>배열 용량:</b> {string_result.배열용량_kW} kW<br>
<b>인버터 용량:</b> {inv.용량_kW} kW ({inv.단상삼상})<br>
<b>MPPT 범위:</b> {inv.MPPT범위} V<br>
<b>판정:</b> {msg}
</div>""", unsafe_allow_html=True)
        if string_result.경고메시지:
            for w in string_result.경고메시지:
                st.warning(w)

    with c2:
        st.markdown("**인버터 정보**")
        st.markdown(f"""<div class="card">
<b>모델:</b> {inv.품명}<br>
<b>용량:</b> {inv.용량_kW} kW<br>
<b>DC 입력전압:</b> {inv.DC입력전압} V<br>
<b>MPPT 범위:</b> {inv.MPPT범위} V<br>
<b>MPPT 수:</b> {inv.MPPT수}<br>
<b>효율:</b> {inv.효율_pct}%<br>
<b>무게:</b> {inv.무게_kg} kg
</div>""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("**단선결선도**")

    # matplotlib 이미지 우선, 실패 시 텍스트 폴백
    try:
        img_bytes = ed.generate_diagram_image(
            string_design=string_result,
            인버터품명=inv.품명,
            인버터용량_kW=inv.용량_kW,
            단상삼상=inv.단상삼상,
            인버터수=인버터수량,
            모듈품명=ms0.품명,
            모듈용량_W=ms0.용량_W,
        )
        st.image(img_bytes, use_container_width=True)
        show_text = st.checkbox("텍스트 도면도 보기", value=False)
    except Exception as img_err:
        st.warning(f"이미지 생성 실패 ({img_err}) — 텍스트 모드로 표시합니다.")
        show_text = True

    diag = ed.generate_diagram(
        string_design=string_result,
        인버터품명=inv.품명,
        인버터용량_kW=inv.용량_kW,
        단상삼상=inv.단상삼상,
        인버터수=인버터수량,
        모듈품명=ms0.품명,
        모듈용량_W=ms0.용량_W,
    )
    if show_text:
        st.code(diag.텍스트도면, language=None)

    # 다운로드
    st.download_button("📥 단선결선도 텍스트 다운로드", diag.텍스트도면,
                       f"단선결선도_{현장명 or '현장'}.txt", mime="text/plain")


# ──────────────────────────────────────────────────────────
# ──────────────────────────────────────────────────────────
# 탭: 배치도
# ──────────────────────────────────────────────────────────
with tabs[TAB_OFF + 3]:
    st.markdown('<div class="sec">배치도 (카카오 위성지도 오버레이)</div>', unsafe_allow_html=True)

    # ── API 키 입력 ──────────────────────────────────────────
    with st.expander("카카오 REST API 키 설정", expanded="kakao_key" not in st.session_state):
        st.markdown(
            "**API 키 발급:** [카카오 개발자 콘솔](https://developers.kakao.com) → 앱 생성 → REST API 키 복사"
        )
        kakao_key_input = st.text_input(
            "카카오 REST API 키",
            value=st.session_state.get("kakao_key", ""),
            type="password",
            placeholder="xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx",
        )
        if st.button("API 키 저장"):
            st.session_state["kakao_key"] = kakao_key_input
            st.success("API 키가 저장되었습니다.")

    _DEFAULT_KAKAO_KEY = st.secrets.get("kakao_api_key", "")
    if "kakao_key" not in st.session_state:
        st.session_state["kakao_key"] = _DEFAULT_KAKAO_KEY
    kakao_key = st.session_state.get("kakao_key", _DEFAULT_KAKAO_KEY)

    # ── 주소 검색 ────────────────────────────────────────────
    st.markdown("#### 1. 현장 위치 설정")
    col_addr, col_btn = st.columns([4, 1])
    with col_addr:
        layout_addr = st.text_input(
            "현장 주소",
            value=주소,
            key="layout_addr",
            placeholder="예) 전남 나주시 금천면 ..."
        )
    with col_btn:
        st.write("")
        st.write("")
        search_clicked = st.button("주소 검색")

    # 좌표 세션 저장
    if search_clicked:
        if not kakao_key:
            st.error("카카오 REST API 키를 먼저 입력하세요.")
        elif not layout_addr:
            st.warning("주소를 입력하세요.")
        else:
            try:
                km = KakaoMap(kakao_key)
                coords = km.geocode(layout_addr)
                if coords:
                    st.session_state["layout_lat"] = coords[0]
                    st.session_state["layout_lng"] = coords[1]
                    st.success(f"좌표 확인: 위도 {coords[0]:.6f}, 경도 {coords[1]:.6f}")
                else:
                    st.error("주소를 찾을 수 없습니다. 더 구체적인 주소를 입력하세요.")
            except Exception as e:
                st.error(str(e))

    layout_lat = st.session_state.get("layout_lat", 37.5)
    layout_lng = st.session_state.get("layout_lng", 127.0)

    c_lat, c_lng, c_level = st.columns(3)
    with c_lat:
        layout_lat = st.number_input("위도", value=layout_lat, format="%.6f", key="lat_in")
    with c_lng:
        layout_lng = st.number_input("경도", value=layout_lng, format="%.6f", key="lng_in")
    with c_level:
        map_level = st.selectbox(
            "지도 레벨 (숫자 클수록 넓은 범위)",
            options=list(range(1, 13)),
            index=1,   # level 2 기본값 (~1.2m/px, 태양광 현장 적합)
            format_func=lambda v: f"레벨 {v}  (~{[0.6,1.2,2.4,4.8,9.6,19,38,76,153,306,611,1222][v-1]:.0f}m/px)",
        )

    # ── 배열 위치 입력 ────────────────────────────────────────
    st.markdown("#### 2. 배열 위치 설정")
    st.caption("기준점(현장 주소 좌표)에서 각 배열까지의 거리를 입력합니다. 동쪽/북쪽이 (+)")

    drawer = LayoutDrawer()
    array_positions = []

    for i, res in enumerate(all_results):
        폭 = round(res.구조물가로_m, 2)
        길이 = round(res.구조물세로_m, 2)
        방위각 = array_configs[i].설치방향각도 if i < len(array_configs) else 180.0

        with st.expander(f"배열 {i+1}  ({폭}m × {길이}m, 방위각 {방위각}°)", expanded=(i == 0)):
            cc1, cc2, cc3 = st.columns(3)
            with cc1:
                x_m = st.number_input(
                    "동(+)/서(-) 거리 (m)", value=float(i * (폭 + 2)),
                    key=f"layout_x_{i}", step=0.5
                )
            with cc2:
                y_m = st.number_input(
                    "북(+)/남(-) 거리 (m)", value=0.0,
                    key=f"layout_y_{i}", step=0.5
                )
            with cc3:
                az = st.number_input(
                    "방위각 (°)", value=float(방위각),
                    min_value=0.0, max_value=360.0,
                    key=f"layout_az_{i}", step=1.0,
                    help="0=북, 90=동, 180=남(일반), 270=서"
                )

        array_positions.append(ArrayPosition(
            번호=i + 1, x_m=x_m, y_m=y_m,
            폭_m=폭, 길이_m=길이, 방위각_deg=az,
        ))

    # ── 배치도 생성 ───────────────────────────────────────────
    st.markdown("#### 3. 배치도 생성")
    map_mode = st.radio(
        "배경",
        ["위성지도 (카카오)", "도면만 (지도 없음)"],
        horizontal=True,
    )

    if st.button("배치도 생성", type="primary"):
        with st.spinner("배치도를 생성하는 중..."):
            try:
                if map_mode == "위성지도 (카카오)":
                    if not kakao_key:
                        st.error("카카오 API 키가 필요합니다.")
                    else:
                        km = KakaoMap(kakao_key)
                        base_img = km.fetch_map_image(
                            layout_lat, layout_lng,
                            level=map_level,
                            width=900, height=650,
                        )
                        result_img = drawer.draw_on_map(
                            base_img, array_positions,
                            level=map_level, lat=layout_lat,
                        )
                        st.session_state["layout_img"] = drawer.to_png_bytes(result_img)
                else:
                    result_img = drawer.draw_schematic(array_positions)
                    st.session_state["layout_img"] = drawer.to_png_bytes(result_img)

                st.success("배치도 생성 완료!")
            except Exception as e:
                st.error(f"배치도 생성 실패: {e}")

    if "layout_img" in st.session_state:
        st.image(st.session_state["layout_img"], use_container_width=True)
        st.download_button(
            "📥 배치도 PNG 다운로드",
            st.session_state["layout_img"],
            file_name=f"배치도_{현장명 or '현장'}_{설계일}.png",
            mime="image/png",
        )


# ──────────────────────────────────────────────────────────
# 탭: 지역 일사량
# ──────────────────────────────────────────────────────────
with tabs[TAB_OFF + 4]:
    st.markdown('<div class="sec">지역별 일조시간 / 일사량</div>', unsafe_allow_html=True)

    if 선택지역 and sun_data:
        c1, c2, c3 = st.columns(3)
        c1.metric("지역", 선택지역)
        c2.metric("일평균 일조시간", f"{sun_data['일평균_h']:.2f} h/일")
        c3.metric("연간 일조시간", f"{sun_data['월별_h'][0]:,.0f} h/년" if sun_data['월별_h'][0] > 100 else "")

        # 월별 일조시간 차트
        월별 = sun_data["월별_h"]
        if 월별[0] > 100:  # 연간 총계가 첫번째인 경우 제외
            월별_h = 월별[1:]
        else:
            월별_h = 월별

        if len(월별_h) == 12:
            df_monthly = pd.DataFrame({
                "월": [f"{i+1}월" for i in range(12)],
                "일조시간(h)": 월별_h,
            })
            st.bar_chart(df_monthly.set_index("월"))

        # 일사량
        rad_data = get_irradiance_kwh(선택지역)
        if rad_data:
            st.markdown(f"**일사량 (kWh/m²/day):** 연평균 {rad_data['연평균_kWh']:.2f}")

        # 발전량 계산 반영
        st.markdown("---")
        st.markdown("**현재 선택 지역 기준 발전량 추정**")
        gen_local = calc.estimate_generation(총용량_kW, sun_data["일평균_h"], pr)
        c1, c2, c3 = st.columns(3)
        c1.metric("일 발전량",  f"{gen_local['일발전량_kWh']:,.1f} kWh")
        c2.metric("월 발전량",  f"{gen_local['월발전량_kWh']:,.0f} kWh")
        c3.metric("연간 발전량", f"{gen_local['연발전량_kWh']:,.0f} kWh")
    else:
        st.info("왼쪽 사이드바에서 지역을 검색·선택해주세요.")

    # 전국 지역 목록
    with st.expander("전국 지역 목록 보기"):
        all_regions = sorted(SUNSHINE_DATA.keys())
        df_regions = pd.DataFrame([{
            "지역명": r,
            "일평균일조(h)": SUNSHINE_DATA[r]["일평균_h"],
        } for r in all_regions])
        st.dataframe(df_regions, use_container_width=True, hide_index=True)


# ──────────────────────────────────────────────────────────
# ──────────────────────────────────────────────────────────
# 탭: 수익 분석
# ──────────────────────────────────────────────────────────
with tabs[TAB_OFF + 5]:
    st.markdown('<div class="sec">20년 수익 분석</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**투자비**")
        총투자비 = st.number_input("총 투자비 (원)", 0, 10_000_000_000,
                                  int(총용량_kW * 1_500_000), 1_000_000,
                                  help="kW당 약 150만원 기준")
        자기자본 = st.number_input("자기 자본 (원)", 0, 10_000_000_000, int(총투자비), 1_000_000)
        타인자본 = 총투자비 - 자기자본
        st.caption(f"대출금: {fmt_won(타인자본)}")
        대출이자율 = st.number_input("대출 이자율 (%)", 0.0, 20.0, 3.5, 0.1)
        거치기간 = st.number_input("거치 기간 (년)", 0, 10, 0, help="거치기간 중 이자만 납부, 이후 원리금 균등상환")
        대출기간 = st.number_input("상환 기간 (년)", 1, 30, 15)

    with c2:
        st.markdown("**효율 설정**")
        환경효율 = st.number_input("환경 효율 (%)", 50.0, 100.0,
                                  round(ms0.효율_pct * pr * 100 / ms0.효율_pct * 95.78 / 100, 2)
                                  if ms0.효율_pct else 95.78, 0.01)
        모듈1년차 = st.number_input("1년차 모듈 효율 (%)", 80.0, 100.0, 98.5, 0.1)
        모듈감소율 = st.number_input("연간 효율 감소율 (%)", 0.1, 2.0,
                                    ms0.연간효율감소_pct, 0.01)
        분석기간 = st.selectbox("분석 기간", [20, 25, 30], index=0)

    with c3:
        st.markdown("**운영 비용**")
        연간운영비 = st.number_input("연간 운영비 (원)", 0, 100_000_000, 1_500_000, 100_000)
        연간보험비 = st.number_input("연간 보험비 (원)", 0, 50_000_000, 0, 100_000)

    st.markdown("---")

    # 수익 분석 실행
    rev_inp = RevenueInput(
        설비용량_kW=총용량_kW,
        일평균발전시간_h=일평균일조시간,
        환경효율_pct=환경효율,
        인버터효율_pct=inverter_spec.효율_pct,
        모듈1년차효율_pct=모듈1년차,
        모듈연간감소율_pct=모듈감소율,
        가중치=REC가중치,
        SMP_원kWh=SMP단가,
        REC_원kWh=REC단가,
        판매구분=판매구분,
        총투자비_원=총투자비,
        자기자본_원=자기자본,
        타인자본_원=타인자본,
        대출이자율_pct=대출이자율,
        거치기간_년=거치기간,
        대출기간_년=대출기간,
        연간운영비_원=연간운영비,
        연간보험비_원=연간보험비,
        분석기간_년=분석기간,
    )

    rev_calc = RevenueCalculator()
    rev_result = rev_calc.analyze(rev_inp)

    # 핵심 지표
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("총 수입",       fmt_won(rev_result.총수입_원))
    c2.metric("총 순수입",     fmt_won(rev_result.총순수입_원))
    c3.metric("내부수익률(IRR)", f"{rev_result.내부수익률_pct:.2f}%")
    c4.metric("투자 회수 기간",  f"{rev_result.투자회수기간_년:.1f} 년")

    kW당투자 = round(총투자비 / 총용량_kW) if 총용량_kW > 0 else 0
    st.caption(f"kW당 투자비: {kW당투자:,.0f} 원/kW | 설비 용량: {총용량_kW:.2f} kW")

    # 연도별 표
    st.markdown('<div class="sec">연도별 수익 분석표</div>', unsafe_allow_html=True)
    df_rev = pd.DataFrame([{
        "년차":         r.년차,
        "발전량(kWh)":  f"{r.발전량_kWh:,.1f}",
        "연수입(원)":   f"{r.연수입_원:,.0f}",
        "운영비(원)":   f"{r.연운영비_원:,.0f}",
        "대출원리금(원)": f"{r.대출원리금_원:,.0f}",
        "순수입(원)":   f"{r.순수입_원:,.0f}",
        "누적순수입(원)": f"{r.누적순수입_원:,.0f}",
    } for r in rev_result.연도별])
    st.dataframe(df_rev, use_container_width=True, hide_index=True, height=400)

    # 누적 수익 차트
    st.markdown('<div class="sec">누적 순수입 추이</div>', unsafe_allow_html=True)
    df_chart = pd.DataFrame({
        "년차": [r.년차 for r in rev_result.연도별],
        "누적순수입": [r.누적순수입_원 / 1e6 for r in rev_result.연도별],
        "투자비(백만원)": [총투자비 / 1e6] * len(rev_result.연도별),
    }).set_index("년차")
    st.line_chart(df_chart)

    # Excel 다운로드 (수익분석 포함)
    excel_bytes = export_to_excel(all_results, project_info, gen_info)
    if excel_bytes:
        st.download_button("📥 전체 결과 Excel 다운로드", excel_bytes,
                           f"태양광가설계_{현장명 or '현장'}_{설계일}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
