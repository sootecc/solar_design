"""
배치도 생성기
- 카카오 지도 API: 주소→좌표(Geocoding), 위성지도 이미지
- PIL 기반 배열 사각형 오버레이
VBA 원본: Module11.bas 배열위치/방위각 로직
"""
import io
import math
import os
from dataclasses import dataclass, field
from typing import Optional, Tuple, List

# 한글 지원 TrueType 폰트 경로 (Windows 우선)
_KR_FONT_CANDIDATES = [
    r"C:\Windows\Fonts\malgun.ttf",
    r"C:\Windows\Fonts\gulim.ttc",
    r"C:\Windows\Fonts\batang.ttc",
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
]

def _get_pil_font(size: int = 12):
    from PIL import ImageFont
    for path in _KR_FONT_CANDIDATES:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size)
            except Exception:
                continue
    return ImageFont.load_default()

KAKAO_GEOCODE_URL = "https://dapi.kakao.com/v2/local/search/address.json"

# 카카오 레벨 → Web Mercator zoom 변환 (UI 표시용 m/pixel 포함)
# zoom = 19 - level  (level 1=가장 확대, level 12=가장 축소)
KAKAO_MPP = {
    1: 0.60,  2: 1.19,  3: 2.39,  4: 4.77,
    5: 9.55,  6: 19.1,  7: 38.2,  8: 76.4,
    9: 152.8, 10: 305.6, 11: 611.2, 12: 1222.4,
}

def _level_to_zoom(level: int) -> int:
    """카카오 레벨 → Web Mercator 줌"""
    return max(1, min(19, 19 - level))

ARRAY_COLORS = [
    "#2196F3", "#4CAF50", "#FF5722", "#9C27B0", "#FF9800",
    "#00BCD4", "#E91E63", "#795548", "#607D8B", "#3F51B5",
]


@dataclass
class ArrayPosition:
    """배열 하나의 위치·치수 정보"""
    번호: int
    x_m: float = 0.0        # 기준점 동쪽(+) / 서쪽(-) 거리 (m)
    y_m: float = 0.0        # 기준점 북쪽(+) / 남쪽(-) 거리 (m)
    폭_m: float = 0.0       # 구조물 폭 (m)
    길이_m: float = 0.0     # 구조물 길이_수평투영 (m)
    방위각_deg: float = 180.0  # 패널 전면이 향하는 방위 (0=북, 90=동, 180=남)


class KakaoMap:
    """카카오 REST API 연동"""

    def __init__(self, api_key: str):
        self.headers = {"Authorization": f"KakaoAK {api_key}"}

    def geocode(self, address: str) -> Optional[Tuple[float, float]]:
        """주소 → (위도, 경도)"""
        try:
            import requests
            r = requests.get(
                KAKAO_GEOCODE_URL,
                headers=self.headers,
                params={"query": address},
                timeout=6,
            )
            r.raise_for_status()
            docs = r.json().get("documents", [])
            if not docs:
                return None
            d = docs[0]
            return float(d["y"]), float(d["x"])
        except Exception as e:
            raise RuntimeError(f"주소 검색 실패: {e}")

    def fetch_map_image(
        self,
        lat: float,
        lng: float,
        level: int = 4,
        width: int = 800,
        height: int = 600,
        maptype: str = "skyview",
    ):
        """
        위성(ESRI) 또는 도로(OSM) 타일을 합성해 PIL 이미지 반환.
        카카오 REST API는 geocoding에만 사용; 정적 지도는 공개 타일 서버 활용.
        """
        try:
            import requests
            from PIL import Image

            zoom = _level_to_zoom(level)
            lat_rad = math.radians(lat)
            n = 2 ** zoom
            x_f = (lng + 180.0) / 360.0 * n
            y_f = (1.0 - math.log(math.tan(lat_rad) + 1.0 / math.cos(lat_rad)) / math.pi) / 2.0 * n

            cx, cy = int(x_f), int(y_f)
            ox, oy = x_f - cx, y_f - cy  # 타일 내 소수점 오프셋

            TILE = 256
            nx = math.ceil(width  / TILE) + 2
            ny = math.ceil(height / TILE) + 2
            canvas = Image.new("RGB", (nx * TILE, ny * TILE), (100, 100, 100))

            tx0 = cx - nx // 2
            ty0 = cy - ny // 2

            sess = requests.Session()
            sess.headers["User-Agent"] = "SolarDesignApp/1.0"

            for ti in range(nx):
                for tj in range(ny):
                    tx = tx0 + ti
                    ty = ty0 + tj
                    if tx < 0 or ty < 0 or tx >= n or ty >= n:
                        continue
                    if maptype == "skyview":
                        # ESRI World Imagery (무료 위성)
                        url = (
                            f"https://server.arcgisonline.com/ArcGIS/rest/services"
                            f"/World_Imagery/MapServer/tile/{zoom}/{ty}/{tx}"
                        )
                    else:
                        # OpenStreetMap 도로지도
                        url = f"https://tile.openstreetmap.org/{zoom}/{tx}/{ty}.png"
                    try:
                        r = sess.get(url, timeout=8)
                        if r.status_code == 200:
                            tile = Image.open(io.BytesIO(r.content)).convert("RGB")
                            canvas.paste(tile, (ti * TILE, tj * TILE))
                    except Exception:
                        pass

            # 중심 픽셀 기준 크롭
            cpx = (nx // 2) * TILE + int(ox * TILE)
            cpy = (ny // 2) * TILE + int(oy * TILE)
            l, t = cpx - width // 2, cpy - height // 2
            return canvas.crop((l, t, l + width, t + height)).convert("RGBA")

        except Exception as e:
            raise RuntimeError(f"지도 이미지 로드 실패: {e}")


class LayoutDrawer:
    """배치도 드로어"""

    def _mpp(self, level: int, lat: float) -> float:
        """m/pixel — Web Mercator 표준 공식 (타일 실제 스케일과 일치)"""
        zoom = _level_to_zoom(level)
        return 156543.03392 * math.cos(math.radians(lat)) / (2 ** zoom)

    def _rotated_rect(
        self,
        cx: float, cy: float,
        w_px: float, l_px: float,
        angle_rad: float,
    ) -> List[Tuple[float, float]]:
        """회전 사각형 꼭짓점 반환 (이미지 좌표계)"""
        hw, hl = w_px / 2, l_px / 2
        pts_local = [(-hw, -hl), (hw, -hl), (hw, hl), (-hw, hl)]
        result = []
        for lx, ly in pts_local:
            rx = lx * math.cos(angle_rad) - ly * math.sin(angle_rad)
            ry = lx * math.sin(angle_rad) + ly * math.cos(angle_rad)
            result.append((cx + rx, cy + ry))
        return result

    def draw_on_map(
        self,
        base_img,
        arrays: List[ArrayPosition],
        level: int,
        lat: float,
        alpha: int = 160,
    ):
        """지도 이미지 위에 배열 오버레이 (두꺼운 테두리 + 모듈행 선 + 텍스트 그림자)"""
        from PIL import Image, ImageDraw, ImageColor

        img = base_img.copy().convert("RGBA")
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)

        cx0, cy0 = img.width / 2, img.height / 2
        mpp = self._mpp(level, lat)
        px_per_m = 1.0 / mpp

        # 배열이 너무 작으면 최소 크기 보장
        MIN_PX = 8

        font_size = max(11, min(18, int(30 * px_per_m)))
        font = _get_pil_font(font_size)

        for i, ap in enumerate(arrays):
            color_hex = ARRAY_COLORS[i % len(ARRAY_COLORS)]
            rgb = ImageColor.getrgb(color_hex)

            cx = cx0 + ap.x_m * px_per_m
            cy = cy0 - ap.y_m * px_per_m

            w_px = max(MIN_PX, ap.폭_m * px_per_m)
            l_px = max(MIN_PX, ap.길이_m * px_per_m)

            angle_rad = math.radians(ap.방위각_deg)
            corners = self._rotated_rect(cx, cy, w_px, l_px, angle_rad)
            ipts = [(int(x), int(y)) for x, y in corners]

            # 반투명 채우기
            draw.polygon(ipts, fill=rgb + (alpha,))

            # 두꺼운 테두리 (outline 3px)
            border = [(int(x), int(y)) for x, y in corners + [corners[0]]]
            draw.line(border, fill=rgb + (255,), width=3)

            # 모듈 행 구분선 (방위각 방향 수직)
            if ap.길이_m > 0 and l_px > 12:
                n_rows = max(1, round(ap.길이_m / 1.0))  # 약 1m 간격
                step = l_px / n_rows
                perp_x = math.cos(angle_rad)
                perp_y = math.sin(angle_rad)
                along_x = -math.sin(angle_rad)
                along_y = math.cos(angle_rad)
                for r in range(1, n_rows):
                    row_cx = cx + perp_x * (r * step - l_px / 2)
                    row_cy = cy + perp_y * (r * step - l_px / 2)
                    p1 = (int(row_cx + along_x * w_px / 2),
                          int(row_cy + along_y * w_px / 2))
                    p2 = (int(row_cx - along_x * w_px / 2),
                          int(row_cy - along_y * w_px / 2))
                    draw.line([p1, p2], fill=rgb + (200,), width=1)

            # 텍스트 (그림자 효과)
            label = f"배열{ap.번호}\n{ap.폭_m:.1f}×{ap.길이_m:.1f}m"
            tx, ty = int(cx) - font_size * 2, int(cy) - font_size
            draw.text((tx + 1, ty + 1), label, fill=(0, 0, 0, 200), font=font)
            draw.text((tx,     ty    ), label, fill=(255, 255, 255, 255), font=font)

        return Image.alpha_composite(img, overlay)

    def draw_schematic(
        self,
        arrays: List[ArrayPosition],
        margin_m: float = 15.0,
        canvas_w: int = 900,
    ):
        """지도 없이 배치 평면도 생성 (흰 배경, 치수선 포함)"""
        from PIL import Image, ImageDraw, ImageColor

        if not arrays:
            img = Image.new("RGB", (500, 300), (255, 255, 255))
            ImageDraw.Draw(img).text((10, 10), "배열 위치를 입력하세요", fill=(0, 0, 0))
            return img

        all_x = [ap.x_m for ap in arrays]
        all_y = [ap.y_m for ap in arrays]
        max_half = max(
            max(ap.폭_m, ap.길이_m) for ap in arrays
        ) / 2

        min_x = min(all_x) - max_half - margin_m
        max_x = max(all_x) + max_half + margin_m
        min_y = min(all_y) - max_half - margin_m
        max_y = max(all_y) + max_half + margin_m

        span_x = max(max_x - min_x, 1.0)
        span_y = max(max_y - min_y, 1.0)

        pad = 60
        draw_w = canvas_w - 2 * pad
        scale = draw_w / span_x
        canvas_h = int(span_y * scale) + 2 * pad + 80

        img = Image.new("RGB", (canvas_w, canvas_h), (250, 250, 250))
        draw = ImageDraw.Draw(img, "RGBA")

        # 그리드 (10m 간격)
        grid_step_m = 10
        grid_px = grid_step_m * scale
        gx0 = pad + (0 - min_x) * scale
        for gi in range(-200, 200):
            gx = int(gx0 + gi * grid_px)
            if pad <= gx <= canvas_w - pad:
                draw.line([(gx, pad), (gx, canvas_h - pad - 60)],
                          fill=(220, 220, 220, 255), width=1)
            gy_base = canvas_h - pad - 60 - (0 - min_y) * scale
            gy = int(gy_base - gi * grid_px)
            if pad <= gy <= canvas_h - pad - 60:
                draw.line([(pad, gy), (canvas_w - pad, gy)],
                          fill=(220, 220, 220, 255), width=1)

        def to_px(x_m, y_m):
            px = pad + (x_m - min_x) * scale
            py = canvas_h - pad - 60 - (y_m - min_y) * scale
            return px, py

        font_l = _get_pil_font(14)
        font_s = _get_pil_font(11)

        # 배열 그리기
        for i, ap in enumerate(arrays):
            color_hex = ARRAY_COLORS[i % len(ARRAY_COLORS)]
            rgb = ImageColor.getrgb(color_hex)

            cx, cy = to_px(ap.x_m, ap.y_m)
            w_px = ap.폭_m * scale
            l_px = ap.길이_m * scale
            angle_rad = math.radians(ap.방위각_deg)

            corners = self._rotated_rect(cx, cy, w_px, l_px, angle_rad)
            ipts = [(int(x), int(y)) for x, y in corners]

            draw.polygon(ipts, fill=rgb + (160,), outline=rgb + (255,))

            draw.text(
                (int(cx) - 20, int(cy) - 20),
                f"배열{ap.번호}",
                fill=(0, 0, 0, 255),
                font=font_l,
            )
            draw.text(
                (int(cx) - 28, int(cy) - 4),
                f"{ap.폭_m:.1f}×{ap.길이_m:.1f}m",
                fill=(50, 50, 50, 255),
                font=font_s,
            )

        # 스케일바
        scale_m = 10
        scale_px = int(scale_m * scale)
        sx, sy = pad, canvas_h - 50
        draw.line([(sx, sy), (sx + scale_px, sy)], fill=(0, 0, 0, 255), width=3)
        draw.line([(sx, sy - 5), (sx, sy + 5)], fill=(0, 0, 0, 255), width=2)
        draw.line([(sx + scale_px, sy - 5), (sx + scale_px, sy + 5)],
                  fill=(0, 0, 0, 255), width=2)
        draw.text((sx + scale_px // 2 - 15, sy + 6), f"{scale_m}m",
                  fill=(0, 0, 0), font=font_s)

        # 방위 (북 화살표)
        nx, ny = canvas_w - 50, pad + 20
        draw.line([(nx, ny + 20), (nx, ny)], fill=(0, 0, 100, 255), width=2)
        draw.polygon([(nx, ny), (nx - 6, ny + 10), (nx + 6, ny + 10)],
                     fill=(0, 0, 180, 255))
        draw.text((nx - 4, ny + 22), "N", fill=(0, 0, 100), font=font_s)

        return img

    def to_png_bytes(self, img) -> bytes:
        buf = io.BytesIO()
        img.convert("RGB").save(buf, format="PNG", dpi=(150, 150))
        buf.seek(0)
        return buf.read()
