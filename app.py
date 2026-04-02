"""
送迎ルート自動作成システム - Streamlit Webアプリ
=================================================
放課後等デイサービス / 就労継続支援A型・B型 対応

動作環境:
  - Streamlit Community Cloud（無料枠）
  - OR-Tools 未インストール時はグリーディアルゴリズムで自動フォールバック
  - Google Maps API 不使用（ハーバーサイン距離で推定）
"""

from __future__ import annotations

import io
import math
from dataclasses import dataclass, field
from typing import Optional
from enum import Enum

import pandas as pd
import streamlit as st

# ---- オプションライブラリ（なくても動作する）----
try:
    import folium
    from streamlit_folium import st_folium
    FOLIUM_AVAILABLE = True
except ImportError:
    FOLIUM_AVAILABLE = False

try:
    from ortools.constraint_solver import routing_enums_pb2, pywrapcp
    ORTOOLS_AVAILABLE = True
except ImportError:
    ORTOOLS_AVAILABLE = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ==============================================================
# ページ設定（必ず先頭）
# ==============================================================
st.set_page_config(
    page_title="送迎ルート最適化システム",
    page_icon="🚌",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==============================================================
# カスタム CSS
# ==============================================================
st.markdown("""
<style>
  /* フォント */
  html, body, [class*="css"] { font-family: 'Noto Sans JP', 'メイリオ', sans-serif; }

  /* メインヘッダー */
  .main-header {
    background: linear-gradient(135deg, #1C2833 0%, #2C3E50 100%);
    border-radius: 12px;
    padding: 24px 28px;
    margin-bottom: 24px;
    color: white;
  }
  .main-header h1 { font-size: 26px; margin: 0; font-weight: 800; letter-spacing: 0.04em; }
  .main-header p  { font-size: 13px; margin: 6px 0 0; opacity: 0.7; }

  /* ステップバッジ */
  .step-badge {
    display: inline-block;
    background: #2D5A3D;
    color: white;
    font-size: 11px;
    font-weight: 700;
    padding: 3px 10px;
    border-radius: 20px;
    margin-bottom: 8px;
    letter-spacing: 0.05em;
  }

  /* 結果カード */
  .result-card {
    background: #F0F4F0;
    border-left: 4px solid #2D5A3D;
    border-radius: 0 8px 8px 0;
    padding: 14px 18px;
    margin-bottom: 12px;
  }
  .result-card h3 { font-size: 15px; margin: 0 0 4px; font-weight: 700; }
  .result-card p  { font-size: 12px; color: #555; margin: 0; }

  /* 警告 */
  .constraint-ok   { color: #27AE60; font-weight: 700; }
  .constraint-fail { color: #E74C3C; font-weight: 700; }

  /* メトリクス */
  .metric-box {
    background: white;
    border: 1px solid #DDD;
    border-radius: 8px;
    padding: 14px;
    text-align: center;
  }
  .metric-box .val  { font-size: 28px; font-weight: 800; color: #2C3E50; }
  .metric-box .label{ font-size: 12px; color: #888; margin-top: 2px; }

  /* テーブルスタイル調整 */
  .stDataFrame { border-radius: 8px; overflow: hidden; }

  /* ボタン */
  .stButton > button {
    border-radius: 8px !important;
    font-weight: 600 !important;
    transition: all 0.15s !important;
  }
</style>
""", unsafe_allow_html=True)


# ==============================================================
# コアアルゴリズム（前回のコードをインライン統合）
# ==============================================================

class ServiceType(Enum):
    HOUKAGO_DEI = "放課後等デイサービス"
    A_TYPE      = "就労継続支援A型"
    B_TYPE      = "就労継続支援B型"


class TripType(Enum):
    PICKUP  = "迎え"
    DROPOFF = "送り"


@dataclass
class User:
    user_id:       str
    name:          str
    address:       str
    lat:           float
    lng:           float
    service_type:  ServiceType
    wheelchair:    bool = False
    incompatible:  list[str] = field(default_factory=list)
    pickup_latest: Optional[int] = 540   # 分（デフォルト9:00）
    dropoff_target:Optional[int] = 1050  # 分（デフォルト17:30）


@dataclass
class Vehicle:
    vehicle_id:   str
    name:         str
    vehicle_type: str
    capacity:     int
    wheelchair_ok:bool = False
    depot_lat:    float = 36.6953
    depot_lng:    float = 137.2113


@dataclass
class Staff:
    staff_id:  str
    name:      str
    can_drive: bool = True


@dataclass
class AssignedRoute:
    vehicle:   Vehicle
    driver:    Optional[Staff]
    trip_type: TripType
    stops:     list[dict]
    total_min: int


class DistanceMatrixBuilder:
    """ハーバーサイン距離から所要時間行列を生成"""

    def build(self, locations: list[tuple[float, float]]) -> list[list[int]]:
        n = len(locations)
        return [
            [self._haversine_sec(locations[i], locations[j]) for j in range(n)]
            for i in range(n)
        ]

    @staticmethod
    def _haversine_sec(p1, p2, avg_speed_kmh: float = 30.0) -> int:
        R = 6371.0
        lat1, lng1 = math.radians(p1[0]), math.radians(p1[1])
        lat2, lng2 = math.radians(p2[0]), math.radians(p2[1])
        dlat, dlng = lat2 - lat1, lng2 - lng1
        a = math.sin(dlat/2)**2 + math.cos(lat1)*math.cos(lat2)*math.sin(dlng/2)**2
        dist_km = 2 * R * math.asin(math.sqrt(a))
        return int(dist_km / avg_speed_kmh * 3600)


class ConstraintChecker:
    @staticmethod
    def validate(users, vehicles, staff) -> list[str]:
        errors = []
        drivable = [s for s in staff if s.can_drive]
        if not drivable:
            errors.append("❌ 運転可能なスタッフが0人です")
        wc_users = [u for u in users if u.wheelchair]
        wc_veh   = [v for v in vehicles if v.wheelchair_ok]
        if wc_users and not wc_veh:
            errors.append(f"❌ 車椅子利用者 {[u.name for u in wc_users]} がいますが車椅子対応車両がありません")
        total_cap = sum(v.capacity for v in vehicles)
        if len(users) > total_cap:
            errors.append(f"❌ 利用者数({len(users)}名) > 全車両定員合計({total_cap}名)")
        return errors

    @staticmethod
    def get_forbidden_pairs(users) -> set[tuple[str, str]]:
        pairs = set()
        for u in users:
            for iid in u.incompatible:
                pairs.add(tuple(sorted([u.user_id, iid])))
        return pairs


class TransportVRPSolver:
    """VRPソルバー（OR-Tools / グリーディフォールバック）"""

    def __init__(self, users, vehicles, staff, distance_matrix,
                 trip_type=TripType.PICKUP,
                 depot_arrival_limit_min=540,
                 start_time_min=480):
        self.users    = users
        self.vehicles = vehicles
        self.staff    = [s for s in staff if s.can_drive]
        self.matrix   = distance_matrix
        self.trip_type = trip_type
        self.depot_arrival_limit = depot_arrival_limit_min * 60
        self.start_time = start_time_min * 60
        checker = ConstraintChecker()
        self.forbidden_pairs = checker.get_forbidden_pairs(users)

    def solve(self) -> list[AssignedRoute]:
        if ORTOOLS_AVAILABLE:
            result = self._solve_with_ortools()
            if result:
                return result
        return self._greedy_fallback()

    def _solve_with_ortools(self) -> list[AssignedRoute]:
        n_users    = len(self.users)
        n_vehicles = len(self.vehicles)
        n_nodes    = n_users + 1

        manager = pywrapcp.RoutingIndexManager(n_nodes, n_vehicles, 0)
        routing = pywrapcp.RoutingModel(manager)

        def time_callback(fi, ti):
            return self.matrix[manager.IndexToNode(fi)][manager.IndexToNode(ti)]

        cb_idx = routing.RegisterTransitCallback(time_callback)
        routing.SetArcCostEvaluatorOfAllVehicles(cb_idx)

        routing.AddDimension(cb_idx, 300, self.depot_arrival_limit, True, "Time")
        time_dim = routing.GetDimensionOrDie("Time")
        if self.trip_type == TripType.PICKUP:
            time_dim.CumulVar(manager.NodeToIndex(0)).SetMax(self.depot_arrival_limit)

        def demand_cb(fi):
            return 0 if manager.IndexToNode(fi) == 0 else 1
        dcb = routing.RegisterUnaryTransitCallback(demand_cb)
        routing.AddDimensionWithVehicleCapacity(
            dcb, 0, [v.capacity for v in self.vehicles], True, "Cap"
        )

        # 車椅子制約
        for i, u in enumerate(self.users):
            if u.wheelchair:
                idx = manager.NodeToIndex(i + 1)
                for vi, v in enumerate(self.vehicles):
                    if not v.wheelchair_ok:
                        routing.VehicleVar(idx).RemoveValue(vi)

        # 同乗不可制約
        for uid1, uid2 in self.forbidden_pairs:
            i1 = next((i+1 for i,u in enumerate(self.users) if u.user_id==uid1), None)
            i2 = next((i+1 for i,u in enumerate(self.users) if u.user_id==uid2), None)
            if i1 and i2:
                ni1 = manager.NodeToIndex(i1)
                ni2 = manager.NodeToIndex(i2)
                routing.solver().Add(
                    routing.VehicleVar(ni1) != routing.VehicleVar(ni2)
                )

        params = pywrapcp.DefaultRoutingSearchParameters()
        params.first_solution_strategy = (
            routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC
        )
        params.local_search_metaheuristic = (
            routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
        )
        params.time_limit.FromSeconds(30)
        solution = routing.SolveWithParameters(params)
        if not solution:
            return []

        time_dim = routing.GetDimensionOrDie("Time")
        routes = []
        for vi in range(n_vehicles):
            idx   = routing.Start(vi)
            stops = []
            while not routing.IsEnd(idx):
                node = manager.IndexToNode(idx)
                if node != 0:
                    u = self.users[node - 1]
                    arr_sec = solution.Min(time_dim.CumulVar(idx))
                    arr_min = self.start_time // 60 + arr_sec // 60
                    stops.append({"user": u, "arrival_min": arr_min,
                                  "address": u.address, "lat": u.lat, "lng": u.lng})
                idx = solution.Value(routing.NextVar(idx))
            if not stops:
                continue
            driver = self.staff[vi % len(self.staff)] if self.staff else None
            routes.append(AssignedRoute(
                vehicle=self.vehicles[vi], driver=driver,
                trip_type=self.trip_type, stops=stops, total_min=0
            ))
        return routes

    def _greedy_fallback(self) -> list[AssignedRoute]:
        vehicles_sorted = sorted(
            self.vehicles,
            key=lambda v: (0 if v.wheelchair_ok else 1, -v.capacity)
        )
        routes     = []
        unassigned = list(self.users)

        for vi, vehicle in enumerate(vehicles_sorted):
            if not unassigned:
                break
            driver   = self.staff[vi % len(self.staff)] if self.staff else None
            assigned = []

            # 車椅子優先
            if vehicle.wheelchair_ok:
                for u in [u for u in unassigned if u.wheelchair]:
                    if len(assigned) < vehicle.capacity:
                        assigned.append(u)
                        unassigned.remove(u)

            # 同乗不可を守りながら残り枠を埋める
            forbidden_ids = set()
            for a in assigned:
                forbidden_ids.update(a.incompatible)

            for u in list(unassigned):
                if len(assigned) >= vehicle.capacity:
                    break
                if u.wheelchair and not vehicle.wheelchair_ok:
                    continue
                if u.user_id in forbidden_ids:
                    continue
                assigned.append(u)
                unassigned.remove(u)
                forbidden_ids.update(u.incompatible)

            if not assigned:
                continue

            ordered     = self._nearest_neighbor(assigned)
            stops       = []
            cur_node    = 0
            cur_time    = self.start_time

            for u in ordered:
                uid      = self.users.index(u) + 1
                cur_time += self.matrix[cur_node][uid]
                stops.append({
                    "user": u, "arrival_min": cur_time // 60,
                    "address": u.address, "lat": u.lat, "lng": u.lng
                })
                cur_node = uid

            routes.append(AssignedRoute(
                vehicle=vehicle, driver=driver,
                trip_type=self.trip_type, stops=stops,
                total_min=(cur_time - self.start_time) // 60
            ))

        if unassigned:
            st.warning(f"⚠️ 以下の利用者が割り当て不可です: {[u.name for u in unassigned]}")

        return routes

    def _nearest_neighbor(self, users):
        if not users:
            return []
        remaining, ordered, cur = list(users), [], 0
        while remaining:
            nearest = min(remaining, key=lambda u: self.matrix[cur][self.users.index(u)+1])
            ordered.append(nearest)
            cur = self.users.index(nearest) + 1
            remaining.remove(nearest)
        return ordered


# ==============================================================
# Excel 入出力
# ==============================================================

def parse_excel_upload(uploaded_file) -> tuple[list[User], list[Vehicle], list[Staff]]:
    """
    アップロードされたExcelを読み込んでデータクラスに変換する。

    Excelシート構成:
      シート名「利用者」: 氏名, 住所, 緯度, 経度, サービス種別, 車椅子, 同乗不可(カンマ区切りID)
      シート名「車両」:   車両名, 種別, 定員, 車椅子対応, デポ緯度, デポ経度
      シート名「スタッフ」: 氏名, 運転可否
    """
    xl = pd.ExcelFile(uploaded_file)

    # ---- 利用者 ----
    df_users = xl.parse("利用者")
    service_map = {
        "放課後等デイサービス": ServiceType.HOUKAGO_DEI,
        "A型": ServiceType.A_TYPE,
        "B型": ServiceType.B_TYPE,
    }
    users = []
    for i, row in df_users.iterrows():
        incomp_raw = str(row.get("同乗不可ID", "")).strip()
        incomp     = [x.strip() for x in incomp_raw.split(",") if x.strip()] \
                     if incomp_raw not in ("", "nan") else []
        users.append(User(
            user_id      = str(row.get("ID", f"u{i+1}")),
            name         = str(row["氏名"]),
            address      = str(row.get("住所", "")),
            lat          = float(row.get("緯度", 36.695)),
            lng          = float(row.get("経度", 137.211)),
            service_type = service_map.get(str(row.get("サービス種別", "")),
                                           ServiceType.HOUKAGO_DEI),
            wheelchair   = bool(row.get("車椅子", False)),
            incompatible = incomp,
            pickup_latest  = int(row.get("到着リミット(分)", 540)),
            dropoff_target = int(row.get("送り目標(分)", 1050)),
        ))

    # ---- 車両 ----
    df_veh = xl.parse("車両")
    type_cap = {"large": 7, "normal": 4, "kei": 3}
    vehicles = []
    for i, row in df_veh.iterrows():
        vtype = str(row.get("種別コード", "normal"))
        vehicles.append(Vehicle(
            vehicle_id    = str(row.get("ID", f"v{i+1}")),
            name          = str(row["車両名"]),
            vehicle_type  = vtype,
            capacity      = int(row.get("定員", type_cap.get(vtype, 4))),
            wheelchair_ok = bool(row.get("車椅子対応", False)),
            depot_lat     = float(row.get("デポ緯度", 36.695)),
            depot_lng     = float(row.get("デポ経度", 137.211)),
        ))

    # ---- スタッフ ----
    df_staff = xl.parse("スタッフ")
    staff = []
    for i, row in df_staff.iterrows():
        staff.append(Staff(
            staff_id  = str(row.get("ID", f"s{i+1}")),
            name      = str(row["氏名"]),
            can_drive = bool(row.get("運転可否", True)),
        ))

    return users, vehicles, staff


def build_excel_output(routes: list[AssignedRoute]) -> bytes:
    """結果をExcelバイト列として返す"""
    if not OPENPYXL_AVAILABLE:
        # openpyxl がなければCSVで代替
        rows = []
        for r in routes:
            for i, s in enumerate(r.stops):
                h, m = divmod(s["arrival_min"], 60)
                rows.append({
                    "車両": r.vehicle.name,
                    "運転者": r.driver.name if r.driver else "未定",
                    "順番": i + 1,
                    "氏名": s["user"].name,
                    "住所": s["address"],
                    "到着予定": f"{h:02d}:{m:02d}",
                    "車椅子": "♿" if s["user"].wheelchair else "",
                })
        buf = io.StringIO()
        pd.DataFrame(rows).to_csv(buf, index=False, encoding="utf-8-sig")
        return buf.getvalue().encode("utf-8-sig")

    wb = Workbook()
    ws = wb.active
    ws.title = "送迎ルート"

    HDR_FILL = PatternFill("solid", fgColor="2C3E50")
    VEH_FILL = PatternFill("solid", fgColor="D6EAF8")
    WC_FILL  = PatternFill("solid", fgColor="FADBD8")
    ALT_FILL = PatternFill("solid", fgColor="F8F9FA")

    def s(style="thin", color="CCCCCC"):
        sd = Side(style=style, color=color)
        return Border(left=sd, right=sd, top=sd, bottom=sd)

    # タイトル
    trip = routes[0].trip_type.value if routes else "送迎"
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value     = f"🚌  送迎ルート表　【{trip}便】"
    c.font      = Font(bold=True, size=14, color="FFFFFF", name="メイリオ")
    c.fill      = HDR_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # ヘッダー
    headers = ["車両名", "運転担当", "順番", "利用者氏名", "住所", "サービス", "到着予定", "備考"]
    widths  = [20,       14,         6,      16,          36,     16,          10,        14]
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=10, name="メイリオ")
        c.fill      = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = s()
        ws.column_dimensions[chr(64+col)].width = w
    ws.row_dimensions[2].height = 22

    row = 3
    for ri, route in enumerate(routes):
        driver_name = route.driver.name if route.driver else "未定"
        for i, stop in enumerate(route.stops):
            h, m   = divmod(stop["arrival_min"], 60)
            wc     = "♿ 車椅子" if stop["user"].wheelchair else ""
            svc    = stop["user"].service_type.value

            fill   = WC_FILL if stop["user"].wheelchair else (
                VEH_FILL if i == 0 else (ALT_FILL if ri % 2 else None)
            )
            data = [
                route.vehicle.name if i == 0 else "",
                driver_name        if i == 0 else "",
                i + 1,
                stop["user"].name,
                stop["address"],
                svc,
                f"{h:02d}:{m:02d}",
                wc,
            ]
            for col, val in enumerate(data, 1):
                c = ws.cell(row=row, column=col, value=val)
                c.font      = Font(bold=(i == 0), size=10, name="メイリオ")
                c.alignment = Alignment(
                    horizontal="center" if col in [1, 2, 3, 7, 8] else "left",
                    vertical="center"
                )
                c.border = s()
                if fill:
                    c.fill = fill
            ws.row_dimensions[row].height = 20
            row += 1
        row += 1  # 車両間の空白行

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==============================================================
# デモデータ
# ==============================================================

def get_demo_data() -> tuple[list[User], list[Vehicle], list[Staff]]:
    """すぐ動かせるデモデータ（富山市を想定）"""
    users = [
        User("u1",  "山田 太郎",   "富山市上袋100",     36.720, 137.210, ServiceType.HOUKAGO_DEI, False, [],      540, 1050),
        User("u2",  "鈴木 花子",   "富山市堀川200",     36.695, 137.220, ServiceType.HOUKAGO_DEI, False, ["u3"], 540, 1050),
        User("u3",  "田中 一郎",   "富山市婦中300",     36.660, 137.160, ServiceType.A_TYPE,      True,  ["u2"], 540, 1050),
        User("u4",  "佐藤 愛",     "富山市大沢野400",   36.630, 137.230, ServiceType.B_TYPE,      False, [],      540, 1050),
        User("u5",  "高橋 健太",   "富山市八尾500",     36.590, 137.270, ServiceType.HOUKAGO_DEI, False, [],      540, 1050),
        User("u6",  "渡辺 さくら", "富山市上袋600",     36.725, 137.215, ServiceType.B_TYPE,      False, [],      540, 1050),
        User("u7",  "伊藤 翔",     "富山市堀川700",     36.700, 137.225, ServiceType.A_TYPE,      False, [],      540, 1050),
        User("u8",  "中村 みな",   "富山市婦中800",     36.655, 137.155, ServiceType.B_TYPE,      False, [],      540, 1050),
        User("u9",  "小林 大輝",   "富山市大沢野900",   36.625, 137.235, ServiceType.HOUKAGO_DEI, False, [],      540, 1050),
        User("u10", "加藤 りん",   "富山市八尾1000",    36.585, 137.265, ServiceType.A_TYPE,      False, [],      540, 1050),
    ]
    vehicles = [
        Vehicle("v1", "A-1号車（大型）", "large",  7, wheelchair_ok=True,  depot_lat=36.695, depot_lng=137.211),
        Vehicle("v2", "A-2号車（普通）", "normal", 4, wheelchair_ok=False, depot_lat=36.695, depot_lng=137.211),
        Vehicle("v3", "A-3号車（軽）",  "kei",    3, wheelchair_ok=False, depot_lat=36.695, depot_lng=137.211),
    ]
    staff = [
        Staff("s1", "林 誠一",   can_drive=True),
        Staff("s2", "森 美咲",   can_drive=True),
        Staff("s3", "池田 裕二", can_drive=False),
        Staff("s4", "青木 健太", can_drive=True),
    ]
    return users, vehicles, staff


def get_sample_excel() -> bytes:
    """サンプルExcelファイルを生成してダウンロード用バイト列を返す"""
    users, vehicles, staff = get_demo_data()

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame([{
            "ID": u.user_id, "氏名": u.name, "住所": u.address,
            "緯度": u.lat, "経度": u.lng,
            "サービス種別": u.service_type.value,
            "車椅子": u.wheelchair,
            "同乗不可ID": ",".join(u.incompatible),
            "到着リミット(分)": u.pickup_latest,
            "送り目標(分)": u.dropoff_target,
        } for u in users]).to_excel(writer, sheet_name="利用者", index=False)

        pd.DataFrame([{
            "ID": v.vehicle_id, "車両名": v.name,
            "種別コード": v.vehicle_type, "定員": v.capacity,
            "車椅子対応": v.wheelchair_ok,
            "デポ緯度": v.depot_lat, "デポ経度": v.depot_lng,
        } for v in vehicles]).to_excel(writer, sheet_name="車両", index=False)

        pd.DataFrame([{
            "ID": s.staff_id, "氏名": s.name, "運転可否": s.can_drive,
        } for s in staff]).to_excel(writer, sheet_name="スタッフ", index=False)

    return buf.getvalue()


# ==============================================================
# 地図描画
# ==============================================================

def render_map(routes: list[AssignedRoute], depot_lat: float, depot_lng: float):
    """folium で送迎ルートを地図上に描画"""
    if not FOLIUM_AVAILABLE:
        st.info("📦 `folium` / `streamlit-folium` をインストールすると地図が表示されます")
        return

    m = folium.Map(location=[depot_lat, depot_lng], zoom_start=12,
                   tiles="CartoDB positron")

    # デポマーカー
    folium.Marker(
        [depot_lat, depot_lng],
        tooltip="🏠 事業所（デポ）",
        icon=folium.Icon(color="black", icon="home", prefix="fa")
    ).add_to(m)

    COLORS = ["blue", "green", "red", "purple", "orange", "darkred", "cadetblue"]

    for ri, route in enumerate(routes):
        color    = COLORS[ri % len(COLORS)]
        coords   = [[depot_lat, depot_lng]]
        for stop in route.stops:
            coords.append([stop["lat"], stop["lng"]])
        coords.append([depot_lat, depot_lng])

        # ルートライン
        folium.PolyLine(
            coords, color=color, weight=3, opacity=0.75,
            tooltip=route.vehicle.name
        ).add_to(m)

        # 各停留所マーカー
        for i, stop in enumerate(route.stops):
            h, m_min = divmod(stop["arrival_min"], 60)
            wc = "♿ " if stop["user"].wheelchair else ""
            popup_html = f"""
            <b>{i+1}. {stop['user'].name}</b><br>
            {wc}{stop['address']}<br>
            🕐 到着予定: {h:02d}:{m_min:02d}<br>
            🚗 {route.vehicle.name}
            """
            folium.CircleMarker(
                location=[stop["lat"], stop["lng"]],
                radius=8, color=color, fill=True, fill_opacity=0.85,
                tooltip=f"{i+1}. {stop['user'].name}",
                popup=folium.Popup(popup_html, max_width=220),
            ).add_to(m)

    st_folium(m, width=None, height=460, returned_objects=[])


# ==============================================================
# 結果テーブル生成
# ==============================================================

def routes_to_dataframe(routes: list[AssignedRoute]) -> pd.DataFrame:
    rows = []
    for route in routes:
        driver_name = route.driver.name if route.driver else "未定"
        for i, stop in enumerate(route.stops):
            h, m = divmod(stop["arrival_min"], 60)
            rows.append({
                "車両名":     route.vehicle.name,
                "運転担当":   driver_name,
                "順番":       i + 1,
                "利用者氏名": stop["user"].name,
                "サービス":   stop["user"].service_type.value,
                "住所":       stop["address"],
                "到着予定":   f"{h:02d}:{m:02d}",
                "車椅子":     "♿" if stop["user"].wheelchair else "",
            })
    return pd.DataFrame(rows)


# ==============================================================
# Streamlit UI メイン
# ==============================================================

def main():
    # ---- ヘッダー ----
    st.markdown("""
    <div class="main-header">
      <h1>🚌 送迎ルート最適化システム</h1>
      <p>放課後等デイサービス・就労継続支援A型/B型 対応　｜　VRP（配送計画問題）自動最適化</p>
    </div>
    """, unsafe_allow_html=True)

    # ---- サイドバー：設定 ----
    with st.sidebar:
        st.markdown("### ⚙️ 実行設定")

        trip_type_label = st.radio(
            "便の種別",
            options=["迎え便（自宅→施設）", "送り便（施設→自宅）"],
            index=0,
        )
        trip_type = TripType.PICKUP if "迎え" in trip_type_label else TripType.DROPOFF

        st.markdown("---")
        st.markdown("**時刻設定**")

        col1, col2 = st.columns(2)
        with col1:
            start_h = st.number_input("出発 時", min_value=5, max_value=12, value=8)
            start_m = st.number_input("出発 分", min_value=0, max_value=55, value=0, step=5)
        with col2:
            limit_h = st.number_input("リミット 時", min_value=6, max_value=13, value=9)
            limit_m = st.number_input("リミット 分", min_value=0, max_value=55, value=0, step=5)

        start_min = start_h * 60 + start_m
        limit_min = limit_h * 60 + limit_m

        st.markdown("---")
        algo_label = "🤖 OR-Tools（高精度）" if ORTOOLS_AVAILABLE else "⚡ グリーディ（高速）"
        st.info(f"使用アルゴリズム:\n{algo_label}")

        st.markdown("---")
        st.markdown("**サンプルExcelをダウンロード**")
        if OPENPYXL_AVAILABLE:
            st.download_button(
                "📥 サンプルExcel",
                data=get_sample_excel(),
                file_name="送迎マスタ_サンプル.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # ============================================================
    # STEP 1: データ読み込み
    # ============================================================
    st.markdown('<div class="step-badge">STEP 1　データの読み込み</div>', unsafe_allow_html=True)
    st.markdown("#### 利用者・車両・スタッフのマスタデータを読み込みます")

    col_upload, col_demo = st.columns([2, 1])

    users, vehicles, staff = None, None, None

    with col_upload:
        uploaded = st.file_uploader(
            "Excelファイルをアップロード（シート名: 利用者 / 車両 / スタッフ）",
            type=["xlsx", "xls"],
            help="左サイドバーからサンプルExcelをダウンロードして記入してください"
        )
        if uploaded:
            try:
                with st.spinner("Excelを読み込み中..."):
                    users, vehicles, staff = parse_excel_upload(uploaded)
                st.success(f"✅ 読み込み完了　利用者 {len(users)}名 / 車両 {len(vehicles)}台 / スタッフ {len(staff)}名")
            except Exception as e:
                st.error(f"❌ ファイルの読み込みに失敗しました: {e}")
                st.info("シート名が「利用者」「車両」「スタッフ」になっているか確認してください")

    with col_demo:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🎯 デモデータで実行", use_container_width=True, type="secondary"):
            users, vehicles, staff = get_demo_data()
            st.success(f"✅ デモデータ読み込み完了　利用者 {len(users)}名 / 車両 {len(vehicles)}台 / スタッフ {len(staff)}名")

    # データプレビュー
    if users and vehicles and staff:
        with st.expander("📋 読み込みデータのプレビュー", expanded=False):
            tab1, tab2, tab3 = st.tabs(["👶 利用者", "🚗 車両", "👤 スタッフ"])
            with tab1:
                st.dataframe(pd.DataFrame([{
                    "ID": u.user_id, "氏名": u.name, "住所": u.address,
                    "サービス": u.service_type.value,
                    "車椅子": "♿" if u.wheelchair else "",
                    "同乗不可": ",".join(u.incompatible) or "なし",
                    "到着リミット": f"{u.pickup_latest//60:02d}:{u.pickup_latest%60:02d}",
                } for u in users]), use_container_width=True, hide_index=True)

            with tab2:
                st.dataframe(pd.DataFrame([{
                    "車両名": v.name, "種別": v.vehicle_type,
                    "定員": v.capacity,
                    "車椅子対応": "✅" if v.wheelchair_ok else "✗",
                } for v in vehicles]), use_container_width=True, hide_index=True)

            with tab3:
                st.dataframe(pd.DataFrame([{
                    "氏名": s.name,
                    "運転可否": "✅ 可能" if s.can_drive else "❌ 不可",
                } for s in staff]), use_container_width=True, hide_index=True)

    st.divider()

    # ============================================================
    # STEP 2: 最適化実行
    # ============================================================
    st.markdown('<div class="step-badge">STEP 2　最適化の実行</div>', unsafe_allow_html=True)

    run_disabled = not (users and vehicles and staff)
    run_clicked  = st.button(
        "🚀　最適化を実行する",
        disabled=run_disabled,
        type="primary",
        use_container_width=True,
    )
    if run_disabled:
        st.caption("👆 STEP 1 でデータを読み込んでから実行してください")

    if run_clicked and users and vehicles and staff:
        # 制約チェック
        checker = ConstraintChecker()
        errors  = checker.validate(users, vehicles, staff)
        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        with st.spinner("🔄 ルートを最適化しています... しばらくお待ちください"):
            # 距離行列構築
            depot   = (vehicles[0].depot_lat, vehicles[0].depot_lng)
            locs    = [depot] + [(u.lat, u.lng) for u in users]
            builder = DistanceMatrixBuilder()
            matrix  = builder.build(locs)

            # ソルバー実行
            solver = TransportVRPSolver(
                users=users, vehicles=vehicles, staff=staff,
                distance_matrix=matrix,
                trip_type=trip_type,
                depot_arrival_limit_min=limit_min,
                start_time_min=start_min,
            )
            routes = solver.solve()

        st.session_state["routes"]    = routes
        st.session_state["depot"]     = depot
        st.session_state["trip_type"] = trip_type
        st.success(f"✅ 最適化完了！　{len(routes)} 台のルートを生成しました")

    st.divider()

    # ============================================================
    # STEP 3: 結果表示・ダウンロード
    # ============================================================
    st.markdown('<div class="step-badge">STEP 3　結果の確認とダウンロード</div>', unsafe_allow_html=True)

    if "routes" not in st.session_state:
        st.info("👆 STEP 2 で最適化を実行すると、ここに結果が表示されます")
        return

    routes    = st.session_state["routes"]
    depot     = st.session_state["depot"]
    trip_type = st.session_state["trip_type"]

    if not routes:
        st.error("ルートを生成できませんでした。データを確認してください。")
        return

    # ---- サマリーメトリクス ----
    total_users  = sum(len(r.stops) for r in routes)
    total_veh    = len(routes)
    wc_count     = sum(1 for r in routes for s in r.stops if s["user"].wheelchair)

    checker      = ConstraintChecker()
    forbidden    = checker.get_forbidden_pairs(users) if users else set()

    m1, m2, m3, m4 = st.columns(4)
    for col, (val, label) in zip(
        [m1, m2, m3, m4],
        [(total_users, "割り当て済み利用者"),
         (total_veh,   "稼働車両数"),
         (wc_count,    "車椅子利用者"),
         (len(forbidden), "同乗不可ペア")]
    ):
        with col:
            st.markdown(f"""
            <div class="metric-box">
              <div class="val">{val}</div>
              <div class="label">{label}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---- 制約検証 ----
    with st.expander("🔍 制約条件 検証サマリー", expanded=True):
        all_ok = True
        for route in routes:
            users_in  = [s["user"] for s in route.stops]
            ok_cap    = len(users_in) <= route.vehicle.capacity
            ok_wc     = not (any(u.wheelchair for u in users_in) and not route.vehicle.wheelchair_ok)
            ok_incomp = not any(
                tuple(sorted([u1.user_id, u2.user_id])) in forbidden
                for i, u1 in enumerate(users_in)
                for u2 in users_in[i+1:]
            )
            ok_driver = route.driver is not None and route.driver.can_drive

            all_ok = all_ok and ok_cap and ok_wc and ok_incomp and ok_driver

            stat = lambda ok: f'<span class="constraint-ok">✅</span>' if ok else f'<span class="constraint-fail">❌ 違反</span>'
            driver_name = route.driver.name if route.driver else "未定"

            st.markdown(
                f"**{route.vehicle.name}** ({len(users_in)}/{route.vehicle.capacity}名)　"
                f"定員:{stat(ok_cap)}　車椅子:{stat(ok_wc)}　"
                f"同乗不可:{stat(ok_incomp)}　"
                f"運転者:{driver_name} {stat(ok_driver)}",
                unsafe_allow_html=True
            )

        if all_ok:
            st.success("🎉 全制約条件をクリアしています！")

    # ---- 結果テーブル ----
    st.markdown("#### 📋 送迎ルート一覧")
    df = routes_to_dataframe(routes)
    st.dataframe(df, use_container_width=True, hide_index=True,
                 column_config={
                     "順番":    st.column_config.NumberColumn(width="small"),
                     "到着予定": st.column_config.TextColumn(width="small"),
                     "車椅子":  st.column_config.TextColumn(width="small"),
                 })

    # ---- ダウンロードボタン ----
    col_dl1, col_dl2 = st.columns([1, 3])
    with col_dl1:
        excel_bytes = build_excel_output(routes)
        ext         = "xlsx" if OPENPYXL_AVAILABLE else "csv"
        mime        = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                       if OPENPYXL_AVAILABLE else "text/csv")
        st.download_button(
            label=f"📥 Excel でダウンロード (.{ext})",
            data=excel_bytes,
            file_name=f"送迎ルート_{trip_type.value}.{ext}",
            mime=mime,
            type="primary",
            use_container_width=True,
        )

    st.divider()

    # ---- 地図表示 ----
    st.markdown("#### 🗺️ 送迎ルートマップ")
    if FOLIUM_AVAILABLE:
        render_map(routes, depot[0], depot[1])
    else:
        st.info(
            "📦 地図表示には `folium` と `streamlit-folium` が必要です。\n"
            "`requirements.txt` に追加してデプロイしてください。"
        )
        # 代替：座標テーブル
        coord_rows = []
        for route in routes:
            for i, stop in enumerate(route.stops):
                coord_rows.append({
                    "車両": route.vehicle.name, "順番": i+1,
                    "氏名": stop["user"].name,
                    "緯度": stop["lat"], "経度": stop["lng"],
                })
        st.dataframe(pd.DataFrame(coord_rows), use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
