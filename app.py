"""
送迎ルート自動作成システム - Streamlit Webアプリ v2
=====================================================
放課後等デイサービス / 就労継続支援A型・B型 対応

v2 追加要件:
  1. 迎え・送り 1回実行で両便計算、Excelは2シート出力
  2. 店舗ごとに独立してVRPを実行（混載禁止）
  3. スタッフ優先度 + Fixed Cost による負荷分散
  4. 乗降時間（Service Time）を時間ディメンションに加算
  5. 当日欠席トグル UI（st.data_editor）

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
# ページ設定
# ==============================================================
st.set_page_config(
    page_title="送迎ルート最適化システム v2",
    page_icon="🚌",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ==============================================================
# カスタム CSS
# ==============================================================
st.markdown("""
<style>
  html, body, [class*="css"] { font-family: 'Noto Sans JP', 'メイリオ', sans-serif; }

  .main-header {
    background: linear-gradient(135deg, #1C2833 0%, #2C3E50 100%);
    border-radius: 12px;
    padding: 24px 28px;
    margin-bottom: 24px;
    color: white;
  }
  .main-header h1 { font-size: 24px; margin: 0; font-weight: 800; letter-spacing: 0.04em; }
  .main-header p  { font-size: 12px; margin: 6px 0 0; opacity: 0.7; }

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

  .shop-tag {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 12px;
    font-size: 11px;
    font-weight: 700;
    margin-right: 4px;
  }

  .metric-box {
    background: white;
    border: 1px solid #DDD;
    border-radius: 8px;
    padding: 14px;
    text-align: center;
  }
  .metric-box .val   { font-size: 28px; font-weight: 800; color: #2C3E50; }
  .metric-box .label { font-size: 12px; color: #888; margin-top: 2px; }

  .constraint-ok   { color: #27AE60; font-weight: 700; }
  .constraint-fail { color: #E74C3C; font-weight: 700; }

  .stButton > button { border-radius: 8px !important; font-weight: 600 !important; }
</style>
""", unsafe_allow_html=True)


# ==============================================================
# データクラス定義
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
    user_id:        str
    name:           str
    address:        str
    lat:            float
    lng:            float
    service_type:   ServiceType
    shop:           str                          # 【v2追加】店舗名（例: "A店"）
    wheelchair:     bool = False
    incompatible:   list[str] = field(default_factory=list)
    pickup_latest:  int = 540                    # 施設到着リミット（分）
    dropoff_target: int = 1050                   # 送り目標時刻（分）
    # 【v2追加】乗降時間（秒）: 通常300秒(5分)、車椅子600秒(10分)
    service_time:   int = 300

    def __post_init__(self):
        # 車椅子フラグに応じてservice_timeを自動設定
        if self.wheelchair:
            self.service_time = 600


@dataclass
class Vehicle:
    vehicle_id:    str
    name:          str
    vehicle_type:  str
    capacity:      int
    shop:          str           # 【v2追加】所属店舗
    wheelchair_ok: bool = False
    depot_lat:     float = 36.6953
    depot_lng:     float = 137.2113


@dataclass
class Staff:
    staff_id:  str
    name:      str
    shop:      str       # 【v2追加】所属店舗
    can_drive: bool = True
    priority:  int  = 1  # 【v2追加】優先度（1=最優先, 数字が大きいほど低優先）


@dataclass
class AssignedRoute:
    vehicle:   Vehicle
    driver:    Optional[Staff]
    trip_type: TripType
    shop:      str
    stops:     list[dict]
    total_min: int


# ==============================================================
# 距離行列ビルダー（ハーバーサイン距離）
# ==============================================================

class DistanceMatrixBuilder:
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
        return max(1, int(dist_km / avg_speed_kmh * 3600))


# ==============================================================
# 制約チェッカー
# ==============================================================

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
            names = [u.name for u in wc_users]
            errors.append(f"❌ 車椅子利用者 {names} がいますが車椅子対応車両がありません")
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


# ==============================================================
# VRP ソルバー本体（v2: 優先度Fixed Cost + 乗降時間 対応）
# ==============================================================

# 優先度→Fixed Cost のマッピング（優先度が低いほど高コスト）
PRIORITY_COST_MAP = {
    1: 0,
    2: 5000,
    3: 15000,
    4: 30000,
    9: 999999,  # 実質使用不可
}


class TransportVRPSolver:
    """
    店舗単位で独立して呼び出されることを前提とするVRPソルバー。
    （混載禁止はデータを店舗ごとに分割することで実現）

    v2変更点:
      - SetFixedCostOfVehicle で優先度に応じたペナルティを設定
      - time_callback に service_time（乗降時間）を加算
    """

    TIME_LIMIT_SEC = 30

    def __init__(
        self,
        users:           list[User],
        vehicles:        list[Vehicle],
        staff:           list[Staff],
        distance_matrix: list[list[int]],
        trip_type:       TripType = TripType.PICKUP,
        depot_arrival_limit_min: int = 540,
        start_time_min:          int = 480,
    ):
        self.users    = users
        self.vehicles = vehicles
        # 運転不可スタッフを除外し、優先度順にソート
        self.staff    = sorted(
            [s for s in staff if s.can_drive],
            key=lambda s: s.priority
        )
        self.matrix    = distance_matrix
        self.trip_type = trip_type
        self.depot_arrival_limit = depot_arrival_limit_min * 60
        self.start_time          = start_time_min * 60

        checker = ConstraintChecker()
        self.forbidden_pairs = checker.get_forbidden_pairs(users)

    def solve(self) -> list[AssignedRoute]:
        if ORTOOLS_AVAILABLE and len(self.users) > 0:
            result = self._solve_with_ortools()
            if result is not None:
                return result
        return self._greedy_fallback()

    # ----------------------------------------------------------
    # OR-Tools 本体
    # ----------------------------------------------------------
    def _solve_with_ortools(self) -> Optional[list[AssignedRoute]]:
        n_users    = len(self.users)
        n_vehicles = len(self.vehicles)
        if n_vehicles == 0 or n_users == 0:
            return []

        n_nodes = n_users + 1   # ノード0=デポ, 1..n_users=各利用者

        manager = pywrapcp.RoutingIndexManager(n_nodes, n_vehicles, 0)
        routing = pywrapcp.RoutingModel(manager)

        # ---- 【v2】時間コールバック: 移動時間 + 乗降時間(service_time) ----
        def time_callback(from_idx, to_idx):
            from_node = manager.IndexToNode(from_idx)
            to_node   = manager.IndexToNode(to_idx)
            travel    = self.matrix[from_node][to_node]
            # from_node が利用者ノードの場合は乗降時間を加算
            svc = self.users[from_node - 1].service_time if from_node > 0 else 0
            return travel + svc

        transit_cb = routing.RegisterTransitCallback(time_callback)
        routing.SetArcCostEvaluatorOfAllVehicles(transit_cb)

        # ---- 時間ディメンション ----
        max_time = max(self.depot_arrival_limit, self.start_time + 3 * 3600)
        routing.AddDimension(
            transit_cb,
            600,        # 待機許容（秒）
            max_time,   # 最大積算時間（秒）
            True,       # 始点を0に固定
            "Time",
        )
        time_dim = routing.GetDimensionOrDie("Time")

        # 迎え便のみ: デポ到着時刻をリミット以内に制約
        if self.trip_type == TripType.PICKUP:
            depot_idx = manager.NodeToIndex(0)
            time_dim.CumulVar(depot_idx).SetMax(self.depot_arrival_limit)

        # ---- 定員ディメンション ----
        def demand_cb(from_idx):
            return 0 if manager.IndexToNode(from_idx) == 0 else 1

        demand_cb_idx = routing.RegisterUnaryTransitCallback(demand_cb)
        routing.AddDimensionWithVehicleCapacity(
            demand_cb_idx,
            0,
            [v.capacity for v in self.vehicles],
            True,
            "Capacity",
        )

        # ---- 【v2】Fixed Cost: 優先度が低い車両ほど高コスト ----
        for vi, vehicle in enumerate(self.vehicles):
            driver = self._get_driver_for_vehicle(vi)
            priority = driver.priority if driver else 9
            fixed_cost = PRIORITY_COST_MAP.get(priority, 999999)
            routing.SetFixedCostOfVehicle(fixed_cost, vi)

        # ---- 車椅子制約 ----
        for i, u in enumerate(self.users):
            if u.wheelchair:
                node_idx = manager.NodeToIndex(i + 1)
                for vi, v in enumerate(self.vehicles):
                    if not v.wheelchair_ok:
                        routing.VehicleVar(node_idx).RemoveValue(vi)

        # ---- 同乗不可制約 ----
        for uid1, uid2 in self.forbidden_pairs:
            i1 = next((i+1 for i, u in enumerate(self.users) if u.user_id == uid1), None)
            i2 = next((i+1 for i, u in enumerate(self.users) if u.user_id == uid2), None)
            if i1 and i2:
                ni1 = manager.NodeToIndex(i1)
                ni2 = manager.NodeToIndex(i2)
                routing.solver().Add(
                    routing.VehicleVar(ni1) != routing.VehicleVar(ni2)
                )

        # ---- 探索パラメータ ----
        params = pywrapcp.DefaultRoutingSearchParameters()
        params.first_solution_strategy = (
            routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC
        )
        params.local_search_metaheuristic = (
            routing_enums_pb2.LocalSearchMetaheuristic.GUIDED_LOCAL_SEARCH
        )
        params.time_limit.FromSeconds(self.TIME_LIMIT_SEC)

        solution = routing.SolveWithParameters(params)
        if not solution:
            return None

        return self._extract_routes(manager, routing, solution)

    def _extract_routes(self, manager, routing, solution) -> list[AssignedRoute]:
        time_dim = routing.GetDimensionOrDie("Time")
        routes   = []

        for vi in range(len(self.vehicles)):
            idx   = routing.Start(vi)
            stops = []

            while not routing.IsEnd(idx):
                node = manager.IndexToNode(idx)
                if node != 0:
                    u       = self.users[node - 1]
                    arr_sec = solution.Min(time_dim.CumulVar(idx))
                    # 乗降時間を除いた「到着時刻」に換算
                    arr_min = (self.start_time + arr_sec) // 60
                    stops.append({
                        "user":        u,
                        "arrival_min": arr_min,
                        "address":     u.address,
                        "lat":         u.lat,
                        "lng":         u.lng,
                    })
                idx = solution.Value(routing.NextVar(idx))

            if not stops:
                continue

            driver = self._get_driver_for_vehicle(vi)
            routes.append(AssignedRoute(
                vehicle   = self.vehicles[vi],
                driver    = driver,
                trip_type = self.trip_type,
                shop      = self.vehicles[vi].shop,
                stops     = stops,
                total_min = 0,
            ))

        return routes

    # ----------------------------------------------------------
    # グリーディフォールバック
    # ----------------------------------------------------------
    def _greedy_fallback(self) -> list[AssignedRoute]:
        # 車椅子対応→大型→普通の順でソート
        vehicles_sorted = sorted(
            self.vehicles,
            key=lambda v: (0 if v.wheelchair_ok else 1, -v.capacity)
        )
        routes     = []
        unassigned = list(self.users)

        for vi, vehicle in enumerate(vehicles_sorted):
            if not unassigned:
                break

            driver   = self._get_driver_for_vehicle(vi)
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

            ordered  = self._nearest_neighbor(assigned)
            stops    = []
            cur_node = 0
            cur_time = self.start_time  # 秒

            for u in ordered:
                uid       = self.users.index(u) + 1
                cur_time += self.matrix[cur_node][uid] + u.service_time
                stops.append({
                    "user":        u,
                    "arrival_min": cur_time // 60,
                    "address":     u.address,
                    "lat":         u.lat,
                    "lng":         u.lng,
                })
                cur_node = uid

            routes.append(AssignedRoute(
                vehicle   = vehicle,
                driver    = driver,
                trip_type = self.trip_type,
                shop      = vehicle.shop,
                stops     = stops,
                total_min = (cur_time - self.start_time) // 60,
            ))

        if unassigned:
            names = [u.name for u in unassigned]
            st.warning(f"⚠️ [{self.vehicles[0].shop if self.vehicles else ''}] 以下の利用者が割り当て不可: {names}")

        return routes

    def _nearest_neighbor(self, users: list[User]) -> list[User]:
        if not users:
            return []
        remaining, ordered, cur = list(users), [], 0
        while remaining:
            nearest = min(
                remaining,
                key=lambda u: self.matrix[cur][self.users.index(u) + 1]
            )
            ordered.append(nearest)
            cur = self.users.index(nearest) + 1
            remaining.remove(nearest)
        return ordered

    def _get_driver_for_vehicle(self, vehicle_idx: int) -> Optional[Staff]:
        """優先度順に並んだスタッフから車両インデックスに対応するドライバーを取得"""
        if not self.staff:
            return None
        return self.staff[vehicle_idx % len(self.staff)]


# ==============================================================
# 【v2】店舗別にVRPを実行してルートをまとめる
# ==============================================================

def run_all_shops(
    users:    list[User],
    vehicles: list[Vehicle],
    staff:    list[Staff],
    trip_type: TripType,
    start_min: int,
    limit_min: int,
) -> list[AssignedRoute]:
    """
    店舗ごとにデータを分割し、それぞれ独立してVRPを実行。
    結果を店舗名でソートしてマージして返す。
    """
    shops  = sorted(set(u.shop for u in users))
    routes = []

    for shop in shops:
        shop_users    = [u for u in users    if u.shop == shop]
        shop_vehicles = [v for v in vehicles if v.shop == shop]
        shop_staff    = [s for s in staff    if s.shop == shop]

        if not shop_users or not shop_vehicles:
            continue

        # 距離行列（デポ + 利用者）
        if shop_vehicles:
            depot = (shop_vehicles[0].depot_lat, shop_vehicles[0].depot_lng)
        else:
            depot = (36.695, 137.211)
        locs   = [depot] + [(u.lat, u.lng) for u in shop_users]
        matrix = DistanceMatrixBuilder().build(locs)

        solver = TransportVRPSolver(
            users                    = shop_users,
            vehicles                 = shop_vehicles,
            staff                    = shop_staff,
            distance_matrix          = matrix,
            trip_type                = trip_type,
            depot_arrival_limit_min  = limit_min,
            start_time_min           = start_min,
        )
        shop_routes = solver.solve()
        routes.extend(shop_routes)

    return routes


# ==============================================================
# Excel 入出力
# ==============================================================

def parse_excel_upload(
    uploaded_file,
) -> tuple[list[User], list[Vehicle], list[Staff]]:
    """
    アップロードされたExcelを読み込んでデータクラスに変換。

    【v2 新カラム】
      利用者シート: 「店舗」
      スタッフシート: 「店舗」「優先度」
      車両シート:   「店舗」
    """
    xl = pd.ExcelFile(uploaded_file)

    service_map = {
        "放課後等デイサービス": ServiceType.HOUKAGO_DEI,
        "A型":                ServiceType.A_TYPE,
        "B型":                ServiceType.B_TYPE,
    }

    # ---- 利用者 ----
    df_u = xl.parse("利用者")
    users = []
    for i, row in df_u.iterrows():
        incomp_raw = str(row.get("同乗不可ID", "")).strip()
        incomp = (
            [x.strip() for x in incomp_raw.split(",") if x.strip()]
            if incomp_raw not in ("", "nan") else []
        )
        wc = bool(row.get("車椅子", False))
        users.append(User(
            user_id       = str(row.get("ID", f"u{i+1}")),
            name          = str(row["氏名"]),
            address       = str(row.get("住所", "")),
            lat           = float(row.get("緯度", 36.695)),
            lng           = float(row.get("経度", 137.211)),
            service_type  = service_map.get(
                str(row.get("サービス種別", "")), ServiceType.HOUKAGO_DEI
            ),
            shop          = str(row.get("店舗", "A店")),   # 【v2追加】
            wheelchair    = wc,
            incompatible  = incomp,
            pickup_latest = int(row.get("到着リミット(分)", 540)),
            dropoff_target= int(row.get("送り目標(分)", 1050)),
            service_time  = 600 if wc else 300,            # 【v2: 車椅子10分/通常5分】
        ))

    # ---- 車両 ----
    df_v = xl.parse("車両")
    type_cap = {"large": 7, "normal": 4, "kei": 3}
    vehicles = []
    for i, row in df_v.iterrows():
        vtype = str(row.get("種別コード", "normal"))
        vehicles.append(Vehicle(
            vehicle_id    = str(row.get("ID", f"v{i+1}")),
            name          = str(row["車両名"]),
            vehicle_type  = vtype,
            capacity      = int(row.get("定員", type_cap.get(vtype, 4))),
            shop          = str(row.get("店舗", "A店")),   # 【v2追加】
            wheelchair_ok = bool(row.get("車椅子対応", False)),
            depot_lat     = float(row.get("デポ緯度", 36.695)),
            depot_lng     = float(row.get("デポ経度", 137.211)),
        ))

    # ---- スタッフ ----
    df_s = xl.parse("スタッフ")
    staff = []
    for i, row in df_s.iterrows():
        staff.append(Staff(
            staff_id  = str(row.get("ID", f"s{i+1}")),
            name      = str(row["氏名"]),
            shop      = str(row.get("店舗", "A店")),         # 【v2追加】
            can_drive = bool(row.get("運転可否", True)),
            priority  = int(row.get("優先度", 1)),            # 【v2追加】
        ))

    return users, vehicles, staff


def build_excel_output(
    pickup_routes:  list[AssignedRoute],
    dropoff_routes: list[AssignedRoute],
) -> bytes:
    """
    【v2】迎え便・送り便を別シートに出力する1ファイルのExcelを生成。
    店舗ごとにブロックを分けて出力する。
    """
    if not OPENPYXL_AVAILABLE:
        # openpyxl がない場合はCSVで代替（迎えのみ）
        rows = _routes_to_rows(pickup_routes, "迎え")
        rows += _routes_to_rows(dropoff_routes, "送り")
        buf = io.StringIO()
        pd.DataFrame(rows).to_csv(buf, index=False, encoding="utf-8-sig")
        return buf.getvalue().encode("utf-8-sig")

    wb = Workbook()
    wb.remove(wb.active)  # デフォルトシートを削除

    for routes, sheet_name in [
        (pickup_routes,  "迎え便"),
        (dropoff_routes, "送り便"),
    ]:
        ws = wb.create_sheet(title=sheet_name)
        _write_route_sheet(ws, routes, sheet_name)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _write_route_sheet(ws, routes: list[AssignedRoute], sheet_title: str):
    """1シート分の送迎表を書き込む（店舗ごとにブロック分け）"""

    HDR_FILL  = PatternFill("solid", fgColor="1C2833")
    SHOP_FILLS = {
        0: PatternFill("solid", fgColor="154360"),  # 店舗1: 濃青
        1: PatternFill("solid", fgColor="145A32"),  # 店舗2: 濃緑
        2: PatternFill("solid", fgColor="6E2F1A"),  # 店舗3: 濃茶
        3: PatternFill("solid", fgColor="4A235A"),  # 店舗4: 濃紫
    }
    VEH_FILLS = {
        0: PatternFill("solid", fgColor="D6EAF8"),
        1: PatternFill("solid", fgColor="D5F5E3"),
        2: PatternFill("solid", fgColor="FEF9E7"),
        3: PatternFill("solid", fgColor="F9EBF8"),
    }
    WC_FILL  = PatternFill("solid", fgColor="FADBD8")
    ALT_FILL = PatternFill("solid", fgColor="F8F9FA")

    def bdr():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    def cell_style(c, font=None, fill=None, align=None):
        if font:  c.font      = font
        if fill:  c.fill      = fill
        if align: c.alignment = align
        c.border = bdr()

    # タイトル行
    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value     = f"🚌  送迎ルート表　【{sheet_title}】"
    c.font      = Font(bold=True, size=14, color="FFFFFF", name="メイリオ")
    c.fill      = HDR_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # ヘッダー
    headers = ["店舗", "車両名", "運転担当", "優先度", "順番",
               "利用者氏名", "サービス", "住所", "到着予定", "備考"]
    widths  = [10,    20,       14,         6,       6,
               16,            16,       34,     10,      12]
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=2, column=col, value=h)
        cell_style(c,
                   font=Font(bold=True, color="FFFFFF", size=10, name="メイリオ"),
                   fill=HDR_FILL,
                   align=Alignment(horizontal="center", vertical="center"))
        ws.column_dimensions[chr(64+col)].width = w
    ws.row_dimensions[2].height = 22

    # 店舗ごとにルートをグループ化（ソート済み）
    shops = sorted(set(r.shop for r in routes))
    row   = 3

    for shop_idx, shop in enumerate(shops):
        shop_routes = [r for r in routes if r.shop == shop]
        shop_fill   = SHOP_FILLS.get(shop_idx % len(SHOP_FILLS))
        veh_fill    = VEH_FILLS.get(shop_idx % len(VEH_FILLS))

        # 店舗区切りヘッダー
        ws.merge_cells(f"A{row}:J{row}")
        c = ws.cell(row=row, column=1, value=f"　🏠 {shop}")
        c.font      = Font(bold=True, size=12, color="FFFFFF", name="メイリオ")
        c.fill      = shop_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = bdr()
        ws.row_dimensions[row].height = 26
        row += 1

        for route in shop_routes:
            driver_name = route.driver.name     if route.driver else "未定"
            priority    = route.driver.priority if route.driver else "-"

            for i, stop in enumerate(route.stops):
                h, m   = divmod(stop["arrival_min"], 60)
                wc     = "♿ 車椅子" if stop["user"].wheelchair else ""
                svc    = stop["user"].service_type.value

                fill = WC_FILL if stop["user"].wheelchair else (
                    veh_fill if i == 0 else ALT_FILL
                )
                data = [
                    route.shop          if i == 0 else "",
                    route.vehicle.name  if i == 0 else "",
                    driver_name         if i == 0 else "",
                    priority            if i == 0 else "",
                    i + 1,
                    stop["user"].name,
                    svc,
                    stop["address"],
                    f"{h:02d}:{m:02d}",
                    wc,
                ]
                for col, val in enumerate(data, 1):
                    c = ws.cell(row=row, column=col, value=val)
                    cell_style(c,
                               font=Font(bold=(i == 0), size=10, name="メイリオ"),
                               fill=fill,
                               align=Alignment(
                                   horizontal="center" if col in [1,2,3,4,5,9,10] else "left",
                                   vertical="center"
                               ))
                ws.row_dimensions[row].height = 20
                row += 1

        row += 1  # 店舗間空白行


def _routes_to_rows(routes: list[AssignedRoute], trip_label: str) -> list[dict]:
    rows = []
    for r in routes:
        for i, stop in enumerate(r.stops):
            h, m = divmod(stop["arrival_min"], 60)
            rows.append({
                "便":       trip_label,
                "店舗":     r.shop,
                "車両名":   r.vehicle.name,
                "運転担当": r.driver.name if r.driver else "未定",
                "優先度":   r.driver.priority if r.driver else "-",
                "順番":     i + 1,
                "氏名":     stop["user"].name,
                "サービス": stop["user"].service_type.value,
                "住所":     stop["address"],
                "到着予定": f"{h:02d}:{m:02d}",
                "車椅子":   "♿" if stop["user"].wheelchair else "",
            })
    return rows


# ==============================================================
# デモデータ（v2: 店舗・優先度・service_time 対応）
# ==============================================================

def get_demo_data() -> tuple[list[User], list[Vehicle], list[Staff]]:
    users = [
        # A店
        User("u1",  "山田 太郎",   "富山市上袋100",   36.720, 137.210, ServiceType.HOUKAGO_DEI, "A店", False, [],      540, 1050),
        User("u2",  "鈴木 花子",   "富山市堀川200",   36.695, 137.220, ServiceType.HOUKAGO_DEI, "A店", False, ["u3"], 540, 1050),
        User("u3",  "田中 一郎",   "富山市婦中300",   36.660, 137.160, ServiceType.A_TYPE,      "A店", True,  ["u2"], 540, 1050),
        User("u4",  "佐藤 愛",     "富山市大沢野400", 36.630, 137.230, ServiceType.B_TYPE,      "A店", False, [],      540, 1050),
        User("u5",  "高橋 健太",   "富山市八尾500",   36.590, 137.270, ServiceType.HOUKAGO_DEI, "A店", False, [],      540, 1050),
        # B店
        User("u6",  "渡辺 さくら", "富山市上袋600",   36.725, 137.215, ServiceType.B_TYPE,      "B店", False, [],      540, 1050),
        User("u7",  "伊藤 翔",     "富山市堀川700",   36.700, 137.225, ServiceType.A_TYPE,      "B店", False, [],      540, 1050),
        User("u8",  "中村 みな",   "富山市婦中800",   36.655, 137.155, ServiceType.B_TYPE,      "B店", False, [],      540, 1050),
        User("u9",  "小林 大輝",   "富山市大沢野900", 36.625, 137.235, ServiceType.HOUKAGO_DEI, "B店", False, [],      540, 1050),
        User("u10", "加藤 りん",   "富山市八尾1000",  36.585, 137.265, ServiceType.A_TYPE,      "B店", False, [],      540, 1050),
        # C店
        User("u11", "中島 陽斗",   "富山市上袋1100",  36.715, 137.205, ServiceType.HOUKAGO_DEI, "C店", False, [],      540, 1050),
        User("u12", "斉藤 みゆ",   "富山市堀川1200",  36.690, 137.215, ServiceType.B_TYPE,      "C店", True,  [],      540, 1050),
    ]
    # service_time を __post_init__ で自動設定
    for u in users:
        u.service_time = 600 if u.wheelchair else 300

    vehicles = [
        Vehicle("v1", "A-1号車（大型）", "large",  7, "A店", wheelchair_ok=True,  depot_lat=36.695, depot_lng=137.211),
        Vehicle("v2", "A-2号車（普通）", "normal", 4, "A店", wheelchair_ok=False, depot_lat=36.695, depot_lng=137.211),
        Vehicle("v3", "B-1号車（大型）", "large",  7, "B店", wheelchair_ok=False, depot_lat=36.710, depot_lng=137.200),
        Vehicle("v4", "B-2号車（普通）", "normal", 4, "B店", wheelchair_ok=False, depot_lat=36.710, depot_lng=137.200),
        Vehicle("v5", "C-1号車（大型）", "large",  7, "C店", wheelchair_ok=True,  depot_lat=36.680, depot_lng=137.220),
    ]
    staff = [
        Staff("s1", "林 誠一",   "A店", can_drive=True,  priority=1),
        Staff("s2", "森 美咲",   "A店", can_drive=True,  priority=2),
        Staff("s3", "池田 裕二", "A店", can_drive=False, priority=9),
        Staff("s4", "宇野 幸子", "B店", can_drive=True,  priority=1),
        Staff("s5", "川口 拓也", "B店", can_drive=True,  priority=2),
        Staff("s6", "高木 雄介", "C店", can_drive=True,  priority=1),
        Staff("s7", "中島 奈々", "C店", can_drive=True,  priority=3),
    ]
    return users, vehicles, staff


def get_sample_excel() -> bytes:
    """入力フォーマットのサンプルExcelを生成"""
    users, vehicles, staff = get_demo_data()
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame([{
            "ID": u.user_id, "氏名": u.name, "住所": u.address,
            "緯度": u.lat, "経度": u.lng,
            "サービス種別": u.service_type.value,
            "店舗": u.shop,
            "車椅子": u.wheelchair,
            "同乗不可ID": ",".join(u.incompatible),
            "到着リミット(分)": u.pickup_latest,
            "送り目標(分)": u.dropoff_target,
        } for u in users]).to_excel(writer, sheet_name="利用者", index=False)

        pd.DataFrame([{
            "ID": v.vehicle_id, "車両名": v.name,
            "種別コード": v.vehicle_type, "定員": v.capacity,
            "店舗": v.shop,
            "車椅子対応": v.wheelchair_ok,
            "デポ緯度": v.depot_lat, "デポ経度": v.depot_lng,
        } for v in vehicles]).to_excel(writer, sheet_name="車両", index=False)

        pd.DataFrame([{
            "ID": s.staff_id, "氏名": s.name,
            "店舗": s.shop,
            "運転可否": s.can_drive,
            "優先度": s.priority,
        } for s in staff]).to_excel(writer, sheet_name="スタッフ", index=False)

    return buf.getvalue()


# ==============================================================
# 結果テーブル生成
# ==============================================================

def routes_to_dataframe(routes: list[AssignedRoute]) -> pd.DataFrame:
    rows = []
    for route in routes:
        driver_name = route.driver.name     if route.driver else "未定"
        priority    = route.driver.priority if route.driver else "-"
        for i, stop in enumerate(route.stops):
            h, m = divmod(stop["arrival_min"], 60)
            rows.append({
                "店舗":     route.shop,
                "車両名":   route.vehicle.name,
                "運転担当": driver_name,
                "優先度":   priority,
                "順番":     i + 1,
                "氏名":     stop["user"].name,
                "サービス": stop["user"].service_type.value,
                "住所":     stop["address"],
                "到着予定": f"{h:02d}:{m:02d}",
                "車椅子":   "♿" if stop["user"].wheelchair else "",
            })
    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ==============================================================
# 地図描画
# ==============================================================

SHOP_MAP_COLORS = ["blue", "green", "red", "purple", "orange"]

def render_map(routes: list[AssignedRoute]):
    if not FOLIUM_AVAILABLE:
        st.info("📦 `folium` / `streamlit-folium` をインストールすると地図が表示されます")
        return
    if not routes:
        return

    all_lats = [s["lat"] for r in routes for s in r.stops]
    all_lngs = [s["lng"] for r in routes for s in r.stops]
    center   = [sum(all_lats)/len(all_lats), sum(all_lngs)/len(all_lngs)]

    m = folium.Map(location=center, zoom_start=12, tiles="CartoDB positron")

    shops      = sorted(set(r.shop for r in routes))
    shop_color = {s: SHOP_MAP_COLORS[i % len(SHOP_MAP_COLORS)] for i, s in enumerate(shops)}

    for route in routes:
        color  = shop_color[route.shop]
        depot  = [route.vehicle.depot_lat, route.vehicle.depot_lng]
        coords = [depot] + [[s["lat"], s["lng"]] for s in route.stops] + [depot]

        folium.Marker(
            depot,
            tooltip=f"🏠 {route.shop} デポ",
            icon=folium.Icon(color=color, icon="home", prefix="fa")
        ).add_to(m)

        folium.PolyLine(
            coords, color=color, weight=3, opacity=0.75,
            tooltip=f"{route.shop} - {route.vehicle.name}"
        ).add_to(m)

        for i, stop in enumerate(route.stops):
            h, mn = divmod(stop["arrival_min"], 60)
            wc    = "♿ " if stop["user"].wheelchair else ""
            popup = (
                f"<b>{i+1}. {stop['user'].name}</b><br>"
                f"{wc}{stop['address']}<br>"
                f"🕐 {h:02d}:{mn:02d}<br>"
                f"🚗 {route.vehicle.name} ({route.shop})"
            )
            folium.CircleMarker(
                location=[stop["lat"], stop["lng"]],
                radius=8, color=color, fill=True, fill_opacity=0.85,
                tooltip=f"{i+1}. {stop['user'].name}",
                popup=folium.Popup(popup, max_width=220),
            ).add_to(m)

    st_folium(m, width=None, height=480, returned_objects=[])


# ==============================================================
# Streamlit UI メイン
# ==============================================================

def main():

    # ---- ヘッダー ----
    st.markdown("""
    <div class="main-header">
      <h1>🚌 送迎ルート最適化システム v2</h1>
      <p>放課後等デイサービス・就労継続支援A型/B型　｜　3店舗対応　｜　迎え・送り同時計算</p>
    </div>
    """, unsafe_allow_html=True)

    # ---- サイドバー ----
    with st.sidebar:
        st.markdown("### ⚙️ 実行設定")
        st.markdown("**迎え便 時刻設定**")
        c1, c2 = st.columns(2)
        with c1:
            pu_start_h = st.number_input("出発 時", 5,  12, 8,  key="pu_sh")
            pu_start_m = st.number_input("出発 分", 0,  55, 0,  step=5, key="pu_sm")
        with c2:
            pu_limit_h = st.number_input("リミット 時", 6,  13, 9,  key="pu_lh")
            pu_limit_m = st.number_input("リミット 分", 0,  55, 0,  step=5, key="pu_lm")

        st.markdown("**送り便 時刻設定**")
        c3, c4 = st.columns(2)
        with c3:
            do_start_h = st.number_input("出発 時", 15, 20, 17, key="do_sh")
            do_start_m = st.number_input("出発 分", 0,  55, 0,  step=5, key="do_sm")
        with c4:
            do_limit_h = st.number_input("目標到着 時", 16, 22, 19, key="do_lh")
            do_limit_m = st.number_input("目標到着 分", 0,  55, 0,  step=5, key="do_lm")

        pu_start = pu_start_h * 60 + pu_start_m
        pu_limit = pu_limit_h * 60 + pu_limit_m
        do_start = do_start_h * 60 + do_start_m
        do_limit = do_limit_h * 60 + do_limit_m

        st.divider()
        algo_label = "🤖 OR-Tools（高精度）" if ORTOOLS_AVAILABLE else "⚡ グリーディ（高速）"
        st.info(f"使用アルゴリズム:\n{algo_label}")

        st.divider()
        st.markdown("**サンプルExcelをダウンロード**")
        if OPENPYXL_AVAILABLE:
            st.download_button(
                "📥 サンプルExcel（v2対応）",
                data=get_sample_excel(),
                file_name="送迎マスタ_サンプルv2.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        st.caption("シート名: 利用者 / 車両 / スタッフ")

    # ==========================================================
    # STEP 1: データ読み込み
    # ==========================================================
    st.markdown('<div class="step-badge">STEP 1　データの読み込み</div>', unsafe_allow_html=True)

    col_up, col_demo = st.columns([2, 1])
    with col_up:
        uploaded = st.file_uploader(
            "Excelファイルをアップロード",
            type=["xlsx", "xls"],
            help="サイドバーからサンプルExcel(v2)をDLして記入してください",
        )
        if uploaded:
            try:
                with st.spinner("読み込み中..."):
                    users, vehicles, staff = parse_excel_upload(uploaded)
                st.session_state.update(
                    {"users": users, "vehicles": vehicles, "staff": staff}
                )
                st.success(
                    f"✅ 読み込み完了　利用者 {len(users)}名 /"
                    f" 車両 {len(vehicles)}台 / スタッフ {len(staff)}名"
                )
            except Exception as e:
                st.error(f"❌ 読み込み失敗: {e}")

    with col_demo:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🎯 デモデータで実行", use_container_width=True, type="secondary"):
            users, vehicles, staff = get_demo_data()
            st.session_state.update(
                {"users": users, "vehicles": vehicles, "staff": staff}
            )
            st.success(f"✅ デモデータ読み込み完了")

    if "users" not in st.session_state:
        st.info("👆 STEP 1 でデータを読み込んでください")
        return

    users    = st.session_state["users"]
    vehicles = st.session_state["vehicles"]
    staff    = st.session_state["staff"]

    # データプレビュー
    with st.expander("📋 読み込みデータのプレビュー", expanded=False):
        t1, t2, t3 = st.tabs(["👶 利用者", "🚗 車両", "👤 スタッフ"])
        with t1:
            st.dataframe(pd.DataFrame([{
                "ID": u.user_id, "氏名": u.name, "店舗": u.shop,
                "サービス": u.service_type.value,
                "住所": u.address,
                "車椅子": "♿" if u.wheelchair else "",
                "乗降時間(分)": u.service_time // 60,
                "同乗不可": ",".join(u.incompatible) or "なし",
            } for u in users]), use_container_width=True, hide_index=True)
        with t2:
            st.dataframe(pd.DataFrame([{
                "車両名": v.name, "店舗": v.shop,
                "種別": v.vehicle_type, "定員": v.capacity,
                "車椅子対応": "✅" if v.wheelchair_ok else "✗",
            } for v in vehicles]), use_container_width=True, hide_index=True)
        with t3:
            st.dataframe(pd.DataFrame([{
                "氏名": s.name, "店舗": s.shop,
                "運転可否": "✅" if s.can_drive else "❌",
                "優先度": s.priority,
            } for s in staff]), use_container_width=True, hide_index=True)

    st.divider()

    # ==========================================================
    # STEP 2: 当日欠席トグル（v2追加）
    # ==========================================================
    st.markdown('<div class="step-badge">STEP 2　当日の出欠確認</div>', unsafe_allow_html=True)
    st.markdown("#### 本日欠席の利用者のチェックを外してください")
    st.caption("✅ = 出席（送迎対象）　□ = 欠席（送迎対象外）")

    # st.data_editor 用DataFrameを構築
    shops_list = sorted(set(u.shop for u in users))
    attend_df  = pd.DataFrame([{
        "出席":     True,
        "店舗":     u.shop,
        "氏名":     u.name,
        "サービス": u.service_type.value,
        "住所":     u.address,
        "車椅子":   "♿" if u.wheelchair else "",
        "_user_id": u.user_id,          # 内部キー（非表示）
    } for u in sorted(users, key=lambda u: (u.shop, u.name))])

    edited_df = st.data_editor(
        attend_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "出席":    st.column_config.CheckboxColumn("出席", width="small"),
            "店舗":    st.column_config.TextColumn("店舗",    width="small"),
            "氏名":    st.column_config.TextColumn("氏名",    width="medium"),
            "サービス":st.column_config.TextColumn("サービス",width="medium"),
            "住所":    st.column_config.TextColumn("住所",    width="large"),
            "車椅子":  st.column_config.TextColumn("車椅子",  width="small"),
            "_user_id":st.column_config.TextColumn("_user_id", disabled=True),
        },
        disabled=["店舗", "氏名", "サービス", "住所", "車椅子", "_user_id"],
        column_order=["出席", "店舗", "氏名", "サービス", "住所", "車椅子"],
        key="attendance_editor",
    )

    # 出席者のみ抽出
    attending_ids = set(
        edited_df.loc[edited_df["出席"] == True, "_user_id"].tolist()
    )
    attending_users = [u for u in users if u.user_id in attending_ids]

    absent_count  = len(users) - len(attending_users)
    attend_count  = len(attending_users)

    col_a, col_b = st.columns(2)
    col_a.metric("本日 出席", f"{attend_count} 名")
    col_b.metric("本日 欠席", f"{absent_count} 名",
                 delta=f"-{absent_count}" if absent_count else None,
                 delta_color="inverse")

    st.divider()

    # ==========================================================
    # STEP 3: 最適化実行（v2: 迎え+送り両便同時計算）
    # ==========================================================
    st.markdown('<div class="step-badge">STEP 3　最適化の実行</div>', unsafe_allow_html=True)

    if attend_count == 0:
        st.warning("出席者が0名です。STEP 2 で出席者を選択してください。")
        return

    # 制約チェック（出席者のみ対象）
    checker = ConstraintChecker()
    errors  = checker.validate(attending_users, vehicles, staff)
    if errors:
        for e in errors:
            st.error(e)

    run_clicked = st.button(
        "🚀　迎え便・送り便を同時に最適化する",
        disabled=bool(errors) or attend_count == 0,
        type="primary",
        use_container_width=True,
    )

    if run_clicked:
        with st.spinner("🔄 迎え便・送り便のルートを同時最適化中..."):
            pickup_routes  = run_all_shops(
                attending_users, vehicles, staff,
                TripType.PICKUP, pu_start, pu_limit
            )
            dropoff_routes = run_all_shops(
                attending_users, vehicles, staff,
                TripType.DROPOFF, do_start, do_limit
            )

        st.session_state["pickup_routes"]  = pickup_routes
        st.session_state["dropoff_routes"] = dropoff_routes
        n_pu = len(pickup_routes)
        n_do = len(dropoff_routes)
        st.success(
            f"✅ 最適化完了！　迎え便: {n_pu} ルート　送り便: {n_do} ルート"
        )

    st.divider()

    # ==========================================================
    # STEP 4: 結果表示・ダウンロード（v2: 両便表示）
    # ==========================================================
    st.markdown('<div class="step-badge">STEP 4　結果の確認とダウンロード</div>', unsafe_allow_html=True)

    if "pickup_routes" not in st.session_state:
        st.info("👆 STEP 3 で最適化を実行すると、ここに結果が表示されます")
        return

    pickup_routes  = st.session_state["pickup_routes"]
    dropoff_routes = st.session_state["dropoff_routes"]

    all_routes = pickup_routes + dropoff_routes
    if not all_routes:
        st.error("ルートを生成できませんでした。データを確認してください。")
        return

    # ---- サマリーメトリクス ----
    total_pu = sum(len(r.stops) for r in pickup_routes)
    total_do = sum(len(r.stops) for r in dropoff_routes)
    wc_count = sum(1 for r in pickup_routes for s in r.stops if s["user"].wheelchair)
    shops_used = sorted(set(r.shop for r in all_routes))

    m1, m2, m3, m4 = st.columns(4)
    for col, (val, label) in zip(
        [m1, m2, m3, m4],
        [(total_pu,        "迎え便 割り当て数"),
         (total_do,        "送り便 割り当て数"),
         (wc_count,        "車椅子利用者"),
         (len(shops_used), "稼働店舗数")]
    ):
        with col:
            st.markdown(
                f'<div class="metric-box">'
                f'<div class="val">{val}</div>'
                f'<div class="label">{label}</div>'
                f'</div>',
                unsafe_allow_html=True
            )
    st.markdown("<br>", unsafe_allow_html=True)

    # ---- 制約検証 ----
    with st.expander("🔍 制約条件 検証サマリー", expanded=False):
        forbidden = checker.get_forbidden_pairs(attending_users)
        all_ok    = True

        for label, routes in [("迎え便", pickup_routes), ("送り便", dropoff_routes)]:
            st.markdown(f"**{label}**")
            for route in routes:
                users_in  = [s["user"] for s in route.stops]
                ok_cap    = len(users_in) <= route.vehicle.capacity
                ok_wc     = not (any(u.wheelchair for u in users_in) and not route.vehicle.wheelchair_ok)
                ok_incomp = not any(
                    tuple(sorted([u1.user_id, u2.user_id])) in forbidden
                    for i, u1 in enumerate(users_in) for u2 in users_in[i+1:]
                )
                ok_driver = route.driver is not None and route.driver.can_drive
                ok_shop   = all(u.shop == route.shop for u in users_in)
                all_ok    = all_ok and all([ok_cap, ok_wc, ok_incomp, ok_driver, ok_shop])

                stat = lambda ok: (
                    '<span class="constraint-ok">✅</span>' if ok
                    else '<span class="constraint-fail">❌ 違反</span>'
                )
                dn = route.driver.name if route.driver else "未定"
                st.markdown(
                    f"　**{route.shop} - {route.vehicle.name}** "
                    f"({len(users_in)}/{route.vehicle.capacity}名)　"
                    f"定員:{stat(ok_cap)}　車椅子:{stat(ok_wc)}　"
                    f"同乗不可:{stat(ok_incomp)}　"
                    f"混載禁止:{stat(ok_shop)}　"
                    f"運転者:{dn} {stat(ok_driver)}",
                    unsafe_allow_html=True,
                )
        if all_ok:
            st.success("🎉 全制約条件をクリアしています！")

    # ---- 結果テーブル（迎え・送り タブ切り替え）----
    st.markdown("#### 📋 送迎ルート一覧")
    tab_pu, tab_do = st.tabs(["▶ 迎え便", "◀ 送り便"])

    col_cfg = {
        "順番":    st.column_config.NumberColumn(width="small"),
        "到着予定":st.column_config.TextColumn(width="small"),
        "車椅子":  st.column_config.TextColumn(width="small"),
        "優先度":  st.column_config.NumberColumn(width="small"),
    }

    with tab_pu:
        df_pu = routes_to_dataframe(pickup_routes)
        if not df_pu.empty:
            st.dataframe(df_pu, use_container_width=True,
                         hide_index=True, column_config=col_cfg)
        else:
            st.info("迎え便のルートがありません")

    with tab_do:
        df_do = routes_to_dataframe(dropoff_routes)
        if not df_do.empty:
            st.dataframe(df_do, use_container_width=True,
                         hide_index=True, column_config=col_cfg)
        else:
            st.info("送り便のルートがありません")

    # ---- ダウンロードボタン ----
    st.markdown("<br>", unsafe_allow_html=True)
    col_dl, _ = st.columns([1, 2])
    with col_dl:
        excel_bytes = build_excel_output(pickup_routes, dropoff_routes)
        ext  = "xlsx" if OPENPYXL_AVAILABLE else "csv"
        mime = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                if OPENPYXL_AVAILABLE else "text/csv")
        st.download_button(
            label=f"📥 送迎ルート表をダウンロード (.{ext})",
            data=excel_bytes,
            file_name=f"送迎ルート_迎え送り.{ext}",
            mime=mime,
            type="primary",
            use_container_width=True,
        )
        st.caption("迎え便・送り便が1ファイルの2シートで出力されます")

    st.divider()

    # ---- 地図表示（全店舗・両便重ねて表示）----
    st.markdown("#### 🗺️ 送迎ルートマップ（迎え便）")
    render_map(pickup_routes)

    if dropoff_routes:
        with st.expander("◀ 送り便マップも表示する", expanded=False):
            render_map(dropoff_routes)


if __name__ == "__main__":
    main()
