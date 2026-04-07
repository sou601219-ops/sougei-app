"""
送迎ルート自動作成システム - Streamlit Webアプリ v3
=====================================================
放課後等デイサービス / 就労継続支援A型・B型 対応

v4 追加要件:
  【第1】時刻入力の HH:MM 形式化 + 個別Time Windows の厳格適用
  【第2】スタッフシフト表対応 + 稼働時間のVRP制約化
  【第3】月間予定表（カレンダーシート）＋ 対象日選択 UI

v2 継続機能（デグレなし）:
  - 迎え・送り 1回実行・2シートExcel出力
  - 店舗ごとのVRP分離（混載禁止）
  - スタッフ優先度 Fixed Cost による負荷分散
  - 乗降時間（Service Time）の時間ディメンション加算
  - 当日欠席トグル UI（st.data_editor）
"""

from __future__ import annotations

import io
import math
import datetime
from dataclasses import dataclass, field
from typing import Optional
from enum import Enum

import pandas as pd
import streamlit as st

# ---- オプションライブラリ ----
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
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, GradientFill
    )
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ==============================================================
# ページ設定（必ず先頭）
# ==============================================================
st.set_page_config(
    page_title="送迎ルート最適化 v5",
    page_icon="🚌",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ==============================================================
# テーマ定数
# ==============================================================
THEME = {
    "primary":      "#1B3A5C",
    "primary_lt":   "#2D5F8A",
    "accent":       "#2E7D52",
    "accent_lt":    "#E8F5EE",
    "warn":         "#C0392B",
    "warn_lt":      "#FDECEA",
    "surface":      "#FFFFFF",
    "surface2":     "#F6F8FA",
    "border":       "#DDE1E7",
    "text":         "#1C2330",
    "text2":        "#5A6478",
    "shopA":        "#1A5276",
    "shopA_lt":     "#D4E6F1",
    "shopB":        "#145A32",
    "shopB_lt":     "#D5F5E3",
    "shopC":        "#6E2F1A",
    "shopC_lt":     "#FDEBD0",
    "shopD":        "#4A235A",
    "shopD_lt":     "#F3E6FA",
}

SHOP_COLORS = [
    (THEME["shopA"], THEME["shopA_lt"]),
    (THEME["shopB"], THEME["shopB_lt"]),
    (THEME["shopC"], THEME["shopC_lt"]),
    (THEME["shopD"], THEME["shopD_lt"]),
]

# ==============================================================
# カスタム CSS
# ==============================================================
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@300;400;500;700;900&family=DM+Mono:wght@400;500&display=swap');

  html, body, [class*="css"] {{
    font-family: 'Noto Sans JP', 'メイリオ', sans-serif;
  }}

  /* ヘッダー */
  .hero {{
    background: linear-gradient(135deg, {THEME["primary"]} 0%, {THEME["primary_lt"]} 60%, #3A7BD5 100%);
    border-radius: 16px;
    padding: 28px 32px;
    margin-bottom: 28px;
    color: white;
    position: relative;
    overflow: hidden;
  }}
  .hero::before {{
    content: "🚌";
    position: absolute;
    right: 24px;
    top: 50%;
    transform: translateY(-50%);
    font-size: 80px;
    opacity: 0.12;
  }}
  .hero h1  {{ font-size: 26px; margin: 0; font-weight: 900; letter-spacing: 0.03em; }}
  .hero p   {{ font-size: 13px; margin: 8px 0 0; opacity: 0.75; line-height: 1.7; }}
  .hero .v-badge {{
    display: inline-block;
    background: rgba(255,255,255,0.25);
    border: 1px solid rgba(255,255,255,0.4);
    border-radius: 20px;
    padding: 2px 12px;
    font-size: 11px;
    font-weight: 700;
    letter-spacing: 0.08em;
    margin-bottom: 10px;
  }}

  /* ステップ見出し */
  .step-header {{
    display: flex;
    align-items: center;
    gap: 12px;
    margin: 28px 0 16px;
  }}
  .step-num {{
    width: 36px; height: 36px;
    background: {THEME["primary"]};
    color: white;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 15px; font-weight: 900;
    flex-shrink: 0;
    font-family: 'DM Mono', monospace;
  }}
  .step-title {{
    font-size: 17px;
    font-weight: 700;
    color: {THEME["text"]};
  }}
  .step-sub {{
    font-size: 12px;
    color: {THEME["text2"]};
    margin-top: 2px;
  }}

  /* カード */
  .info-card {{
    background: {THEME["surface"]};
    border: 1px solid {THEME["border"]};
    border-radius: 12px;
    padding: 18px 22px;
    margin-bottom: 16px;
    box-shadow: 0 1px 6px rgba(0,0,0,0.06);
  }}
  .info-card-accent {{
    border-left: 4px solid {THEME["accent"]};
  }}

  /* 店舗バッジ */
  .shop-badge {{
    display: inline-block;
    padding: 3px 10px;
    border-radius: 20px;
    font-size: 11px;
    font-weight: 700;
    margin-right: 4px;
  }}

  /* メトリクス */
  .metric-grid {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin: 16px 0;
  }}
  .metric-item {{
    background: {THEME["surface"]};
    border: 1px solid {THEME["border"]};
    border-radius: 10px;
    padding: 16px 12px;
    text-align: center;
  }}
  .metric-item .val   {{ font-size: 32px; font-weight: 900; color: {THEME["primary"]}; line-height: 1; }}
  .metric-item .unit  {{ font-size: 13px; font-weight: 600; color: {THEME["primary"]}; }}
  .metric-item .label {{ font-size: 11px; color: {THEME["text2"]}; margin-top: 4px; }}

  /* 制約バッジ */
  .ok   {{ color: #27AE60; font-weight: 700; }}
  .fail {{ color: #E74C3C; font-weight: 700; }}

  /* カレンダー関連 */
  .cal-date-badge {{
    background: {THEME["accent"]};
    color: white;
    border-radius: 8px;
    padding: 4px 14px;
    font-size: 14px;
    font-weight: 700;
    font-family: 'DM Mono', monospace;
    display: inline-block;
    margin-bottom: 8px;
  }}

  /* タイムライン */
  .timeline-item {{
    display: flex;
    align-items: flex-start;
    gap: 12px;
    padding: 10px 0;
    border-bottom: 1px dashed {THEME["border"]};
  }}
  .timeline-item:last-child {{ border-bottom: none; }}
  .timeline-dot {{
    width: 10px; height: 10px;
    border-radius: 50%;
    background: {THEME["accent"]};
    margin-top: 5px;
    flex-shrink: 0;
  }}
  .timeline-time {{
    font-family: 'DM Mono', monospace;
    font-size: 13px;
    font-weight: 600;
    color: {THEME["accent"]};
    width: 50px;
    flex-shrink: 0;
  }}
  .timeline-name {{ font-size: 13px; font-weight: 600; }}
  .timeline-detail {{ font-size: 11px; color: {THEME["text2"]}; }}

  /* ダウンロードボタン */
  .stDownloadButton > button {{
    border-radius: 10px !important;
    font-weight: 700 !important;
    background: {THEME["accent"]} !important;
    color: white !important;
    border: none !important;
    padding: 10px 24px !important;
    font-size: 14px !important;
    transition: all 0.2s !important;
  }}
  .stDownloadButton > button:hover {{
    background: #2E9D60 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 16px rgba(46,125,82,0.35) !important;
  }}

  /* 実行ボタン */
  .stButton > button[kind="primary"] {{
    background: linear-gradient(135deg, {THEME["primary"]} 0%, {THEME["primary_lt"]} 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    font-size: 15px !important;
    padding: 14px 32px !important;
    transition: all 0.2s !important;
  }}
  .stButton > button[kind="primary"]:hover {{
    transform: translateY(-2px) !important;
    box-shadow: 0 6px 20px rgba(27,58,92,0.35) !important;
  }}
  .stButton > button[kind="secondary"] {{
    border-radius: 8px !important;
    font-weight: 600 !important;
  }}

  /* サイドバー */
  section[data-testid="stSidebar"] {{
    background: {THEME["surface2"]};
  }}
  .sidebar-section {{
    background: white;
    border-radius: 10px;
    padding: 14px 16px;
    margin-bottom: 12px;
    border: 1px solid {THEME["border"]};
  }}
  .sidebar-section h4 {{
    font-size: 13px;
    font-weight: 700;
    color: {THEME["primary"]};
    margin: 0 0 10px;
    padding-bottom: 6px;
    border-bottom: 1px solid {THEME["border"]};
  }}

  /* アニメーション */
  @keyframes fadeUp {{
    from {{ opacity: 0; transform: translateY(10px); }}
    to   {{ opacity: 1; transform: translateY(0); }}
  }}
  .fade-up {{ animation: fadeUp 0.3s ease; }}

  /* 印刷 */
  @media print {{
    header, section[data-testid="stSidebar"],
    .stButton, .stDownloadButton, .step-header {{ display: none !important; }}
    .stDataFrame {{ font-size: 10px; }}
  }}
</style>
""", unsafe_allow_html=True)


# ==============================================================
# ユーティリティ関数
# ==============================================================

def hhmm_to_min(s: str, default: int = 540) -> int:
    """
    【第1改修】HH:MM 形式の文字列を分数（整数）に変換。
    例: "09:00" → 540、"15:30" → 930
    変換失敗時は default を返す。
    """
    if not s or str(s).strip() in ("", "nan", "None"):
        return default
    s = str(s).strip()
    try:
        parts = s.split(":")
        return int(parts[0]) * 60 + int(parts[1])
    except (ValueError, IndexError):
        try:
            # 数値で渡された場合（互換性）
            return int(float(s))
        except ValueError:
            return default


def min_to_hhmm(m: int) -> str:
    """分数を HH:MM 文字列に変換。例: 540 → '09:00'"""
    h, mn = divmod(m, 60)
    return f"{h:02d}:{mn:02d}"


def step_header(num: int, title: str, sub: str = ""):
    """ステップヘッダーを描画"""
    st.markdown(f"""
    <div class="step-header fade-up">
      <div class="step-num">{num}</div>
      <div>
        <div class="step-title">{title}</div>
        {"<div class='step-sub'>" + sub + "</div>" if sub else ""}
      </div>
    </div>
    """, unsafe_allow_html=True)


def metric_row(items: list[tuple]):
    """
    items: [(value, unit, label), ...]
    """
    cols = "".join([
        f'<div class="metric-item">'
        f'<div class="val">{v}</div>'
        f'<div class="unit">{u}</div>'
        f'<div class="label">{l}</div>'
        f'</div>'
        for v, u, l in items
    ])
    st.markdown(f'<div class="metric-grid">{cols}</div>', unsafe_allow_html=True)


# ==============================================================
# データクラス
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
    shop:           str
    wheelchair:     bool  = False
    incompatible:   list  = field(default_factory=list)
    # 【第1改修】分単位で保持（HH:MM から変換済み）
    pickup_latest:  int   = 540     # 迎え便 施設到着リミット（分）
    dropoff_target: int   = 1050    # 送り便 自宅到着目標（分）
    # 乗降時間（秒）: 通常 300秒(5分)、車椅子 600秒(10分)
    service_time:   int   = 300

    def __post_init__(self):
        if self.wheelchair:
            self.service_time = 600


@dataclass
class Vehicle:
    vehicle_id:    str
    name:          str
    vehicle_type:  str
    capacity:      int
    shop:          str
    wheelchair_ok: bool  = False
    depot_lat:     float = 36.6953
    depot_lng:     float = 137.2113


@dataclass
class Staff:
    staff_id:    str
    name:        str
    shop:        str
    can_drive:   bool = True
    priority:    int  = 1
    # 【第2改修】シフト稼働時間（分）。None = 終日
    shift_start: Optional[int] = None
    shift_end:   Optional[int] = None


@dataclass
class AssignedRoute:
    vehicle:   Vehicle
    driver:    Optional[Staff]
    trip_type: TripType
    shop:      str
    stops:     list
    total_min: int


# ==============================================================
# 優先度コスト（v2継続）
# ==============================================================
PRIORITY_COST_MAP = {
    1: 0,
    2: 5_000,
    3: 15_000,
    4: 30_000,
    9: 999_999,
}


# ==============================================================
# 距離行列ビルダー
# ==============================================================

class DistanceMatrixBuilder:
    def build(self, locations: list[tuple[float, float]]) -> list[list[int]]:
        n = len(locations)
        return [
            [self._h(locations[i], locations[j]) for j in range(n)]
            for i in range(n)
        ]

    @staticmethod
    def _h(p1, p2, spd: float = 30.0) -> int:
        R = 6371.0
        la1, lo1 = math.radians(p1[0]), math.radians(p1[1])
        la2, lo2 = math.radians(p2[0]), math.radians(p2[1])
        dla, dlo = la2 - la1, lo2 - lo1
        a = math.sin(dla/2)**2 + math.cos(la1)*math.cos(la2)*math.sin(dlo/2)**2
        return max(1, int(2 * R * math.asin(math.sqrt(a)) / spd * 3600))


# ==============================================================
# 制約チェッカー
# ==============================================================

class ConstraintChecker:
    @staticmethod
    def validate(users, vehicles, staff) -> list[str]:
        errors = []
        if not [s for s in staff if s.can_drive]:
            errors.append("❌ 運転可能なスタッフが0人です")
        wcu = [u for u in users if u.wheelchair]
        wcv = [v for v in vehicles if v.wheelchair_ok]
        if wcu and not wcv:
            errors.append(f"❌ 車椅子利用者 {[u.name for u in wcu]} がいますが車椅子対応車両がありません")
        if len(users) > sum(v.capacity for v in vehicles):
            errors.append(f"❌ 利用者数({len(users)}名) > 全車両定員合計({sum(v.capacity for v in vehicles)}名)")
        return errors

    @staticmethod
    def get_forbidden_pairs(users) -> set[tuple[str, str]]:
        pairs = set()
        for u in users:
            for iid in u.incompatible:
                pairs.add(tuple(sorted([u.user_id, iid])))
        return pairs


# ==============================================================
# VRP ソルバー（v3: 個別TimeWindow + シフト稼働時間）
# ==============================================================

class TransportVRPSolver:
    """
    店舗単位で独立して呼び出されるVRPソルバー。

    v3 変更点:
      - 個別 Time Windows（利用者ごとの pickup_latest / dropoff_target）を
        CumulVar.SetMax / SetMin で Hard Constraint として適用
      - スタッフのシフト時間を Vehicle の稼働時間上限として制約
    """

    TIME_LIMIT_SEC = 30

    def __init__(
        self,
        users:                   list[User],
        vehicles:                list[Vehicle],
        staff:                   list[Staff],
        distance_matrix:         list[list[int]],
        trip_type:               TripType = TripType.PICKUP,
        depot_arrival_limit_min: int = 540,   # 全体デフォルト（個別設定がない場合のみ使用）
        start_time_min:          int = 480,
    ):
        self.users     = users
        self.vehicles  = vehicles
        self.staff     = sorted(
            [s for s in staff if s.can_drive],
            key=lambda s: s.priority
        )
        self.matrix    = distance_matrix
        self.trip_type = trip_type
        # 全体デフォルト（秒換算）
        self.global_limit  = depot_arrival_limit_min * 60
        self.start_time    = start_time_min * 60

        checker = ConstraintChecker()
        self.forbidden_pairs = checker.get_forbidden_pairs(users)

    def solve(self) -> list[AssignedRoute]:
        if ORTOOLS_AVAILABLE and self.users:
            result = self._solve_with_ortools()
            if result is not None:
                return result
        return self._greedy_fallback()

    # ----------------------------------------------------------
    # OR-Tools 本体（v3: 個別TimeWindow + シフト制約）
    # ----------------------------------------------------------
    def _solve_with_ortools(self) -> Optional[list[AssignedRoute]]:
        nu = len(self.users)
        nv = len(self.vehicles)
        if nv == 0 or nu == 0:
            return []

        # ノード: 0=デポ, 1..nu=各利用者
        nn = nu + 1

        manager = pywrapcp.RoutingIndexManager(nn, nv, 0)
        routing = pywrapcp.RoutingModel(manager)

        # -------- コールバック: 移動時間 + 乗降時間(service_time) --------
        # 【v2継続】乗降時間は from_node が利用者の場合に加算
        def time_callback(fi, ti):
            fn = manager.IndexToNode(fi)
            tn = manager.IndexToNode(ti)
            travel = self.matrix[fn][tn]
            # from_node が利用者ノード（0以外）なら乗降時間を加算
            svc = self.users[fn - 1].service_time if fn > 0 else 0
            return travel + svc

        transit_cb = routing.RegisterTransitCallback(time_callback)
        routing.SetArcCostEvaluatorOfAllVehicles(transit_cb)

        # -------- Time Dimension --------
        # 上限は全体リミットと個別TimeWindowの最大値を包括できる値に設定
        max_tw = max(
            self.global_limit,
            *[u.pickup_latest  * 60 for u in self.users],
            *[u.dropoff_target * 60 for u in self.users],
            self.start_time + 6 * 3600,
        )
        routing.AddDimension(
            transit_cb,
            600,       # 待機許容スラック（秒）
            max_tw,    # ディメンション最大値
            True,      # 始点を 0 に固定
            "Time",
        )
        time_dim = routing.GetDimensionOrDie("Time")

        # -------- 【第1改修】個別 Time Windows を Hard Constraint で適用 --------
        # 各利用者ノードに「その利用者専用の到着リミット」を SetMax で設定する。
        # これにより同一車両内でも「Aさんは9時まで、Bさんは10時まで」が機能する。
        for i, user in enumerate(self.users):
            node_idx = manager.NodeToIndex(i + 1)  # ノード0はデポ、利用者は1始まり

            if self.trip_type == TripType.PICKUP:
                # 迎え便: 施設到着リミット（pickup_latest）を SetMax で Hard Constraint 化
                # pickup_latest は絶対時刻（分）なので、始点(start_time)からの相対秒に変換
                limit_sec = (user.pickup_latest * 60) - self.start_time
                if limit_sec > 0:
                    # CumulVar はノード到着時の「出発からの経過時間（秒）」を表す
                    time_dim.CumulVar(node_idx).SetMax(limit_sec)

            else:
                # 送り便: 自宅到着目標（dropoff_target）を SetMax で適用（Soft→Hard）
                # 送り便は遅延許容だが、目標+1時間をハードリミットとして設定
                soft_limit_sec = (user.dropoff_target * 60) - self.start_time
                hard_limit_sec = soft_limit_sec + 3600  # 1時間の余裕
                if hard_limit_sec > 0:
                    time_dim.CumulVar(node_idx).SetMax(hard_limit_sec)

        # デポ（ノード0）への帰着リミットは全体設定を使用
        depot_idx = manager.NodeToIndex(0)
        if self.trip_type == TripType.PICKUP:
            time_dim.CumulVar(depot_idx).SetMax(self.global_limit)

        # -------- 【第2改修】Vehicle ごとのシフト稼働時間を制約として設定 --------
        # スタッフのシフト開始・終了時間を Vehicle の TimeWindow として設定する。
        # シフト開始 → 車両の出発可能最早時刻（CumulVar Start の SetMin）
        # シフト終了 → 車両の帰着最遅時刻（CumulVar End の SetMax）
        for vi, vehicle in enumerate(self.vehicles):
            driver = self._get_driver(vi)
            if driver is None:
                continue

            # シフト開始時刻（秒）: start_time からの相対値に変換
            if driver.shift_start is not None:
                # 車両の出発時刻が shift_start 以降になるよう制約
                shift_start_rel = max(0, (driver.shift_start * 60) - self.start_time)
                # Start ノードの CumulVar に SetMin でシフト開始を適用
                time_dim.CumulVar(routing.Start(vi)).SetMin(shift_start_rel)

            # シフト終了時刻（秒）: start_time からの相対値に変換
            if driver.shift_end is not None:
                # 車両が帰着（End ノード到達）するのが shift_end 以内であるよう制約
                shift_end_rel = (driver.shift_end * 60) - self.start_time
                if shift_end_rel > 0:
                    # End ノードの CumulVar に SetMax でシフト終了を適用
                    time_dim.CumulVar(routing.End(vi)).SetMax(shift_end_rel)

        # -------- 【v2継続】定員制約 --------
        def demand_cb(fi):
            return 0 if manager.IndexToNode(fi) == 0 else 1

        demand_idx = routing.RegisterUnaryTransitCallback(demand_cb)
        routing.AddDimensionWithVehicleCapacity(
            demand_idx,
            0,
            [v.capacity for v in self.vehicles],
            True,
            "Capacity",
        )

        # -------- 【v2継続】Fixed Cost: 優先度が低い車両ほど高コスト --------
        for vi in range(len(self.vehicles)):
            driver = self._get_driver(vi)
            priority = driver.priority if driver else 9
            routing.SetFixedCostOfVehicle(
                PRIORITY_COST_MAP.get(priority, 999_999), vi
            )

        # -------- 【v2継続】車椅子制約 --------
        for i, u in enumerate(self.users):
            if u.wheelchair:
                ni = manager.NodeToIndex(i + 1)
                for vi, v in enumerate(self.vehicles):
                    if not v.wheelchair_ok:
                        routing.VehicleVar(ni).RemoveValue(vi)

        # -------- 【v2継続】同乗不可制約 --------
        for uid1, uid2 in self.forbidden_pairs:
            i1 = next((i+1 for i, u in enumerate(self.users) if u.user_id == uid1), None)
            i2 = next((i+1 for i, u in enumerate(self.users) if u.user_id == uid2), None)
            if i1 and i2:
                ni1 = manager.NodeToIndex(i1)
                ni2 = manager.NodeToIndex(i2)
                routing.solver().Add(
                    routing.VehicleVar(ni1) != routing.VehicleVar(ni2)
                )

        # -------- 探索パラメータ --------
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

            driver = self._get_driver(vi)
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
    # グリーディフォールバック（OR-Tools 未インストール時）
    # ----------------------------------------------------------
    def _greedy_fallback(self) -> list[AssignedRoute]:
        vs_sorted = sorted(
            self.vehicles,
            key=lambda v: (0 if v.wheelchair_ok else 1, -v.capacity)
        )
        routes     = []
        unassigned = list(self.users)

        for vi, vehicle in enumerate(vs_sorted):
            if not unassigned:
                break
            driver   = self._get_driver(vi)
            assigned = []

            if vehicle.wheelchair_ok:
                for u in [u for u in unassigned if u.wheelchair]:
                    if len(assigned) < vehicle.capacity:
                        assigned.append(u)
                        unassigned.remove(u)

            fids = set()
            for a in assigned:
                fids.update(a.incompatible)

            for u in list(unassigned):
                if len(assigned) >= vehicle.capacity:
                    break
                if u.wheelchair and not vehicle.wheelchair_ok:
                    continue
                if u.user_id in fids:
                    continue
                assigned.append(u)
                unassigned.remove(u)
                fids.update(u.incompatible)

            if not assigned:
                continue

            ordered  = self._nn(assigned)
            stops    = []
            cur_node = 0
            cur_time = self.start_time

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
            st.warning(
                f"⚠️ [{self.vehicles[0].shop if self.vehicles else ''}] "
                f"割り当て不可: {[u.name for u in unassigned]}"
            )
        return routes

    def _nn(self, users: list[User]) -> list[User]:
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

    def _get_driver(self, vi: int) -> Optional[Staff]:
        return self.staff[vi % len(self.staff)] if self.staff else None


# ==============================================================
# 店舗別VRP実行（v2継続 + 第2改修: シフトフィルタリング）
# ==============================================================

def run_all_shops(
    users:     list[User],
    vehicles:  list[Vehicle],
    staff:     list[Staff],
    trip_type: TripType,
    start_min: int,
    limit_min: int,
) -> list[AssignedRoute]:
    """
    店舗ごとにデータを分割し、独立したVRPを実行してマージして返す。

    【第2改修】シフトフィルタリング:
    その便の時間帯に勤務していないスタッフを事前に除外することで
    VRPの計算量を削減し、不正な配車を防止する。
    """
    shops  = sorted(set(u.shop for u in users))
    routes = []

    for shop in shops:
        su = [u for u in users   if u.shop == shop]
        sv = [v for v in vehicles if v.shop == shop]

        # 【第2改修】シフト時間でスタッフをフィルタリング
        # - can_drive=False は除外
        # - shift_start / shift_end が設定されている場合、
        #   その便の出発時間帯に勤務しているスタッフのみ残す
        def is_on_shift(s: Staff) -> bool:
            if not s.can_drive:
                return False
            if s.shift_start is not None and s.shift_end is not None:
                # 出発時刻（start_min）がシフト内かチェック
                return s.shift_start <= start_min < s.shift_end
            if s.shift_start is not None:
                return s.shift_start <= start_min
            if s.shift_end is not None:
                return start_min < s.shift_end
            return True  # シフト未設定 = 終日勤務

        ss = [s for s in staff if s.shop == shop and is_on_shift(s)]

        if not su or not sv:
            continue
        if not ss:
            st.warning(f"⚠️ [{shop}] 該当時間帯に勤務可能なスタッフがいません")
            continue

        depot  = (sv[0].depot_lat, sv[0].depot_lng)
        locs   = [depot] + [(u.lat, u.lng) for u in su]
        matrix = DistanceMatrixBuilder().build(locs)

        solver = TransportVRPSolver(
            users                    = su,
            vehicles                 = sv,
            staff                    = ss,
            distance_matrix          = matrix,
            trip_type                = trip_type,
            depot_arrival_limit_min  = limit_min,
            start_time_min           = start_min,
        )
        routes.extend(solver.solve())

    return routes


# ==============================================================
# Excel 入出力（v3: HH:MM形式 + シフト + カレンダー対応）
# ==============================================================




# ==============================================================
# カレンダーパース ユーティリティ（v4新規）
# ==============================================================

def parse_time_range(cell_val, default_start: int = 480, default_end: int = 1140):
    """
    カレンダーセルの値を解析して (start_min, end_min) のタプルを返す。

    対応フォーマット:
      "09:00-15:30"  → (540, 930)   ← 基本形式（時間直接入力）
      "09:00～15:30" → (540, 930)   ← 全角ハイフン対応
      "09:00"        → (540, default_end) ← 開始時刻のみ
      "〇" / "○"     → (default_start, default_end) ← フォールバック
      空欄 / nan     → None  ← 欠席・休み

    戻り値:
      (start_min, end_min) : 出席/出勤の場合
      None                 : 欠席/休みの場合
    """
    if cell_val is None:
        return None
    s = str(cell_val).strip()
    if s in ("", "nan", "None", "-", "×", "✕", "欠席", "休み", "休"):
        return None

    # 〇 / ○ フォールバック
    if s in ("〇", "○", "◯", "✓", "✔", "出", "◎"):
        return (default_start, default_end)

    # "HH:MM-HH:MM" 形式（半角・全角ハイフン対応）
    for sep in ["-", "～", "〜", "~", "ー", "−"]:
        if sep in s:
            parts = s.split(sep, 1)
            if len(parts) == 2:
                try:
                    st = hhmm_to_min(parts[0].strip(), -1)
                    en = hhmm_to_min(parts[1].strip(), -1)
                    if st >= 0 and en >= 0:
                        return (st, en)
                except Exception:
                    pass

    # "HH:MM" 単体（開始時刻のみ）
    try:
        st = hhmm_to_min(s, -1)
        if st >= 0:
            return (st, default_end)
    except Exception:
        pass

    return None  # 解析失敗 = 欠席扱い


# ==============================================================
# デモデータ（v4）
# ==============================================================

def get_demo_data() -> tuple[list[User], list[Vehicle], list[Staff]]:
    """デモ用マスタデータ"""
    users = [
        User("u1",  "山田 太郎",   "富山市上袋100",   36.720, 137.210, ServiceType.HOUKAGO_DEI, "A店", False, [],      540,  1050),
        User("u2",  "鈴木 花子",   "富山市堀川200",   36.695, 137.220, ServiceType.HOUKAGO_DEI, "A店", False, ["u3"], 570,  1080),
        User("u3",  "田中 一郎",   "富山市婦中300",   36.660, 137.160, ServiceType.A_TYPE,      "A店", True,  ["u2"], 540,  1050),
        User("u4",  "佐藤 愛",     "富山市大沢野400", 36.630, 137.230, ServiceType.B_TYPE,      "A店", False, [],      540,  1050),
        User("u5",  "高橋 健太",   "富山市八尾500",   36.590, 137.270, ServiceType.HOUKAGO_DEI, "A店", False, [],      540,  1050),
        User("u6",  "渡辺 さくら", "富山市上袋600",   36.725, 137.215, ServiceType.B_TYPE,      "B店", False, [],      540,  1050),
        User("u7",  "伊藤 翔",     "富山市堀川700",   36.700, 137.225, ServiceType.A_TYPE,      "B店", False, [],      540,  1050),
        User("u8",  "中村 みな",   "富山市婦中800",   36.655, 137.155, ServiceType.B_TYPE,      "B店", False, [],      600,  1110),
        User("u9",  "小林 大輝",   "富山市大沢野900", 36.625, 137.235, ServiceType.HOUKAGO_DEI, "B店", False, [],      540,  1050),
        User("u10", "加藤 りん",   "富山市八尾1000",  36.585, 137.265, ServiceType.A_TYPE,      "B店", False, [],      540,  1050),
        User("u11", "中島 陽斗",   "富山市上袋1100",  36.715, 137.205, ServiceType.HOUKAGO_DEI, "C店", False, [],      540,  1050),
        User("u12", "斉藤 みゆ",   "富山市堀川1200",  36.690, 137.215, ServiceType.B_TYPE,      "C店", True,  [],      570,  1080),
    ]
    for u in users:
        u.service_time = 600 if u.wheelchair else 300

    vehicles = [
        Vehicle("v1", "A-1号車（大型）", "large",  7, "A店", True,  36.695, 137.211),
        Vehicle("v2", "A-2号車（普通）", "normal", 4, "A店", False, 36.695, 137.211),
        Vehicle("v3", "B-1号車（大型）", "large",  7, "B店", False, 36.710, 137.200),
        Vehicle("v4", "B-2号車（普通）", "normal", 4, "B店", False, 36.710, 137.200),
        Vehicle("v5", "C-1号車（大型）", "large",  7, "C店", True,  36.680, 137.220),
    ]
    staff = [
        Staff("s1", "林 誠一",   "A店", True,  1, 480,  1140),
        Staff("s2", "森 美咲",   "A店", True,  2, 480,  1020),
        Staff("s3", "池田 裕二", "A店", False, 9, None, None),
        Staff("s4", "宇野 幸子", "B店", True,  1, 480,  1140),
        Staff("s5", "川口 拓也", "B店", True,  2, 540,  1080),
        Staff("s6", "高木 雄介", "C店", True,  1, 480,  1140),
        Staff("s7", "中島 奈々", "C店", True,  3, 600,  1200),
    ]
    return users, vehicles, staff


# ==============================================================
# Excel カレンダー 書き込み（v4: プロ仕様レイアウト）
# ==============================================================

def build_demo_calendar(
    users: list[User],
    staff: list[Staff],
    year:  int = 2026,
    month: int = 4,
) -> dict:
    """
    デモ用の月間カレンダーデータを生成する。
    get_sample_excel と同じ「時間直接入力」形式でメモリ上に構築し、
    extract_for_date に渡せる dict 形式で返す。

    戻り値フォーマット:
      {
        "staff": { スタッフ名: { "YYYY-MM-DD": (start_min, end_min) or None } },
        "users": { 利用者名:   { "YYYY-MM-DD": (start_min, end_min) or None } },
        "year": year, "month": month
      }
    """
    import calendar as cal_mod
    _, days_in_month = cal_mod.monthrange(year, month)
    dates = [datetime.date(year, month, d) for d in range(1, days_in_month + 1)]

    staff_cal: dict = {}
    for s in staff:
        schedule: dict = {}
        for dt in dates:
            is_sun = dt.weekday() == 6
            if is_sun or not s.can_drive:
                schedule[dt.strftime("%Y-%m-%d")] = None  # 日曜・運転不可は休み
            else:
                # シフト時間をデフォルト値として設定
                ss = s.shift_start if s.shift_start is not None else 480
                se = s.shift_end   if s.shift_end   is not None else 1140
                schedule[dt.strftime("%Y-%m-%d")] = (ss, se)
        staff_cal[s.name] = schedule

    users_cal: dict = {}
    for u in users:
        schedule = {}
        for dt in dates:
            is_sat = dt.weekday() == 5
            is_sun = dt.weekday() == 6
            if is_sat or is_sun:
                schedule[dt.strftime("%Y-%m-%d")] = None  # 土日は休み
            else:
                # 迎え: 08:00 出発 → pickup_latest まで
                schedule[dt.strftime("%Y-%m-%d")] = (480, u.pickup_latest)
        users_cal[u.name] = schedule

    return {
        "staff":  staff_cal,
        "users":  users_cal,
        "year":   year,
        "month":  month,
    }




# ==============================================================
# v5 定数: 固定枠定義
# ==============================================================

# 各店舗のスタッフ・利用者の固定枠数
# 現場で増員があっても parse_excel_upload は動的に読み込むので
# これはあくまで「テンプレートに最初から用意する行数」
SHOP_LIST = ["A店", "B店", "C店"]

STAFF_FRAMES: dict[str, int] = {"A店": 10, "B店": 6, "C店": 6}
USER_FRAMES:  dict[str, int] = {"A店": 30, "B店": 30, "C店": 30}

# マスタシートの列定義（1-indexed, openpyxl）
# スタッフ: A=ID B=氏名 C=店舗 D=運転可否 E=優先度 F=出勤時間 G=退勤時間
STAFF_COL_NAME      = "B"
STAFF_COL_SHOP      = "C"
STAFF_COL_SHIFT_ST  = "F"   # 出勤時間
STAFF_COL_SHIFT_EN  = "G"   # 退勤時間

# 利用者: A=ID B=氏名 C=住所 D=緯度 E=経度 F=サービス種別 G=店舗
#          H=車椅子 I=同乗不可ID J=到着リミット K=送り目標
USER_COL_NAME   = "B"
USER_COL_LIMIT  = "J"   # 到着リミット


def _get_master_row_ranges(frames: dict[str, int]) -> dict[str, tuple[int, int]]:
    """
    マスタシートの各店舗の Excel 行範囲（1-indexed）を返す。
    Row 1 = タイトル, Row 2 = ヘッダー, Row 3 以降 = データ。

    例 (STAFF_FRAMES):
      A店 → (3, 12)   ← 10行
      B店 → (13, 18)  ← 6行
      C店 → (19, 24)  ← 6行
    """
    ranges: dict[str, tuple[int, int]] = {}
    cur = 3  # データ開始行
    for shop in SHOP_LIST:
        n = frames.get(shop, 0)
        ranges[shop] = (cur, cur + n - 1)
        cur += n
    return ranges


def _get_calendar_shop_layout(frames: dict[str, int]) -> list[tuple[str, int, int, int, int]]:
    """
    カレンダーシートの各店舗ブロックの行情報を計算する。

    カレンダー構造:
      Row 1: タイトル
      Row 2: 凡例
      Row 3: 日付数字ヘッダー
      Row 4: 曜日ヘッダー
      Row 5: 🏠 A店 ヘッダー行
      Row 6〜: A店データ行 ...

    戻り値: [(shop, shop_hdr_row, data_start_row, data_end_row, master_start_row), ...]
    """
    master_ranges = _get_master_row_ranges(frames)
    layout = []
    cal_row = 5          # 最初の店舗ヘッダー行
    for shop in SHOP_LIST:
        n = frames[shop]
        hdr_row    = cal_row
        data_start = cal_row + 1
        data_end   = data_start + n - 1
        mst_start  = master_ranges[shop][0]
        layout.append((shop, hdr_row, data_start, data_end, mst_start))
        cal_row = data_end + 2   # +1=空白行, +1=次の店舗ヘッダー
    return layout


# ==============================================================
# v5 マスタシート書き込み（固定枠付き・店舗ブロック形式）
# ==============================================================

def _write_master_sheet_v5(
    ws,
    shop_data:    dict[str, list[dict]],   # {shop: [row_dict, ...]}
    frames:       dict[str, int],           # {shop: n_slots}
    headers:      list[str],               # 列名リスト
    title:        str,
    header_color: str = "2C4A6E",
    shop_col_key: str = "店舗",            # 店舗名を入れるキー
):
    """
    v5 マスタシート共通書き込み関数（固定枠付き・店舗ブロック形式）。

    各店舗ごとに frames[shop] 行を確保し、
    デモデータで埋まらない残り行は 店舗名だけ入れて 氏名等は空欄にする。
    これにより、ユーザーが氏名を入力するだけでシステムに認識される。
    """
    TITLE_FILL = PatternFill("solid", fgColor="1B3A5C")
    HDR_FILL   = PatternFill("solid", fgColor=header_color)
    SHOP_FILLS = {
        "A店": PatternFill("solid", fgColor="D4E6F1"),
        "B店": PatternFill("solid", fgColor="D5F5E3"),
        "C店": PatternFill("solid", fgColor="FEF9E7"),
    }
    EMPTY_FILL = PatternFill("solid", fgColor="FAFBFC")
    ODD_FILL   = PatternFill("solid", fgColor="FFFFFF")

    def bdr():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    n_cols = len(headers)
    W = get_column_letter

    # タイトル行（Row 1）
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(row=1, column=1, value=title)
    c.font      = Font(bold=True, size=12, color="FFFFFF", name="メイリオ")
    c.fill      = TITLE_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # ヘッダー行（Row 2）
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=10, name="メイリオ")
        c.fill      = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = bdr()
    ws.row_dimensions[2].height = 22

    # データ行（Row 3〜）: 店舗ブロックごとに固定枠で書き込む
    master_ranges = _get_master_row_ranges(frames)
    for shop in SHOP_LIST:
        start_row, end_row = master_ranges[shop]
        data_rows = shop_data.get(shop, [])
        shop_fill = SHOP_FILLS.get(shop, EMPTY_FILL)

        for slot_idx in range(frames[shop]):
            row_idx = start_row + slot_idx
            # デモデータがあれば使い、なければ空欄枠
            if slot_idx < len(data_rows):
                row_data = data_rows[slot_idx]
                is_empty = False
            else:
                # 空欄枠: 店舗名だけ入れて他は空
                row_data = {h: "" for h in headers}
                row_data[shop_col_key] = shop
                is_empty = True

            fill = EMPTY_FILL if is_empty else (
                shop_fill if slot_idx % 2 == 0 else ODD_FILL
            )

            for col, h in enumerate(headers, 1):
                val = row_data.get(h, "")
                c   = ws.cell(row=row_idx, column=col, value=val)
                c.font      = Font(
                    size=10, name="メイリオ",
                    color="AAAAAA" if is_empty and h not in (shop_col_key, "ID") else "1C2330"
                )
                c.fill      = fill
                c.alignment = Alignment(
                    horizontal="left" if isinstance(val, str) and len(str(val)) > 6 else "center",
                    vertical="center"
                )
                c.border = bdr()
            ws.row_dimensions[row_idx].height = 20

    # 列幅調整（全データを走査）
    all_rows = [r for rows in shop_data.values() for r in rows]
    for col, h in enumerate(headers, 1):
        max_len = max(
            len(str(h)),
            max((len(str(r.get(h, ""))) for r in all_rows), default=0)
        )
        ws.column_dimensions[W(col)].width = min(max(max_len + 2, 8), 40)


# ==============================================================
# v5 スタッフカレンダーシート（Excel数式埋め込み版）
# ==============================================================

def _write_calendar_sheet_staff(wb, staff: list, year: int, month: int):
    """
    スタッフシフト表シート（カレンダー_スタッフ）を書き込む。

    【v5】 氏名セルと平日時間セルにExcel数式を埋め込む。
      氏名数式:   =IF(スタッフ!B{n}="","",スタッフ!B{n})
      平日時間式: =IF(スタッフ!B{n}="","",IF(スタッフ!F{n}="","08:00-19:00",スタッフ!F{n}&"-"&スタッフ!G{n}))
      土日:       空欄固定
    """
    import calendar as cal_mod
    _, days_in_month = cal_mod.monthrange(year, month)
    dates = [datetime.date(year, month, d) for d in range(1, days_in_month + 1)]

    ws = wb.create_sheet("カレンダー_スタッフ")

    C_TITLE    = PatternFill("solid", fgColor="1B3A5C")
    C_HDR_DATE = PatternFill("solid", fgColor="2C4A6E")
    C_SAT      = PatternFill("solid", fgColor="EBF5FB")
    C_SUN      = PatternFill("solid", fgColor="FDECEA")
    C_FORMULA  = PatternFill("solid", fgColor="F0FFF4")   # 数式セル: 薄緑
    C_EMPTY    = PatternFill("solid", fgColor="FAFBFC")   # 空欄枠: ごく薄いグレー
    C_WEEKEND  = PatternFill("solid", fgColor="F5F5F5")   # 土日空欄

    SHOP_DARK = {"A店": "1A5276", "B店": "145A32", "C店": "6E2F1A"}
    SHOP_LIGHT = {"A店": "D4E6F1", "B店": "D5F5E3", "C店": "FEF9E7"}

    def bdr(color="C8CDD2"):
        s = Side(style="thin", color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    NAME_COL  = 1   # A列: 氏名（数式）
    SHOP_COL  = 2   # B列: 店舗（数式）
    FIRST_DAY = 3   # C列〜: 日付

    last_col = FIRST_DAY + days_in_month - 1

    # ── Row 1: タイトル ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws.cell(row=1, column=1)
    c.value     = f"📅 スタッフ シフト表　{year}年{month}月"
    c.font      = Font(bold=True, size=14, color="FFFFFF", name="メイリオ")
    c.fill      = C_TITLE
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # ── Row 2: 凡例 ──
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    c = ws.cell(row=2, column=1)
    c.value = (
        "【運用方法】 氏名・時間はスタッフマスタから自動転記されます。"
        "イレギュラーな日のみ時間を直接上書き入力してください。"
        "空欄 = 休み・非番。"
    )
    c.font      = Font(italic=True, size=10, color="555555", name="メイリオ")
    c.fill      = PatternFill("solid", fgColor="EAF2FA")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20

    # ── Row 3〜4: 日付ヘッダー（数字 + 曜日）──
    DAY_NAMES = ["月", "火", "水", "木", "金", "土", "日"]
    ws.merge_cells(start_row=3, start_column=1, end_row=4, end_column=1)
    ws.merge_cells(start_row=3, start_column=2, end_row=4, end_column=2)
    for col in [1, 2]:
        label = "氏名" if col == 1 else "店舗"
        c = ws.cell(row=3, column=col, value=label)
        c.font      = Font(bold=True, color="FFFFFF", size=10, name="メイリオ")
        c.fill      = C_HDR_DATE
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = bdr()

    for di, dt in enumerate(dates):
        col    = FIRST_DAY + di
        is_sat = dt.weekday() == 5
        is_sun = dt.weekday() == 6
        hfill  = C_SUN if is_sun else (C_SAT if is_sat else C_HDR_DATE)
        hcolor = "C0392B" if is_sun else ("1A5276" if is_sat else "FFFFFF")

        c = ws.cell(row=3, column=col, value=int(dt.day))
        c.font = Font(bold=True, size=10, color=hcolor, name="メイリオ")
        c.fill = hfill; c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr()

        c = ws.cell(row=4, column=col, value=DAY_NAMES[dt.weekday()])
        c.font = Font(bold=True, size=9, color=hcolor, name="メイリオ")
        c.fill = hfill; c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr()
        ws.column_dimensions[get_column_letter(col)].width = 10

    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 16
    ws.column_dimensions[get_column_letter(NAME_COL)].width = 14
    ws.column_dimensions[get_column_letter(SHOP_COL)].width = 7

    # ── データ行（店舗ブロック）──
    layout = _get_calendar_shop_layout(STAFF_FRAMES)

    for shop, hdr_row, data_start, data_end, mst_start in layout:
        dark  = SHOP_DARK.get(shop, "2C3E50")
        light = SHOP_LIGHT.get(shop, "F0F0F0")
        n_slots = STAFF_FRAMES[shop]

        # 店舗ヘッダー行
        ws.merge_cells(start_row=hdr_row, start_column=1, end_row=hdr_row, end_column=last_col)
        c = ws.cell(row=hdr_row, column=1, value=f"  🏠 {shop}")
        c.font      = Font(bold=True, size=11, color="FFFFFF", name="メイリオ")
        c.fill      = PatternFill("solid", fgColor="2C3E50")
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = bdr("2C3E50")
        ws.row_dimensions[hdr_row].height = 22

        # 各スタッフ枠行
        for slot_idx in range(n_slots):
            cal_row = data_start + slot_idx
            mst_row = mst_start + slot_idx   # スタッフマスタの対応行

            # ── 氏名セル（数式）──
            # マスタのB列を参照し、空なら空白を返す
            name_formula = f'=IF(スタッフ!{STAFF_COL_NAME}{mst_row}="","",スタッフ!{STAFF_COL_NAME}{mst_row})'
            c = ws.cell(row=cal_row, column=NAME_COL, value=name_formula)
            c.font      = Font(bold=True, size=10, name="メイリオ")
            c.fill      = PatternFill("solid", fgColor=light)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border    = bdr()

            # ── 店舗セル（数式）──
            shop_formula = f'=IF(スタッフ!{STAFF_COL_NAME}{mst_row}="","",スタッフ!{STAFF_COL_SHOP}{mst_row})'
            c = ws.cell(row=cal_row, column=SHOP_COL, value=shop_formula)
            c.font      = Font(size=9, name="メイリオ")
            c.fill      = PatternFill("solid", fgColor=light)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = bdr()

            # ── 各日セル ──
            for di, dt in enumerate(dates):
                col    = FIRST_DAY + di
                is_sat = dt.weekday() == 5
                is_sun = dt.weekday() == 6

                if is_sat or is_sun:
                    # 土日は空欄固定（数式なし）
                    c = ws.cell(row=cal_row, column=col, value="")
                    c.fill   = C_WEEKEND
                    c.border = bdr()
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    # 平日: マスタのシフト時間を参照する数式
                    # F列(出勤時間)が空なら "08:00-19:00"、あれば F&"-"&G
                    time_formula = (
                        f'=IF(スタッフ!{STAFF_COL_NAME}{mst_row}="","",'
                        f'IF(スタッフ!{STAFF_COL_SHIFT_ST}{mst_row}="",'
                        f'"08:00-19:00",'
                        f'スタッフ!{STAFF_COL_SHIFT_ST}{mst_row}'
                        f'&"-"&'
                        f'スタッフ!{STAFF_COL_SHIFT_EN}{mst_row}))'
                    )
                    c = ws.cell(row=cal_row, column=col, value=time_formula)
                    c.fill   = C_FORMULA
                    c.border = bdr()
                    c.font   = Font(size=9, color="1E6B38", name="メイリオ")
                    c.alignment = Alignment(horizontal="center", vertical="center")

            ws.row_dimensions[cal_row].height = 22

        # 店舗間の空白行
        sep_row = data_end + 1
        for col in range(1, last_col + 1):
            ws.cell(row=sep_row, column=col).fill = PatternFill("solid", fgColor="F0F3F4")
        ws.row_dimensions[sep_row].height = 6

    # 印刷設定・ウィンドウ枠固定
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6)
    ws.freeze_panes = f"{get_column_letter(FIRST_DAY)}5"


# ==============================================================
# v5 利用者カレンダーシート（Excel数式埋め込み版）
# ==============================================================

def _write_calendar_sheet_users(wb, users: list, year: int, month: int):
    """
    利用者予定表シート（カレンダー_利用者）を書き込む。
    店舗ブロックで分割。

    【v5】 氏名セルと平日時間セルにExcel数式を埋め込む。
      氏名数式:   =IF(利用者!B{n}="","",利用者!B{n})
      平日時間式: =IF(利用者!B{n}="","","08:00-"&利用者!J{n})
      土日:       空欄固定
    """
    import calendar as cal_mod
    _, days_in_month = cal_mod.monthrange(year, month)
    dates = [datetime.date(year, month, d) for d in range(1, days_in_month + 1)]

    ws = wb.create_sheet("カレンダー_利用者")

    C_TITLE   = PatternFill("solid", fgColor="1B3A5C")
    C_HDR     = PatternFill("solid", fgColor="2C4A6E")
    C_SAT     = PatternFill("solid", fgColor="EBF5FB")
    C_SUN     = PatternFill("solid", fgColor="FDECEA")
    C_FORMULA = PatternFill("solid", fgColor="F0FFF4")
    C_WEEKEND = PatternFill("solid", fgColor="F5F5F5")
    C_WC      = PatternFill("solid", fgColor="FEF9E7")

    SHOP_DARK  = {"A店": "1A5276", "B店": "145A32", "C店": "6E2F1A"}
    SHOP_LIGHT = {"A店": "D4E6F1", "B店": "D5F5E3", "C店": "FDEBD0"}

    def bdr(color="C8CDD2"):
        s = Side(style="thin", color=color)
        return Border(left=s, right=s, top=s, bottom=s)

    NAME_COL  = 1
    SVC_COL   = 2
    FIRST_DAY = 3
    last_col  = FIRST_DAY + days_in_month - 1
    DAY_NAMES = ["月", "火", "水", "木", "金", "土", "日"]

    # ── Row 1: タイトル ──
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws.cell(row=1, column=1)
    c.value     = f"📋 利用者 月間利用予定表　{year}年{month}月"
    c.font      = Font(bold=True, size=14, color="FFFFFF", name="メイリオ")
    c.fill      = C_TITLE
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # ── Row 2: 凡例 ──
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    c = ws.cell(row=2, column=1)
    c.value = (
        "【運用方法】 氏名・デフォルト時間は利用者マスタから自動転記されます。"
        "変更がある日のみ直接上書きしてください。空欄 = 欠席・利用なし。♿ = 車椅子利用者。"
    )
    c.font  = Font(italic=True, size=10, color="555555", name="メイリオ")
    c.fill  = PatternFill("solid", fgColor="E8F8F5")
    c.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20

    # ── Row 3〜4: 日付ヘッダー ──
    ws.merge_cells(start_row=3, start_column=1, end_row=4, end_column=1)
    ws.merge_cells(start_row=3, start_column=2, end_row=4, end_column=2)
    for col, label in [(1, "利用者氏名"), (2, "サービス")]:
        c = ws.cell(row=3, column=col, value=label)
        c.font      = Font(bold=True, color="FFFFFF", size=10, name="メイリオ")
        c.fill      = C_HDR
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = bdr()

    for di, dt in enumerate(dates):
        col    = FIRST_DAY + di
        is_sat = dt.weekday() == 5
        is_sun = dt.weekday() == 6
        hfill  = C_SUN if is_sun else (C_SAT if is_sat else C_HDR)
        hcolor = "C0392B" if is_sun else ("1A5276" if is_sat else "FFFFFF")

        c = ws.cell(row=3, column=col, value=int(dt.day))
        c.font = Font(bold=True, size=10, color=hcolor, name="メイリオ")
        c.fill = hfill; c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr()

        c = ws.cell(row=4, column=col, value=DAY_NAMES[dt.weekday()])
        c.font = Font(bold=True, size=9, color=hcolor, name="メイリオ")
        c.fill = hfill; c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr()
        ws.column_dimensions[get_column_letter(col)].width = 10

    ws.row_dimensions[3].height = 20
    ws.row_dimensions[4].height = 16
    ws.column_dimensions[get_column_letter(NAME_COL)].width = 14
    ws.column_dimensions[get_column_letter(SVC_COL)].width  = 10

    # ── データ行（店舗ブロック）──
    layout = _get_calendar_shop_layout(USER_FRAMES)

    for shop, hdr_row, data_start, data_end, mst_start in layout:
        dark  = SHOP_DARK.get(shop, "2C3E50")
        light = SHOP_LIGHT.get(shop, "F0F0F0")
        n_slots = USER_FRAMES[shop]

        # 店舗ヘッダー行
        ws.merge_cells(start_row=hdr_row, start_column=1, end_row=hdr_row, end_column=last_col)
        c = ws.cell(row=hdr_row, column=1, value=f"  🏠 {shop}　（{n_slots}名枠）")
        c.font      = Font(bold=True, size=11, color="FFFFFF", name="メイリオ")
        c.fill      = PatternFill("solid", fgColor=dark)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = bdr(dark)
        ws.row_dimensions[hdr_row].height = 24

        for slot_idx in range(n_slots):
            cal_row = data_start + slot_idx
            mst_row = mst_start + slot_idx   # 利用者マスタの対応行

            # ── 氏名セル（数式）──
            # マスタB列（氏名）を参照。空なら空白。
            name_formula = (
                f'=IF(利用者!{USER_COL_NAME}{mst_row}="",'
                f'"♿ "&利用者!{USER_COL_NAME}{mst_row},'
                f'利用者!{USER_COL_NAME}{mst_row})'
            )
            # シンプル版（車椅子判定はマスタH列が必要で複雑になるため省略）
            name_formula = f'=IF(利用者!{USER_COL_NAME}{mst_row}="","",利用者!{USER_COL_NAME}{mst_row})'
            c = ws.cell(row=cal_row, column=NAME_COL, value=name_formula)
            c.font      = Font(bold=True, size=10, name="メイリオ")
            c.fill      = PatternFill("solid", fgColor=light)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border    = bdr()

            # ── サービス種別セル（数式）──
            # マスタF列（サービス種別）を参照
            svc_formula = f'=IF(利用者!{USER_COL_NAME}{mst_row}="","",利用者!F{mst_row})'
            c = ws.cell(row=cal_row, column=SVC_COL, value=svc_formula)
            c.font      = Font(size=9, name="メイリオ")
            c.fill      = PatternFill("solid", fgColor=light)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border    = bdr()

            # ── 各日セル ──
            for di, dt in enumerate(dates):
                col    = FIRST_DAY + di
                is_sat = dt.weekday() == 5
                is_sun = dt.weekday() == 6

                if is_sat or is_sun:
                    # 土日: 空欄固定
                    c = ws.cell(row=cal_row, column=col, value="")
                    c.fill   = C_WEEKEND
                    c.border = bdr()
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    # 平日: "08:00-{到着リミット}" の数式
                    # マスタJ列（到着リミット）を参照
                    # 例: 到着リミットが"09:00"なら → "08:00-09:00"
                    time_formula = (
                        f'=IF(利用者!{USER_COL_NAME}{mst_row}="","",'
                        f'"08:00-"&利用者!{USER_COL_LIMIT}{mst_row})'
                    )
                    c = ws.cell(row=cal_row, column=col, value=time_formula)
                    c.fill   = C_FORMULA
                    c.border = bdr()
                    c.font   = Font(size=9, color="1E6B38", name="メイリオ")
                    c.alignment = Alignment(horizontal="center", vertical="center")

            ws.row_dimensions[cal_row].height = 22

        # 店舗間の空白行
        sep_row = data_end + 1
        for col in range(1, last_col + 1):
            ws.cell(row=sep_row, column=col).fill = PatternFill("solid", fgColor="F0F3F4")
        ws.row_dimensions[sep_row].height = 8

    # 印刷設定
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6)
    ws.freeze_panes = f"{get_column_letter(FIRST_DAY)}5"


# ==============================================================
# v5 parse_excel_upload（空行スキップ・堅牢版）
# ==============================================================

def _is_empty_cell(val) -> bool:
    """
    セル値が「空」と見なせるかどうかを判定する。
    NaN・空文字・"nan"・Excel数式文字列（"="で始まる）はすべて空扱い。
    """
    if val is None:
        return True
    s = str(val).strip()
    return s in ("", "nan", "None") or s.startswith("=")


def parse_excel_upload(
    uploaded_file,
    default_pickup_limit:   int = 540,
    default_dropoff_target: int = 1050,
):
    """
    Excelを読み込んでデータクラスに変換。

    【v5 変更点】
    - 氏名が空・NaN・数式文字列の行を完全スキップ（固定枠の空行対応）
    - 行数制限なし（ユーザーが行を追加してもすべて読み込む）
    - Excelが未計算のままアップロードされた場合（数式文字列）も安全に処理
    """
    import pandas as _pd
    xl = _pd.ExcelFile(uploaded_file)

    svc_map = {
        "放課後等デイサービス": ServiceType.HOUKAGO_DEI,
        "A型":                ServiceType.A_TYPE,
        "B型":                ServiceType.B_TYPE,
    }

    # ---- 利用者シート ----
    df_u = xl.parse("利用者", header=1)
    users = []
    for i, row in df_u.iterrows():
        # 【v5】氏名が空・NaN・数式文字列の行はスキップ（固定枠の空行）
        name_val = row.get("氏名", "")
        if _is_empty_cell(name_val):
            continue

        incomp_raw = str(row.get("同乗不可ID", "")).strip()
        incomp = (
            [x.strip() for x in incomp_raw.split(",") if x.strip()]
            if incomp_raw not in ("", "nan") else []
        )
        wc     = bool(row.get("車椅子", False))
        pu_lim = hhmm_to_min(row.get("到着リミット", ""), default_pickup_limit)
        do_tgt = hhmm_to_min(row.get("送り目標",   ""), default_dropoff_target)

        # 緯度・経度: 空や数式の場合はデフォルト値
        try:
            lat = float(row.get("緯度", 36.695))
        except (ValueError, TypeError):
            lat = 36.695
        try:
            lng = float(row.get("経度", 137.211))
        except (ValueError, TypeError):
            lng = 137.211

        users.append(User(
            user_id        = str(row.get("ID", f"u{i+1}")),
            name           = str(name_val).strip(),
            address        = str(row.get("住所", "")),
            lat            = lat,
            lng            = lng,
            service_type   = svc_map.get(str(row.get("サービス種別", "")), ServiceType.HOUKAGO_DEI),
            shop           = str(row.get("店舗", "A店")),
            wheelchair     = wc,
            incompatible   = incomp,
            pickup_latest  = pu_lim,
            dropoff_target = do_tgt,
            service_time   = 600 if wc else 300,
        ))

    # ---- 車両シート ----
    df_v    = xl.parse("車両", header=1)
    type_cap = {"large": 7, "normal": 4, "kei": 3}
    vehicles = []
    for i, row in df_v.iterrows():
        # 【v5】車両名が空の行はスキップ
        vname = row.get("車両名", "")
        if _is_empty_cell(vname):
            continue
        vtype = str(row.get("種別コード", "normal"))
        try:
            depot_lat = float(row.get("デポ緯度", 36.695))
        except (ValueError, TypeError):
            depot_lat = 36.695
        try:
            depot_lng = float(row.get("デポ経度", 137.211))
        except (ValueError, TypeError):
            depot_lng = 137.211
        vehicles.append(Vehicle(
            vehicle_id    = str(row.get("ID", f"v{i+1}")),
            name          = str(vname).strip(),
            vehicle_type  = vtype,
            capacity      = int(row.get("定員", type_cap.get(vtype, 4))),
            shop          = str(row.get("店舗", "A店")),
            wheelchair_ok = bool(row.get("車椅子対応", False)),
            depot_lat     = depot_lat,
            depot_lng     = depot_lng,
        ))

    # ---- スタッフシート ----
    df_s  = xl.parse("スタッフ", header=1)
    staff = []
    for i, row in df_s.iterrows():
        # 【v5】氏名が空の行はスキップ（固定枠の空行）
        name_val = row.get("氏名", "")
        if _is_empty_cell(name_val):
            continue

        ss_raw = row.get("出勤時間", "")
        se_raw = row.get("退勤時間", "")
        # 数式文字列の場合は None 扱い
        ss = hhmm_to_min(ss_raw, -1) if not _is_empty_cell(ss_raw) else None
        se = hhmm_to_min(se_raw, -1) if not _is_empty_cell(se_raw) else None
        staff.append(Staff(
            staff_id    = str(row.get("ID", f"s{i+1}")),
            name        = str(name_val).strip(),
            shop        = str(row.get("店舗", "A店")),
            can_drive   = bool(row.get("運転可否", True)),
            priority    = int(row.get("優先度", 1)),
            shift_start = ss if ss != -1 else None,
            shift_end   = se if se != -1 else None,
        ))

    # ---- カレンダーシート ----
    calendar_data = _parse_calendar_sheets(xl, staff, users)

    return users, vehicles, staff, calendar_data


# ==============================================================
# v5 カレンダーパース（数式文字列対応・堅牢版）
# ==============================================================

def _parse_calendar_sheet(
    df:            pd.DataFrame,
    name_col:      int,
    first_day_col: int,
    year:          int,
    month:         int,
    default_start: int = 480,
    default_end:   int = 1140,
) -> dict:
    """
    横軸=日付、縦軸=名前 のカレンダーシートを解析する。

    【v5 変更点】
    - 名前セルが数式文字列（"="始まり）の場合スキップ
      （Excel未計算のままアップロードされた場合の安全対策）
    - 時間セルが数式文字列の場合は None（欠席扱い）として処理
    """
    import calendar as cal_mod
    _, days_in_month = cal_mod.monthrange(year, month)
    result: dict = {}

    for row_idx in range(len(df)):
        row = df.iloc[row_idx]

        name_val = row.iloc[name_col] if name_col < len(row) else None

        # 空・NaN・数式未計算はスキップ
        if name_val is None or str(name_val).strip() in ("", "nan"):
            continue
        name = str(name_val).strip()
        if name.startswith("="):
            # Excel数式が未計算のまま → スキップ
            continue

        # スキップ: タイトル・凡例・ヘッダー・店舗区切り行
        SKIP_KEYWORDS = [
            "🏠", "📅", "📋", "【入力", "【運用",
            "スタッフ シフト", "利用者 月間",
            "スタッフ名", "利用者氏名", "氏名",
        ]
        if any(kw in name for kw in SKIP_KEYWORDS):
            continue

        # ♿プレフィックスを除去
        name = name.replace("♿ ", "").replace("♿", "").strip()

        schedule: dict = {}
        for day in range(1, days_in_month + 1):
            day_col_idx = first_day_col + (day - 1)
            if day_col_idx >= len(row):
                break

            cell_val = row.iloc[day_col_idx]
            date_str = datetime.date(year, month, day).strftime("%Y-%m-%d")

            # 【v5】数式文字列（未計算）は None（欠席扱い）
            if cell_val is not None and str(cell_val).strip().startswith("="):
                schedule[date_str] = None
            else:
                schedule[date_str] = parse_time_range(cell_val, default_start, default_end)

        result[name] = schedule

    return result


def _parse_calendar_sheets(xl, staff, users) -> Optional[dict]:
    """
    カレンダー_スタッフ / カレンダー_利用者 シートを解析。
    シートがなければ None を返す。
    """
    has_staff = "カレンダー_スタッフ" in xl.sheet_names
    has_users = "カレンダー_利用者"   in xl.sheet_names

    if not has_staff and not has_users:
        return None

    import datetime as _dt
    today = _dt.date.today()
    year  = today.year
    month = today.month

    staff_cal: dict = {}
    users_cal: dict = {}

    if has_staff:
        raw_s = xl.parse("カレンダー_スタッフ", header=None)
        year, month = _detect_year_month_from_sheet(raw_s, xl, "カレンダー_スタッフ")
        staff_cal   = _parse_calendar_sheet(
            raw_s, name_col=0, first_day_col=2,
            year=year, month=month,
            default_start=480, default_end=1140,
        )

    if has_users:
        raw_u = xl.parse("カレンダー_利用者", header=None)
        year, month = _detect_year_month_from_sheet(raw_u, xl, "カレンダー_利用者")
        users_cal   = _parse_calendar_sheet(
            raw_u, name_col=0, first_day_col=2,
            year=year, month=month,
            default_start=480, default_end=1050,
        )

    return {"staff": staff_cal, "users": users_cal, "year": year, "month": month}


def _detect_year_month_from_sheet(df_raw, xl, sheet_name) -> tuple:
    """タイトル行から YYYY年M月 を抽出して返す。失敗したら今月を返す。"""
    import re, datetime as _dt
    today = _dt.date.today()
    try:
        title_val = str(df_raw.iloc[0, 0])
        m = re.search(r"(\d{4})年\s*(\d{1,2})月", title_val)
        if m:
            return int(m.group(1)), int(m.group(2))
    except Exception:
        pass
    return today.year, today.month


# ==============================================================
# v5 get_sample_excel（固定枠 + 数式埋め込み版）
# ==============================================================

def get_sample_excel(year: Optional[int] = None, month: Optional[int] = None) -> bytes:
    """
    v5対応サンプルExcelを生成。

    【v5 変更点】
    1. マスタに固定枠を設ける（スタッフ A10/B6/C6、利用者 各30）
    2. カレンダーにExcel数式を埋め込み（マスタへの参照）
    3. 空欄枠には店舗名だけ入力して氏名は空に

    シート構成:
      利用者            - マスタ（A30/B30/C30 固定枠）
      車両              - マスタ
      スタッフ          - マスタ（A10/B6/C6 固定枠）
      カレンダー_スタッフ - 数式で氏名・時間をマスタから自動転記
      カレンダー_利用者   - 数式で氏名・時間をマスタから自動転記
      記入例            - 入力ガイド
    """
    users, vehicles, staff = get_demo_data()

    if year is None or month is None:
        today = datetime.date.today()
        year  = today.year
        month = today.month

    if not OPENPYXL_AVAILABLE:
        buf = io.StringIO()
        pd.DataFrame([{"氏名": u.name, "店舗": u.shop} for u in users]).to_csv(buf, index=False)
        return buf.getvalue().encode("utf-8-sig")

    wb = Workbook()
    wb.remove(wb.active)

    # ========== 利用者マスタ（固定枠 各30行）==========
    ws_u = wb.create_sheet("利用者")
    user_headers = [
        "ID", "氏名", "住所", "緯度", "経度",
        "サービス種別", "店舗", "車椅子", "同乗不可ID",
        "到着リミット", "送り目標",
    ]
    # 店舗ごとにデモデータを整理
    user_by_shop = {shop: [] for shop in SHOP_LIST}
    for u in users:
        if u.shop in user_by_shop:
            user_by_shop[u.shop].append({
                "ID": u.user_id, "氏名": u.name, "住所": u.address,
                "緯度": u.lat, "経度": u.lng,
                "サービス種別": u.service_type.value, "店舗": u.shop,
                "車椅子": u.wheelchair, "同乗不可ID": ",".join(u.incompatible),
                "到着リミット": min_to_hhmm(u.pickup_latest),
                "送り目標":     min_to_hhmm(u.dropoff_target),
            })
    _write_master_sheet_v5(
        ws_u, user_by_shop, USER_FRAMES, user_headers,
        title="利用者マスタ　（各店30名枠・空欄行に追記してください）",
        header_color="2C4A6E", shop_col_key="店舗",
    )

    # ========== 車両マスタ（固定枠なし・そのまま）==========
    ws_v = wb.create_sheet("車両")
    _write_master_sheet(ws_v, [
        {"ID": v.vehicle_id, "車両名": v.name, "種別コード": v.vehicle_type,
         "定員": v.capacity, "店舗": v.shop, "車椅子対応": v.wheelchair_ok,
         "デポ緯度": v.depot_lat, "デポ経度": v.depot_lng,
         } for v in vehicles
    ], title="車両マスタ", header_color="1A5276")

    # ========== スタッフマスタ（固定枠 A10/B6/C6）==========
    ws_s = wb.create_sheet("スタッフ")
    staff_headers = [
        "ID", "氏名", "店舗", "運転可否", "優先度", "出勤時間", "退勤時間",
    ]
    staff_by_shop = {shop: [] for shop in SHOP_LIST}
    for s in staff:
        if s.shop in staff_by_shop:
            staff_by_shop[s.shop].append({
                "ID": s.staff_id, "氏名": s.name, "店舗": s.shop,
                "運転可否": s.can_drive, "優先度": s.priority,
                "出勤時間": min_to_hhmm(s.shift_start) if s.shift_start else "",
                "退勤時間": min_to_hhmm(s.shift_end)   if s.shift_end   else "",
            })
    _write_master_sheet_v5(
        ws_s, staff_by_shop, STAFF_FRAMES, staff_headers,
        title="スタッフマスタ　（A店10枠 / B店6枠 / C店6枠）",
        header_color="145A32", shop_col_key="店舗",
    )

    # ========== カレンダーシート（数式埋め込み）==========
    _write_calendar_sheet_staff(wb, staff, year, month)
    _write_calendar_sheet_users(wb, users, year, month)

    # ========== 記入例シート ==========
    ws_ex = wb.create_sheet("記入例")
    _write_master_sheet(ws_ex, [
        {"項目": "カレンダー 基本入力形式",     "例": "09:00-15:30",   "説明": "開始-終了の時刻を直接入力（HH:MM-HH:MM）"},
        {"項目": "カレンダー 欠席・休み",        "例": "（空欄）",       "説明": "空欄にするとその日は計算から除外"},
        {"項目": "カレンダー 数式の上書き方法",  "例": "09:00-17:00",   "説明": "数式を消して直接時間を入力してください"},
        {"項目": "カレンダー 〇フォールバック",  "例": "〇",            "説明": "〇を入力するとマスタのデフォルト時間を適用"},
        {"項目": "到着リミット / 送り目標",      "例": "09:00",          "説明": "HH:MM形式で記入（カレンダーの時間基準になります）"},
        {"項目": "出勤時間 / 退勤時間",          "例": "08:00 / 19:00",  "説明": "スタッフのシフト時間（空欄=08:00-19:00デフォルト）"},
        {"項目": "マスタ 空欄枠への追加",        "例": "氏名を入力",     "説明": "空欄行の氏名を入力するとカレンダーに自動反映されます"},
        {"項目": "マスタ 行の追加",              "例": "行を挿入",       "説明": "枠以上に追加する場合は行を挿入 → カレンダーにも行を追加"},
        {"項目": "車椅子フラグ",                 "例": "TRUE / FALSE",   "説明": "TRUEで乗降時間が10分に設定"},
        {"項目": "同乗不可ID",                  "例": "u1,u3",          "説明": "カンマ区切りで複数指定可"},
    ], title="記入例・入力ガイド（v5）", header_color="7D3C98")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==============================================================
# _write_master_sheet（v4から継続・車両マスタ等で使用）
# ==============================================================

def _write_master_sheet(ws, rows: list[dict], title: str, header_color: str = "2C4A6E"):
    """
    マスタシート共通書き込み（固定枠なし版・車両マスタ等で使用）。
    """
    if not rows:
        return

    TITLE_FILL = PatternFill("solid", fgColor="1B3A5C")
    HDR_FILL   = PatternFill("solid", fgColor=header_color)
    EVEN_FILL  = PatternFill("solid", fgColor="F8F9FA")
    ODD_FILL   = PatternFill("solid", fgColor="FFFFFF")

    def bdr():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    headers = list(rows[0].keys())
    n_cols  = len(headers)

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(row=1, column=1, value=title)
    c.font      = Font(bold=True, size=12, color="FFFFFF", name="メイリオ")
    c.fill      = TITLE_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    for col, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=10, name="メイリオ")
        c.fill      = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = bdr()
        max_len = max(len(str(h)), max((len(str(r.get(h, ""))) for r in rows), default=0))
        ws.column_dimensions[get_column_letter(col)].width = min(max(max_len + 2, 8), 40)
    ws.row_dimensions[2].height = 22

    for ri, row_data in enumerate(rows):
        row_idx = ri + 3
        fill    = EVEN_FILL if ri % 2 == 0 else ODD_FILL
        for col, h in enumerate(headers, 1):
            val = row_data.get(h, "")
            c   = ws.cell(row=row_idx, column=col, value=val)
            c.font      = Font(size=10, name="メイリオ")
            c.fill      = fill
            c.alignment = Alignment(
                horizontal="left" if isinstance(val, str) and len(str(val)) > 6 else "center",
                vertical="center"
            )
            c.border = bdr()
        ws.row_dimensions[row_idx].height = 20


def extract_for_date(
    calendar_data: Optional[dict],
    all_users:     list[User],
    all_staff:     list[Staff],
    target_date:   datetime.date,
    default_pickup_limit:   int = 540,
    default_dropoff_target: int = 1050,
    default_shift_start:    int = 480,
    default_shift_end:      int = 1140,
) -> tuple[list[User], list[Staff]]:
    """
    v4カレンダーから指定日の出席者・出勤者を抽出。
    カレンダーに時間が入力されている場合は User/Staff の時間枠を上書きする。

    戻り値:
      (filtered_users, filtered_staff)
      ※User.pickup_latest / Staff.shift_start/shift_end がカレンダーの値で更新される
    """
    if calendar_data is None:
        return all_users, all_staff

    date_str   = target_date.strftime("%Y-%m-%d")
    staff_cal  = calendar_data.get("staff",  {})
    users_cal  = calendar_data.get("users",  {})

    # ---- スタッフフィルタリング ----
    filtered_staff = []
    for s in all_staff:
        schedule = staff_cal.get(s.name, {})
        tw = schedule.get(date_str, "MISSING")

        if tw == "MISSING":
            # カレンダーにない → 全員参加扱い
            filtered_staff.append(s)
            continue

        if tw is None:
            # 明示的に空欄 → 休み
            continue

        # 時間情報あり → shift_start/shift_end を上書きしたコピーを作る
        import copy
        s_copy             = copy.copy(s)
        s_copy.shift_start = tw[0]
        s_copy.shift_end   = tw[1]
        filtered_staff.append(s_copy)

    # ---- 利用者フィルタリング ----
    filtered_users = []
    for u in all_users:
        schedule = users_cal.get(u.name, {})
        tw = schedule.get(date_str, "MISSING")

        if tw == "MISSING":
            filtered_users.append(u)
            continue

        if tw is None:
            continue

        import copy
        u_copy               = copy.copy(u)
        # カレンダーの開始時刻 → デポ出発時刻の上限として pickup_latest に使用
        # 終了時刻 → 施設到着リミット（迎え便）/ 送り目標（送り便）
        u_copy.pickup_latest  = tw[1]   # カレンダー終了時刻 = 迎え到着リミット
        u_copy.dropoff_target = tw[1] + 150  # 送り目標は+2.5h
        filtered_users.append(u_copy)

    return filtered_users, filtered_staff


# ==============================================================
# Excel 送迎結果出力（v3から継続）
# ==============================================================

def build_excel_output(
    pickup_routes:  list[AssignedRoute],
    dropoff_routes: list[AssignedRoute],
    target_date:    Optional[datetime.date] = None,
) -> bytes:
    """迎え便・送り便を別シートに出力する1ファイルのExcel"""
    if not OPENPYXL_AVAILABLE:
        rows = (
            _routes_to_rows(pickup_routes,  "迎え") +
            _routes_to_rows(dropoff_routes, "送り")
        )
        buf = io.StringIO()
        pd.DataFrame(rows).to_csv(buf, index=False, encoding="utf-8-sig")
        return buf.getvalue().encode("utf-8-sig")

    wb       = Workbook()
    wb.remove(wb.active)
    date_str = target_date.strftime("%Y/%m/%d") if target_date else ""

    for routes, sheet_name in [
        (pickup_routes,  "迎え便"),
        (dropoff_routes, "送り便"),
    ]:
        ws = wb.create_sheet(title=sheet_name)
        _write_route_sheet(ws, routes, sheet_name, date_str)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _write_route_sheet(ws, routes, sheet_title, date_str=""):
    """送迎表シートを書き込む（v3から継続・店舗ブロック）"""
    TITLE_FILL = PatternFill("solid", fgColor="1B3A5C")
    DATE_FILL  = PatternFill("solid", fgColor="2D5F8A")
    HDR_FILL   = PatternFill("solid", fgColor="2C4A6E")
    WC_FILL    = PatternFill("solid", fgColor="FADBD8")
    ALT_FILL   = PatternFill("solid", fgColor="F6F8FA")
    SHOP_FILLS = [
        (PatternFill("solid", fgColor="1A5276"), PatternFill("solid", fgColor="D4E6F1")),
        (PatternFill("solid", fgColor="145A32"), PatternFill("solid", fgColor="D5F5E3")),
        (PatternFill("solid", fgColor="6E2F1A"), PatternFill("solid", fgColor="FDEBD0")),
        (PatternFill("solid", fgColor="4A235A"), PatternFill("solid", fgColor="F3E6FA")),
    ]

    def bdr():
        s = Side(style="thin", color="BBBBBB")
        return Border(left=s, right=s, top=s, bottom=s)

    N_COLS = 10
    W = get_column_letter

    ws.merge_cells(f"A1:{W(N_COLS)}1")
    c = ws["A1"]
    c.value     = f"🚌  送迎ルート表　【{sheet_title}】"
    c.font      = Font(bold=True, size=15, color="FFFFFF", name="メイリオ")
    c.fill      = TITLE_FILL
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    hdr_row = 2
    if date_str:
        ws.merge_cells(f"A2:{W(N_COLS)}2")
        c = ws["A2"]
        c.value     = f"実施日：{date_str}"
        c.font      = Font(bold=True, size=12, color="FFFFFF", name="メイリオ")
        c.fill      = DATE_FILL
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = bdr()
        ws.row_dimensions[2].height = 24
        hdr_row = 3

    headers = ["店舗","車両名","運転担当","優先度","順番","利用者氏名","サービス","住所","到着予定","備考"]
    widths  = [10,    20,      14,        6,       6,     16,          14,       34,    10,       14]
    for col, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=hdr_row, column=col, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=10, name="メイリオ")
        c.fill      = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = bdr()
        ws.column_dimensions[W(col)].width = w
    ws.row_dimensions[hdr_row].height = 24

    shops = sorted(set(r.shop for r in routes))
    row   = hdr_row + 1

    for si, shop in enumerate(shops):
        dark_fill, light_fill = SHOP_FILLS[si % len(SHOP_FILLS)]
        ws.merge_cells(f"A{row}:{W(N_COLS)}{row}")
        c = ws.cell(row=row, column=1, value=f"　🏠 {shop}")
        c.font      = Font(bold=True, size=12, color="FFFFFF", name="メイリオ")
        c.fill      = dark_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border    = bdr()
        ws.row_dimensions[row].height = 28
        row += 1

        for route in [r for r in routes if r.shop == shop]:
            dn = route.driver.name     if route.driver else "未定"
            dp = route.driver.priority if route.driver else "-"
            for i, stop in enumerate(route.stops):
                h, m = divmod(stop["arrival_min"], 60)
                fill = WC_FILL if stop["user"].wheelchair else (
                    light_fill if i == 0 else ALT_FILL
                )
                data = [
                    route.shop         if i == 0 else "",
                    route.vehicle.name if i == 0 else "",
                    dn                 if i == 0 else "",
                    dp                 if i == 0 else "",
                    i + 1,
                    stop["user"].name,
                    stop["user"].service_type.value,
                    stop["address"],
                    f"{h:02d}:{m:02d}",
                    "♿ 車椅子" if stop["user"].wheelchair else "",
                ]
                for col, val in enumerate(data, 1):
                    c = ws.cell(row=row, column=col, value=val)
                    c.font      = Font(bold=(i == 0), size=10, name="メイリオ")
                    c.alignment = Alignment(
                        horizontal="center" if col in [1,2,3,4,5,9,10] else "left",
                        vertical="center"
                    )
                    c.border = bdr()
                    if fill: c.fill = fill
                ws.row_dimensions[row].height = 20
                row += 1
        row += 1

    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75)
    ws.print_title_rows = f"1:{hdr_row}"


def _routes_to_rows(routes, trip_label):
    rows = []
    for r in routes:
        for i, stop in enumerate(r.stops):
            h, m = divmod(stop["arrival_min"], 60)
            rows.append({
                "便": trip_label, "店舗": r.shop, "車両名": r.vehicle.name,
                "運転担当": r.driver.name if r.driver else "未定",
                "優先度": r.driver.priority if r.driver else "-",
                "順番": i+1, "氏名": stop["user"].name,
                "サービス": stop["user"].service_type.value,
                "住所": stop["address"], "到着予定": f"{h:02d}:{m:02d}",
                "車椅子": "♿" if stop["user"].wheelchair else "",
            })
    return rows


def routes_to_dataframe(routes: list[AssignedRoute]) -> pd.DataFrame:
    rows = []
    for route in routes:
        dn = route.driver.name     if route.driver else "未定"
        dp = route.driver.priority if route.driver else "-"
        for i, stop in enumerate(route.stops):
            h, m = divmod(stop["arrival_min"], 60)
            rows.append({
                "店舗":     route.shop,
                "車両名":   route.vehicle.name,
                "運転担当": dn,
                "優先度":   dp,
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

SHOP_MAP_COLORS = ["blue", "green", "red", "purple", "orange", "darkblue"]


def render_map(routes: list[AssignedRoute]):
    if not FOLIUM_AVAILABLE:
        st.info("📦 `folium` と `streamlit-folium` をインストールすると地図が表示されます")
        return
    if not routes:
        return

    all_lats = [s["lat"] for r in routes for s in r.stops]
    all_lngs = [s["lng"] for r in routes for s in r.stops]
    center   = [sum(all_lats)/len(all_lats), sum(all_lngs)/len(all_lngs)]

    m = folium.Map(location=center, zoom_start=12, tiles="CartoDB positron")
    shops = sorted(set(r.shop for r in routes))
    sc    = {s: SHOP_MAP_COLORS[i % len(SHOP_MAP_COLORS)] for i, s in enumerate(shops)}

    for route in routes:
        color  = sc[route.shop]
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
            h, mn  = divmod(stop["arrival_min"], 60)
            wc     = "♿ " if stop["user"].wheelchair else ""
            tw_lim = min_to_hhmm(stop["user"].pickup_latest) if route.trip_type == TripType.PICKUP \
                     else min_to_hhmm(stop["user"].dropoff_target)
            popup  = (
                f"<b>{i+1}. {stop['user'].name}</b><br>"
                f"{wc}{stop['address']}<br>"
                f"🕐 到着: {h:02d}:{mn:02d} / リミット: {tw_lim}<br>"
                f"🚗 {route.vehicle.name} ({route.shop})"
            )
            folium.CircleMarker(
                location=[stop["lat"], stop["lng"]],
                radius=8, color=color, fill=True, fill_opacity=0.85,
                tooltip=f"{i+1}. {stop['user'].name}",
                popup=folium.Popup(popup, max_width=240),
            ).add_to(m)

    st_folium(m, width=None, height=480, returned_objects=[])


# ==============================================================
# Streamlit UI メイン
# ==============================================================

def main():

    # ---- ヒーローヘッダー ----
    st.markdown("""
    <div class="hero fade-up">
      <div class="v-badge">VERSION 5</div>
      <h1>送迎ルート最適化システム</h1>
      <p>
        放課後等デイサービス・就労継続支援A型/B型 対応　｜　3店舗混載禁止　｜　月間カレンダー連動<br>
        個別Time Windows・スタッフシフト制約・迎え／送り同時計算
      </p>
    </div>
    """, unsafe_allow_html=True)

    # ================================================================
    # サイドバー
    # ================================================================
    with st.sidebar:

        # アルゴリズム表示
        algo_color = "#27AE60" if ORTOOLS_AVAILABLE else "#E67E22"
        algo_label = "OR-Tools（高精度VRP）" if ORTOOLS_AVAILABLE else "グリーディ（高速）"
        st.markdown(f"""
        <div class="sidebar-section">
          <h4>⚙️ システム状態</h4>
          <div style="font-size:12px;color:{algo_color};font-weight:700;">● {algo_label}</div>
          <div style="font-size:11px;color:#888;margin-top:4px;">
            {"OR-Tools インストール済み" if ORTOOLS_AVAILABLE else "pip install ortools で高精度モードに変更可能"}
          </div>
        </div>
        """, unsafe_allow_html=True)

        # 対象日選択
        st.markdown('<div class="sidebar-section"><h4>📅 対象日</h4>', unsafe_allow_html=True)
        target_date = st.date_input(
            "送迎実施日",
            value=datetime.date.today(),
            label_visibility="collapsed",
        )
        st.markdown(f"""
        <div class="cal-date-badge">{target_date.strftime("%Y年 %m月 %d日 (%a)")}</div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)

        # 時刻設定
        st.markdown('<div class="sidebar-section"><h4>⏰ デフォルト時刻設定</h4>', unsafe_allow_html=True)
        st.caption("Excelが空欄の場合のデフォルト値")

        st.markdown("**迎え便**")
        c1, c2 = st.columns(2)
        with c1:
            pu_sh = st.number_input("出発 時", 5,  12, 8,  key="pu_sh")
            pu_sm = st.number_input("出発 分", 0,  55, 0,  step=5, key="pu_sm")
        with c2:
            pu_lh = st.number_input("リミット 時", 6,  13, 9,  key="pu_lh")
            pu_lm = st.number_input("リミット 分", 0,  55, 0,  step=5, key="pu_lm")

        st.markdown("**送り便**")
        c3, c4 = st.columns(2)
        with c3:
            do_sh = st.number_input("出発 時", 15, 20, 17, key="do_sh")
            do_sm = st.number_input("出発 分", 0,  55, 0,  step=5, key="do_sm")
        with c4:
            do_lh = st.number_input("目標 時", 16, 22, 19, key="do_lh")
            do_lm = st.number_input("目標 分", 0,  55, 0,  step=5, key="do_lm")

        pu_start = pu_sh * 60 + pu_sm
        pu_limit = pu_lh * 60 + pu_lm
        do_start = do_sh * 60 + do_sm
        do_limit = do_lh * 60 + do_lm
        st.markdown("</div>", unsafe_allow_html=True)

        # サンプルExcelダウンロード
        st.markdown('<div class="sidebar-section"><h4>📥 テンプレート</h4>', unsafe_allow_html=True)
        if OPENPYXL_AVAILABLE:
            st.download_button(
                "サンプルExcel（v3対応）",
                data=get_sample_excel(),
                file_name="送迎マスタ_サンプルv3.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.caption("シート: 利用者 / 車両 / スタッフ / 月間カレンダー / 記入例")
        st.markdown("</div>", unsafe_allow_html=True)

    # ================================================================
    # STEP 1: データ読み込み
    # ================================================================
    step_header(1, "データの読み込み", "Excelアップロード または デモデータで即時起動")

    col_up, col_demo = st.columns([3, 1])

    with col_up:
        uploaded = st.file_uploader(
            "Excelファイルをアップロード（.xlsx / .xls）",
            type=["xlsx", "xls"],
            help="サイドバーから v3 対応サンプルExcelをDLしてご利用ください",
        )
        if uploaded:
            try:
                with st.spinner("読み込み中..."):
                    users, vehicles, staff, cal_df = parse_excel_upload(
                        uploaded, pu_limit, do_limit
                    )
                st.session_state.update({
                    "users": users, "vehicles": vehicles,
                    "staff": staff, "calendar": cal_df,
                })
                st.success(
                    f"✅ 読み込み完了　"
                    f"利用者 **{len(users)}名** / 車両 **{len(vehicles)}台** / "
                    f"スタッフ **{len(staff)}名**"
                    + (f" / 月間カレンダー **あり**" if cal_df is not None else "")
                )
            except Exception as e:
                st.error(f"❌ 読み込み失敗: {e}")
                st.info("シート名・カラム名をサンプルExcelと一致させてください")

    with col_demo:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if st.button("🎯 デモデータで試す", use_container_width=True, type="secondary"):
            users, vehicles, staff = get_demo_data()
            cal_df = build_demo_calendar(users, staff, target_date.year, target_date.month)
            st.session_state.update({
                "users": users, "vehicles": vehicles,
                "staff": staff, "calendar": cal_df,
            })
            st.success("✅ デモデータ読み込み完了")

    if "users" not in st.session_state:
        st.info("👆 Excelをアップロードするか、デモデータで試してください")
        return

    all_users    = st.session_state["users"]
    all_vehicles = st.session_state["vehicles"]
    all_staff    = st.session_state["staff"]
    calendar_df  = st.session_state.get("calendar")

    # プレビュー（折りたたみ）
    with st.expander("📋 マスタデータのプレビュー", expanded=False):
        t1, t2, t3, t4 = st.tabs(["👶 利用者", "🚗 車両", "👤 スタッフ", "📅 カレンダー"])
        with t1:
            st.dataframe(pd.DataFrame([{
                "ID": u.user_id, "氏名": u.name, "店舗": u.shop,
                "サービス": u.service_type.value, "住所": u.address,
                "車椅子": "♿" if u.wheelchair else "",
                "乗降(分)": u.service_time // 60,
                "迎えリミット": min_to_hhmm(u.pickup_latest),
                "送り目標":    min_to_hhmm(u.dropoff_target),
                "同乗不可": ",".join(u.incompatible) or "なし",
            } for u in all_users]), use_container_width=True, hide_index=True)
        with t2:
            st.dataframe(pd.DataFrame([{
                "車両名": v.name, "店舗": v.shop,
                "種別": v.vehicle_type, "定員": v.capacity,
                "車椅子対応": "✅" if v.wheelchair_ok else "✗",
            } for v in all_vehicles]), use_container_width=True, hide_index=True)
        with t3:
            st.dataframe(pd.DataFrame([{
                "氏名": s.name, "店舗": s.shop,
                "運転": "✅" if s.can_drive else "❌",
                "優先度": s.priority,
                "出勤": min_to_hhmm(s.shift_start) if s.shift_start else "終日",
                "退勤": min_to_hhmm(s.shift_end)   if s.shift_end   else "終日",
            } for s in all_staff]), use_container_width=True, hide_index=True)
        with t4:
            if calendar_df is not None:
                # calendar_df は v4以降 dict 型（{"staff":..., "users":..., "year":..., "month":...}）
                year_c  = calendar_df.get("year",  "?")
                month_c = calendar_df.get("month", "?")
                sc      = calendar_df.get("staff", {})
                uc      = calendar_df.get("users", {})
                st.caption(
                    f"📅 {year_c}年{month_c}月　"
                    f"スタッフ: **{len(sc)}名**分　利用者: **{len(uc)}名**分"
                )
                col_s, col_u = st.columns(2)
                with col_s:
                    st.markdown("**スタッフ カレンダー**")
                    if sc:
                        sample_name = next(iter(sc))
                        sample_days = list(sc[sample_name].items())[:7]
                        rows = [
                            {"日付": d,
                             "時間": f"{min_to_hhmm(v[0])}〜{min_to_hhmm(v[1])}" if v else "休み"}
                            for d, v in sample_days
                        ]
                        st.caption(f"（{sample_name} の最初の7日間）")
                        st.dataframe(
                            pd.DataFrame(rows),
                            use_container_width=True, hide_index=True
                        )
                    else:
                        st.info("スタッフカレンダーデータなし")
                with col_u:
                    st.markdown("**利用者 カレンダー**")
                    if uc:
                        sample_name = next(iter(uc))
                        sample_days = list(uc[sample_name].items())[:7]
                        rows = [
                            {"日付": d,
                             "時間": f"{min_to_hhmm(v[0])}〜{min_to_hhmm(v[1])}" if v else "欠席"}
                            for d, v in sample_days
                        ]
                        st.caption(f"（{sample_name} の最初の7日間）")
                        st.dataframe(
                            pd.DataFrame(rows),
                            use_container_width=True, hide_index=True
                        )
                    else:
                        st.info("利用者カレンダーデータなし")
            else:
                st.info("月間カレンダーシートが読み込まれていません")

    st.divider()

    # ================================================================
    # STEP 2: 対象日フィルタリング + 欠席トグル
    # ================================================================
    step_header(
        2,
        f"対象日のデータ確認  {target_date.strftime('%m/%d')}",
        "月間カレンダーから自動抽出 → 手動微調整"
    )

    # 月間カレンダーから対象日のデータを抽出
    if calendar_df is not None:
        filtered_users, filtered_staff = extract_for_date(
            calendar_df, all_users, all_staff, target_date
        )
        n_extracted_u = len(filtered_users)
        n_extracted_s = len(filtered_staff)

        st.markdown(
            f"📅 **{target_date.strftime('%Y年%m月%d日')}** の自動抽出結果　"
            f"利用者 **{n_extracted_u}名** / スタッフ **{n_extracted_s}名**",
        )
    else:
        filtered_users = all_users
        filtered_staff = all_staff
        st.info("月間カレンダーが未読み込みのため全員を表示しています。")

    # 当日欠席トグル（v2継続）
    st.markdown("#### 📝 当日の出欠・シフト最終確認")
    st.caption(
        "チェック = 本日参加／出勤。急な欠席・変更はここで外してください。"
    )

    attend_df = pd.DataFrame([{
        "出席":      True,
        "店舗":      u.shop,
        "氏名":      u.name,
        "サービス":  u.service_type.value,
        "住所":      u.address,
        "車椅子":    "♿" if u.wheelchair else "",
        "迎えリミット": min_to_hhmm(u.pickup_latest),
        "送り目標":  min_to_hhmm(u.dropoff_target),
        "_uid":      u.user_id,
    } for u in sorted(filtered_users, key=lambda x: (x.shop, x.name))])

    edited_attend = st.data_editor(
        attend_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "出席":         st.column_config.CheckboxColumn("出席", width="small"),
            "店舗":         st.column_config.TextColumn("店舗",    width="small"),
            "氏名":         st.column_config.TextColumn("氏名",    width="medium"),
            "サービス":     st.column_config.TextColumn("サービス",width="medium"),
            "住所":         st.column_config.TextColumn("住所",    width="large"),
            "車椅子":       st.column_config.TextColumn("車椅子",  width="small"),
            "迎えリミット": st.column_config.TextColumn("迎えリミット", width="small"),
            "送り目標":     st.column_config.TextColumn("送り目標",    width="small"),
            "_uid":         st.column_config.TextColumn("_uid",    disabled=True),
        },
        disabled=["店舗","氏名","サービス","住所","車椅子","迎えリミット","送り目標","_uid"],
        column_order=["出席","店舗","氏名","サービス","住所","車椅子","迎えリミット","送り目標"],
        key="attendance_editor",
    )

    attending_ids   = set(edited_attend.loc[edited_attend["出席"] == True, "_uid"].tolist())
    attending_users = [u for u in filtered_users if u.user_id in attending_ids]

    # スタッフ確認
    st.markdown("#### 👤 出勤スタッフ最終確認")
    staff_df = pd.DataFrame([{
        "出勤":   True,
        "店舗":   s.shop,
        "氏名":   s.name,
        "優先度": s.priority,
        "運転":   "✅" if s.can_drive else "❌",
        "出勤時間": min_to_hhmm(s.shift_start) if s.shift_start else "終日",
        "退勤時間": min_to_hhmm(s.shift_end)   if s.shift_end   else "終日",
        "_sid":   s.staff_id,
    } for s in sorted(filtered_staff, key=lambda x: (x.shop, x.priority))])

    edited_staff = st.data_editor(
        staff_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "出勤":     st.column_config.CheckboxColumn("出勤", width="small"),
            "店舗":     st.column_config.TextColumn("店舗",  width="small"),
            "氏名":     st.column_config.TextColumn("氏名",  width="medium"),
            "優先度":   st.column_config.NumberColumn("優先度",  width="small"),
            "運転":     st.column_config.TextColumn("運転可否", width="small"),
            "出勤時間": st.column_config.TextColumn("出勤", width="small"),
            "退勤時間": st.column_config.TextColumn("退勤", width="small"),
            "_sid":     st.column_config.TextColumn("_sid", disabled=True),
        },
        disabled=["店舗","氏名","優先度","運転","出勤時間","退勤時間","_sid"],
        column_order=["出勤","店舗","氏名","優先度","運転","出勤時間","退勤時間"],
        key="staff_editor",
    )

    attending_sids  = set(edited_staff.loc[edited_staff["出勤"] == True, "_sid"].tolist())
    attending_staff = [s for s in filtered_staff if s.staff_id in attending_sids]
    active_vehicles = [v for v in all_vehicles if any(s.shop == v.shop for s in attending_staff)]

    # サマリー
    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("本日 出席利用者", f"{len(attending_users)} 名")
    col_b.metric("本日 出勤スタッフ", f"{len(attending_staff)} 名")
    col_c.metric("稼働車両", f"{len(active_vehicles)} 台")
    col_d.metric("欠席",
                 f"{len(filtered_users) - len(attending_users)} 名",
                 delta=f"-{len(filtered_users) - len(attending_users)}" if len(filtered_users) != len(attending_users) else None,
                 delta_color="inverse")

    st.divider()

    # ================================================================
    # STEP 3: 最適化実行
    # ================================================================
    step_header(3, "最適化の実行", "迎え便・送り便を同時に計算します")

    if len(attending_users) == 0:
        st.warning("出席利用者が0名です。STEP 2 で確認してください。")
        return

    checker = ConstraintChecker()
    errors  = checker.validate(attending_users, active_vehicles, attending_staff)
    for e in errors:
        st.error(e)

    run_btn = st.button(
        f"🚀　{target_date.strftime('%m/%d')} の送迎ルートを最適化する　（{len(attending_users)}名）",
        disabled=bool(errors),
        type="primary",
        use_container_width=True,
    )

    if run_btn:
        with st.spinner(f"🔄 {target_date.strftime('%m/%d')} の迎え便・送り便を同時最適化中..."):
            pickup_routes  = run_all_shops(
                attending_users, active_vehicles, attending_staff,
                TripType.PICKUP,  pu_start, pu_limit,
            )
            dropoff_routes = run_all_shops(
                attending_users, active_vehicles, attending_staff,
                TripType.DROPOFF, do_start, do_limit,
            )

        st.session_state.update({
            "pickup_routes":  pickup_routes,
            "dropoff_routes": dropoff_routes,
            "result_date":    target_date,
        })
        n_pu = len(pickup_routes)
        n_do = len(dropoff_routes)
        st.success(
            f"✅ 最適化完了！　迎え便 **{n_pu}** ルート　送り便 **{n_do}** ルート"
        )

    st.divider()

    # ================================================================
    # STEP 4: 結果表示・ダウンロード
    # ================================================================
    step_header(4, "結果の確認とダウンロード", "送迎ルートを確認して印刷・配布")

    if "pickup_routes" not in st.session_state:
        st.info("👆 STEP 3 で最適化を実行すると結果が表示されます")
        return

    pickup_routes  = st.session_state["pickup_routes"]
    dropoff_routes = st.session_state["dropoff_routes"]
    result_date    = st.session_state.get("result_date")
    all_routes     = pickup_routes + dropoff_routes

    if not all_routes:
        st.error("ルートを生成できませんでした。データを確認してください。")
        return

    # メトリクス
    total_pu   = sum(len(r.stops) for r in pickup_routes)
    total_do   = sum(len(r.stops) for r in dropoff_routes)
    wc_count   = sum(1 for r in pickup_routes for s in r.stops if s["user"].wheelchair)
    shops_used = sorted(set(r.shop for r in all_routes))

    metric_row([
        (total_pu,        "名", "迎え便 乗車合計"),
        (total_do,        "名", "送り便 乗車合計"),
        (wc_count,        "名", "車椅子対応"),
        (len(shops_used), "店", "稼働店舗数"),
    ])

    # 制約検証
    with st.expander("🔍 制約条件 検証サマリー", expanded=False):
        forbidden = checker.get_forbidden_pairs(attending_users)
        all_ok    = True

        for label, routes in [("迎え便", pickup_routes), ("送り便", dropoff_routes)]:
            st.markdown(f"**{label}**")
            for route in routes:
                uin      = [s["user"] for s in route.stops]
                ok_cap   = len(uin) <= route.vehicle.capacity
                ok_wc    = not (any(u.wheelchair for u in uin) and not route.vehicle.wheelchair_ok)
                ok_incp  = not any(
                    tuple(sorted([u1.user_id, u2.user_id])) in forbidden
                    for i, u1 in enumerate(uin) for u2 in uin[i+1:]
                )
                ok_drv   = route.driver is not None and route.driver.can_drive
                ok_shop  = all(u.shop == route.shop for u in uin)

                # TimeWindow チェック（迎え便のみ）
                if route.trip_type == TripType.PICKUP:
                    ok_tw = all(s["arrival_min"] <= s["user"].pickup_latest for s in route.stops)
                else:
                    ok_tw = True  # 送り便は目安のみ

                all_ok = all_ok and all([ok_cap, ok_wc, ok_incp, ok_drv, ok_shop])

                def st_icon(ok): return '<span class="ok">✅</span>' if ok else '<span class="fail">❌ 違反</span>'
                dn = route.driver.name if route.driver else "未定"
                st.markdown(
                    f"　**{route.shop} - {route.vehicle.name}** ({len(uin)}/{route.vehicle.capacity}名)　"
                    f"定員:{st_icon(ok_cap)}　車椅子:{st_icon(ok_wc)}　"
                    f"同乗不可:{st_icon(ok_incp)}　混載禁止:{st_icon(ok_shop)}　"
                    f"TimeWindow:{st_icon(ok_tw)}　"
                    f"運転者:{dn} {st_icon(ok_drv)}",
                    unsafe_allow_html=True,
                )
        if all_ok:
            st.success("🎉 全制約条件をクリアしています！")

    # 結果テーブル（タブ）
    st.markdown("#### 📋 送迎ルート一覧")
    tab_pu, tab_do = st.tabs(["▶ 迎え便", "◀ 送り便"])

    col_cfg = {
        "順番":     st.column_config.NumberColumn(width="small"),
        "到着予定": st.column_config.TextColumn(width="small"),
        "車椅子":   st.column_config.TextColumn(width="small"),
        "優先度":   st.column_config.NumberColumn(width="small"),
    }

    with tab_pu:
        df_pu = routes_to_dataframe(pickup_routes)
        if not df_pu.empty:
            st.dataframe(df_pu, use_container_width=True, hide_index=True, column_config=col_cfg)
        else:
            st.info("迎え便のルートがありません")

    with tab_do:
        df_do = routes_to_dataframe(dropoff_routes)
        if not df_do.empty:
            st.dataframe(df_do, use_container_width=True, hide_index=True, column_config=col_cfg)
        else:
            st.info("送り便のルートがありません")

    # タイムラインプレビュー（迎え便のみ）
    with st.expander("🕐 タイムラインプレビュー（迎え便）", expanded=False):
        for route in sorted(pickup_routes, key=lambda r: r.shop):
            dn = route.driver.name if route.driver else "未定"
            st.markdown(f"**{route.shop} - {route.vehicle.name}** 　運転: {dn}")
            html = ""
            for stop in route.stops:
                h, m = divmod(stop["arrival_min"], 60)
                wc   = "♿ " if stop["user"].wheelchair else ""
                tw   = min_to_hhmm(stop["user"].pickup_latest)
                html += f"""
                <div class="timeline-item">
                  <div class="timeline-dot"></div>
                  <div class="timeline-time">{h:02d}:{m:02d}</div>
                  <div>
                    <div class="timeline-name">{wc}{stop["user"].name}</div>
                    <div class="timeline-detail">{stop["address"]}　リミット: {tw}</div>
                  </div>
                </div>
                """
            st.markdown(html, unsafe_allow_html=True)

    # ダウンロード
    st.markdown("<br>", unsafe_allow_html=True)
    col_dl, _ = st.columns([1, 2])
    with col_dl:
        excel_bytes = build_excel_output(pickup_routes, dropoff_routes, result_date)
        date_fname  = result_date.strftime("%Y%m%d") if result_date else "送迎ルート"
        ext   = "xlsx" if OPENPYXL_AVAILABLE else "csv"
        mime  = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                 if OPENPYXL_AVAILABLE else "text/csv")
        st.download_button(
            label=f"📥 送迎ルート表をダウンロード（{date_fname}）",
            data=excel_bytes,
            file_name=f"送迎ルート_{date_fname}.{ext}",
            mime=mime,
            use_container_width=True,
        )
        st.caption("迎え便・送り便が1ファイルの2シートで出力されます（印刷向けA4横）")

    st.divider()

    # 地図表示
    st.markdown("#### 🗺️ 送迎ルートマップ")
    tab_map1, tab_map2 = st.tabs(["▶ 迎え便", "◀ 送り便"])
    with tab_map1:
        render_map(pickup_routes)
    with tab_map2:
        render_map(dropoff_routes)


if __name__ == "__main__":
    main()
