"""
送迎ルート自動作成システム - Streamlit Webアプリ v4
=====================================================
放課後等デイサービス / 就労継続支援A型・B型 対応

v4 改修要件:
  【第1】カレンダーのレイアウト変更（行＝人、列＝日付）
  【第2】スタッフと利用者のカレンダーシートの分離
  【第3】利用者カレンダーの「店舗（事業所）」ごとのシート分割
  【第4】カレンダーの「時間（HH:MM-HH:MM）」直接入力の基本仕様化とパース処理
  【第5】openpyxlを用いたExcelフォーマットの視認性向上（幅調整・色・罫線）
"""

from __future__ import annotations

import io
import math
import datetime
import re
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
        Font, PatternFill, Alignment, Border, Side
    )
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ==========================================
# 1. データモデル
# ==========================================
class NodeType(Enum):
    DEPOT = "DEPOT"
    PICKUP = "PICKUP"
    DROPOFF = "DROPOFF"

@dataclass
class TimeWindow:
    start_m: int
    end_m: int

@dataclass
class FacilityInfo:
    name: str
    lat: float
    lon: float

@dataclass
class StaffInfo:
    id_str: str
    name: str
    priority: int = 1
    shift_tw: Optional[TimeWindow] = None

@dataclass
class UserInfo:
    id_str: str
    name: str
    lat: float
    lon: float
    wheelchair: bool = False
    service_time: int = 5
    facility: str = "デフォルト店舗"
    # 当日の個別時間制約（迎え・送り）
    pickup_tw: Optional[TimeWindow] = None
    dropoff_tw: Optional[TimeWindow] = None

@dataclass
class VRPNode:
    node_id: int
    ntype: NodeType
    lat: float
    lon: float
    name: str
    user_ref: Optional[UserInfo] = None
    tw: Optional[TimeWindow] = None

@dataclass
class VRPInstance:
    facility: FacilityInfo
    staff_list: list[StaffInfo]
    user_list: list[UserInfo]
    nodes: list[VRPNode] = field(default_factory=list)
    dist_matrix: list[list[int]] = field(default_factory=list)
    time_matrix: list[list[int]] = field(default_factory=list)


# ==========================================
# 2. ユーティリティ (時間変換・距離計算)
# ==========================================
def parse_time_str(t_str: str, default_val: int = 0) -> int:
    """ HH:MM 形式の文字列を 0:00 からの経過分数に変換 """
    if not isinstance(t_str, str):
        return default_val
    m = re.match(r"(\d{1,2}):(\d{2})", t_str.strip())
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return default_val

def extract_time_window(val: any, def_start: str, def_end: str) -> tuple[bool, str, str]:
    """
    セル入力値から時間を抽出する
    例: "09:00-15:30" -> (True, "09:00", "15:30")
    例: "〇" -> (True, def_start, def_end)
    空白 -> (False, "", "")
    """
    if pd.isna(val) or str(val).strip() == "":
        return False, def_start, def_end
    
    val_str = str(val).strip()
    # 時間文字列の直接入力をパース
    m = re.search(r'(\d{1,2}:\d{2})\s*[-~～]\s*(\d{1,2}:\d{2})', val_str)
    if m:
        return True, m.group(1), m.group(2)
    # フォールバック処理
    elif val_str == "〇" or val_str.lower() in ['true', 'yes', '1']:
        return True, def_start, def_end
    else:
        return False, def_start, def_end

def format_minutes(m_total: int) -> str:
    if m_total < 0:
        return "00:00"
    h = m_total // 60
    m = m_total % 60
    return f"{h:02d}:{m:02d}"

def haversine(lat1: float, lon1: float, lat2: float, lon2: float) -> float:
    R = 6371000.0
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2)**2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlam / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    return R * c

def compute_matrices(nodes: list[VRPNode]) -> tuple[list[list[int]], list[list[int]]]:
    n = len(nodes)
    dist_m = [[0]*n for _ in range(n)]
    time_m = [[0]*n for _ in range(n)]
    speed_m_per_min = 300.0 # 18km/h

    for i in range(n):
        for j in range(n):
            if i == j:
                continue
            d = haversine(nodes[i].lat, nodes[i].lon, nodes[j].lat, nodes[j].lon)
            # マンハッタン距離係数による補正
            d *= 1.3
            dist_m[i][j] = int(d)
            # 乗降時間加算 (iで乗降にかかる時間をi→jの移動時間に含める)
            add_service = 0
            if nodes[i].ntype != NodeType.DEPOT and nodes[i].user_ref is not None:
                add_service = nodes[i].user_ref.service_time
            
            t_min = d / speed_m_per_min
            time_m[i][j] = int(math.ceil(t_min)) + add_service
            
    return dist_m, time_m


# ==========================================
# 3. Excel生成 (オープンピクセルでの整形)
# ==========================================
def style_excel_sheet(ws, is_calendar: bool = False, month_days: int = 31):
    """ プロ仕様のExcel書式設定（幅、色、罫線） """
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    weekend_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                         top=Side(style='thin'), bottom=Side(style='thin'))

    # ヘッダーの装飾
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # セルの枠線と中央揃え
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border
            if cell.row > 1:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # カレンダー特有の列幅と週末色付け
    if is_calendar:
        ws.column_dimensions['A'].width = 10  # ID
        ws.column_dimensions['B'].width = 18  # 名前
        for col_idx in range(3, 3 + month_days):
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = 14
            # サンプルとして簡易的に土日想定列を色付け (例: 5,6, 12,13...)
            day = col_idx - 2
            if day % 7 in (5, 6):
                for row_idx in range(1, ws.max_row + 1):
                    ws.cell(row=row_idx, column=col_idx).fill = weekend_fill
    else:
        # マスタシートの幅自動調整
        for col in ws.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[col_letter].width = max_length + 4


def get_sample_excel() -> bytes:
    if not OPENPYXL_AVAILABLE:
        return b""
    wb = Workbook()
    wb.remove(wb.active)

    # ---- 1. 利用者マスタ ----
    ws_um = wb.create_sheet(title="マスタ_利用者")
    ws_um.append(["ID", "名前", "事業所", "緯度", "経度", "車椅子", "基本迎え時間", "基本送り時間"])
    users_data = [
        ["U01", "富山 太郎", "富山中央事業所", 36.6953, 137.2113, "", "08:30-09:30", "15:30-16:30"],
        ["U02", "高岡 花子", "富山中央事業所", 36.7011, 137.2155, "1", "08:30-09:30", "15:30-16:30"],
        ["U03", "呉羽 一郎", "富山中央事業所", 36.6899, 137.1999, "", "08:30-09:30", "15:30-16:30"],
        ["U04", "新湊 次郎", "高岡南事業所", 36.7444, 137.0111, "", "09:00-10:00", "16:00-17:00"],
        ["U05", "氷見 桜", "高岡南事業所", 36.7555, 137.0222, "", "09:00-10:00", "16:00-17:00"],
    ]
    for r in users_data: ws_um.append(r)
    style_excel_sheet(ws_um)

    # ---- 2. スタッフマスタ ----
    ws_sm = wb.create_sheet(title="マスタ_スタッフ")
    ws_sm.append(["ID", "名前", "優先度", "基本出勤時間", "基本退勤時間"])
    staff_data = [
        ["S01", "送迎 山田", 1, "08:00", "18:00"],
        ["S02", "送迎 佐藤", 2, "08:30", "17:30"],
        ["S03", "臨時 鈴木", 3, "08:00", "18:00"],
    ]
    for r in staff_data: ws_sm.append(r)
    style_excel_sheet(ws_sm)

    # ---- 3. 事業所マスタ ----
    ws_fm = wb.create_sheet(title="マスタ_事業所")
    ws_fm.append(["事業所名", "緯度", "経度"])
    fac_data = [
        ["富山中央事業所", 36.69595, 137.21368],
        ["高岡南事業所", 36.73333, 137.01667],
    ]
    for r in fac_data: ws_fm.append(r)
    style_excel_sheet(ws_fm)

    # カレンダー用共通ヘッダー
    cal_header = ["ID", "名前"] + [str(i) for i in range(1, 32)]

    # ---- 4. シフト_スタッフ ----
    ws_sc = wb.create_sheet(title="シフト_スタッフ")
    ws_sc.append(cal_header)
    for row in [
        ["S01", "送迎 山田", "08:00-18:00", "", "08:00-18:00"],
        ["S02", "送迎 佐藤", "08:30-17:30", "08:30-17:30", "08:30-17:30"],
        ["S03", "臨時 鈴木", "", "08:00-18:00", ""],
    ]:
        ws_sc.append(row + ["08:00-18:00"] * 28)
    style_excel_sheet(ws_sc, is_calendar=True)

    # ---- 5. 予定_富山中央事業所 ----
    ws_uc1 = wb.create_sheet(title="予定_富山中央事業所")
    ws_uc1.append(cal_header)
    for row in [
        ["U01", "富山 太郎", "08:30-15:30", "08:30-15:30", "08:30-15:30"],
        ["U02", "高岡 花子", "09:00-16:00", "", "09:00-16:00"],
        ["U03", "呉羽 一郎", "08:30-15:30", "08:30-15:30", "08:30-15:30"],
    ]:
        ws_uc1.append(row + ["08:30-15:30"] * 28)
    style_excel_sheet(ws_uc1, is_calendar=True)

    # ---- 6. 予定_高岡南事業所 ----
    ws_uc2 = wb.create_sheet(title="予定_高岡南事業所")
    ws_uc2.append(cal_header)
    for row in [
        ["U04", "新湊 次郎", "09:00-17:00", "09:00-17:00", ""],
        ["U05", "氷見 桜", "09:00-17:00", "09:00-17:00", "09:00-17:00"],
    ]:
        ws_uc2.append(row + ["09:00-17:00"] * 28)
    style_excel_sheet(ws_uc2, is_calendar=True)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ==========================================
# 4. データパース
# ==========================================
def parse_excel_upload(file_bytes: bytes, target_date: datetime.date) -> tuple[dict, list[StaffInfo], list[UserInfo]]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names

    # マスタ読み込み
    df_um = xls.parse([s for s in sheet_names if "利用者" in s and "マスタ" in s][0])
    df_sm = xls.parse([s for s in sheet_names if "スタッフ" in s and "マスタ" in s][0])
    df_fm = xls.parse([s for s in sheet_names if "事業所" in s and "マスタ" in s][0])

    fac_dict = {}
    for _, r in df_fm.iterrows():
        fac_dict[str(r.get("事業所名", "")).strip()] = FacilityInfo(
            name=str(r.get("事業所名", "")).strip(),
            lat=float(r["緯度"]),
            lon=float(r["経度"])
        )

    # スタッフ解析
    staff_list = []
    # シフトシートの検索
    shift_sheets = [s for s in sheet_names if "シフト" in s or ("カレンダー" in s and "スタッフ" in s)]
    df_shift = xls.parse(shift_sheets[0]) if shift_sheets else pd.DataFrame()
    day_col = str(target_date.day)
    # 念のため整数型カラムへのフォールバック
    if target_date.day in df_shift.columns:
        day_col = target_date.day

    for _, r in df_sm.iterrows():
        s_id = str(r["ID"])
        def_s = str(r.get("基本出勤時間", "08:00"))
        def_e = str(r.get("基本退勤時間", "18:00"))
        
        tw = None
        if not df_shift.empty and "ID" in df_shift.columns and day_col in df_shift.columns:
            shift_row = df_shift[df_shift["ID"].astype(str) == s_id]
            if not shift_row.empty:
                val = shift_row.iloc[0][day_col]
                is_active, st_str, ed_str = extract_time_window(val, def_s, def_e)
                if is_active:
                    tw = TimeWindow(parse_time_str(st_str), parse_time_str(ed_str))
        
        if tw:
            staff_list.append(StaffInfo(
                id_str=s_id,
                name=str(r["名前"]),
                priority=int(r.get("優先度", 1)),
                shift_tw=tw
            ))

    # 利用者解析
    user_list = []
    # 予定シートの検索（複数店舗対応）
    plan_sheets = [s for s in sheet_names if "予定" in s or ("カレンダー" in s and "利用者" in s)]
    df_plans = {}
    for ps in plan_sheets:
        df_plans[ps] = xls.parse(ps)

    for _, r in df_um.iterrows():
        u_id = str(r["ID"])
        fac_name = str(r.get("事業所", "デフォルト店舗")).strip()
        
        # マスタからデフォルトの時間枠を分割
        def_p_tw = str(r.get("基本迎え時間", "08:30-09:30"))
        def_d_tw = str(r.get("基本送り時間", "15:30-16:30"))
        
        def_p_s = def_p_tw.split("-")[0] if "-" in def_p_tw else "08:00"
        def_p_e = def_p_tw.split("-")[1] if "-" in def_p_tw else "10:00"
        def_d_s = def_d_tw.split("-")[0] if "-" in def_d_tw else "15:00"
        def_d_e = def_d_tw.split("-")[1] if "-" in def_d_tw else "18:00"

        # 当該利用者が属する事業所の予定シートを探す
        target_sheet = None
        for ps, df_p in df_plans.items():
            if not df_p.empty and "ID" in df_p.columns:
                if not df_p[df_p["ID"].astype(str) == u_id].empty:
                    target_sheet = df_p
                    break

        p_tw_obj, d_tw_obj = None, None
        
        if target_sheet is not None and day_col in target_sheet.columns:
            plan_row = target_sheet[target_sheet["ID"].astype(str) == u_id]
            val = plan_row.iloc[0][day_col]
            
            is_active, act_s, act_e = extract_time_window(val, def_p_s, def_d_e)
            if is_active:
                # 迎え枠（開始〜開始＋2時間などを仮設定、あるいは直接パース結果を利用）
                p_tw_obj = TimeWindow(parse_time_str(act_s), parse_time_str(act_s) + 120)
                # 送り枠（終了ー2時間〜終了などを仮設定）
                d_tw_obj = TimeWindow(parse_time_str(act_e) - 120, parse_time_str(act_e))

        if p_tw_obj and d_tw_obj:
            user_list.append(UserInfo(
                id_str=u_id,
                name=str(r["名前"]),
                lat=float(r["緯度"]),
                lon=float(r["経度"]),
                wheelchair=bool(r.get("車椅子", False)),
                facility=fac_name,
                pickup_tw=p_tw_obj,
                dropoff_tw=d_tw_obj
            ))

    return fac_dict, staff_list, user_list


# ==========================================
# 5. VRP 最適化ロジック (V3継承)
# ==========================================
def solve_vrp(inst: VRPInstance, is_pickup: bool) -> list[list[dict]]:
    if not ORTOOLS_AVAILABLE or not inst.user_list:
        return []

    # ノード構築 (0: デポ)
    nodes = [VRPNode(0, NodeType.DEPOT, inst.facility.lat, inst.facility.lon, inst.facility.name)]
    
    for u in inst.user_list:
        tw = u.pickup_tw if is_pickup else u.dropoff_tw
        nodes.append(VRPNode(len(nodes), NodeType.PICKUP if is_pickup else NodeType.DROPOFF, 
                             u.lat, u.lon, u.name, user_ref=u, tw=tw))
        
    inst.nodes = nodes
    dist_m, time_m = compute_matrices(nodes)
    
    num_vehicles = len(inst.staff_list)
    manager = pywrapcp.RoutingIndexManager(len(nodes), num_vehicles, 0)
    routing = pywrapcp.RoutingModel(manager)

    def time_callback(from_index, to_index):
        f = manager.IndexToNode(from_index)
        t = manager.IndexToNode(to_index)
        return time_m[f][t]

    trans_idx = routing.RegisterTransitCallback(time_callback)
    routing.SetArcCostEvaluatorOfAllVehicles(trans_idx)

    # 時間制約ディメンション
    routing.AddDimension(
        trans_idx,
        30,  # 待機許容時間
        1440, # 1日の最大分数
        False, 
        "Time"
    )
    time_dim = routing.GetDimensionOrDie("Time")

    # Time Windows設定
    for i in range(1, len(nodes)):
        idx = manager.NodeToIndex(i)
        tw = nodes[i].tw
        if tw:
            time_dim.CumulVar(idx).SetRange(tw.start_m, tw.end_m)

    # 車両（スタッフ）のシフト制約と優先度コスト
    for v_idx in range(num_vehicles):
        idx_start = routing.Start(v_idx)
        idx_end = routing.End(v_idx)
        staff = inst.staff_list[v_idx]
        
        if staff.shift_tw:
            time_dim.CumulVar(idx_start).SetRange(staff.shift_tw.start_m, staff.shift_tw.end_m)
            time_dim.CumulVar(idx_end).SetRange(staff.shift_tw.start_m, staff.shift_tw.end_m)
        
        # 固定コストによる分散（優先度1=安価、2=高価）
        routing.SetFixedCostOfVehicle(staff.priority * 1000, v_idx)

    search_parameters = pywrapcp.DefaultRoutingSearchParameters()
    search_parameters.first_solution_strategy = routing_enums_pb2.FirstSolutionStrategy.PATH_CHEAPEST_ARC
    search_parameters.time_limit.seconds = 5

    solution = routing.SolveWithParameters(search_parameters)

    routes = []
    if solution:
        for v_idx in range(num_vehicles):
            idx = routing.Start(v_idx)
            route = []
            while not routing.IsEnd(idx):
                node_idx = manager.IndexToNode(idx)
                n = nodes[node_idx]
                t_var = time_dim.CumulVar(idx)
                arr_min = solution.Min(t_var)
                arr_max = solution.Max(t_var)
                route.append({
                    "node": n,
                    "arr_min": arr_min,
                    "arr_max": arr_max,
                    "staff": inst.staff_list[v_idx].name
                })
                idx = solution.Value(routing.NextVar(idx))
            
            node_idx = manager.IndexToNode(idx)
            n = nodes[node_idx]
            t_var = time_dim.CumulVar(idx)
            route.append({
                "node": n,
                "arr_min": solution.Min(t_var),
                "arr_max": solution.Max(t_var),
                "staff": inst.staff_list[v_idx].name
            })
            
            if len(route) > 2:
                routes.append(route)
    return routes


# ==========================================
# 6. Streamlit UI
# ==========================================
def main():
    st.set_page_config(page_title="送迎ルート自動作成 v4", layout="wide")
    st.title("🚗 送迎ルート自動作成システム v4")

    st.markdown("""
    **【v4の新機能】**
    * カレンダーのレイアウトを「行＝人名、列＝日付」に変更
    * スタッフシフト表と利用者予定表を別シートに分離
    * 利用者予定表の店舗別シート管理に対応
    * カレンダーでの時間入力（HH:MM-HH:MM）に標準対応
    """)

    # 1. テンプレートダウンロード
    with st.expander("📝 1. Excelテンプレートのダウンロード", expanded=False):
        st.write("設定用のExcelファイルをダウンロードします。")
        if OPENPYXL_AVAILABLE:
            st.download_button(
                label="📥 テンプレートExcelダウンロード",
                data=get_sample_excel(),
                file_name="送迎マスタ_カレンダー.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # 2. アップロードと対象日選択
    st.markdown("### 2. データアップロードと対象日指定")
    col_up, col_dt = st.columns([2, 1])
    
    with col_up:
        uploaded_file = st.file_uploader("設定済みExcelファイルをアップロード", type=["xlsx"])
    with col_dt:
        target_date = st.date_input("作成対象日", value=datetime.date.today())

    if not uploaded_file:
        st.info("Excelファイルをアップロードしてください。")
        return

    # パース処理
    try:
        fac_dict, staff_list, user_list = parse_excel_upload(uploaded_file.getvalue(), target_date)
    except Exception as e:
        st.error(f"ファイル読み込みエラー: {e}")
        return

    st.success(f"{target_date.strftime('%Y年%m月%d日')} のデータを読み込みました。"
               f"（事業所: {len(fac_dict)}件、スタッフ: {len(staff_list)}名、利用者: {len(user_list)}名）")

    # 当日欠席調整 UI (V3継承)
    with st.expander("⚙️ 当日の欠席・追加調整（オプション）", expanded=False):
        df_edit = pd.DataFrame([{"ID": u.id_str, "名前": u.name, "事業所": u.facility, "利用": True} for u in user_list])
        edited_df = st.data_editor(df_edit, use_container_width=True, hide_index=True)
        active_ids = edited_df[edited_df["利用"] == True]["ID"].tolist()
        user_list = [u for u in user_list if u.id_str in active_ids]

    # 計算実行
    if st.button("🚀 ルート最適化を実行", type="primary"):
        with st.spinner("計算中..."):
            all_pickup_routes = []
            all_dropoff_routes = []

            # 事業所ごとに分離してVRPを解く
            for fac_name, fac_info in fac_dict.items():
                u_in_fac = [u for u in user_list if u.facility == fac_name]
                if not u_in_fac:
                    continue
                
                inst = VRPInstance(facility=fac_info, staff_list=staff_list, user_list=u_in_fac)
                p_routes = solve_vrp(inst, is_pickup=True)
                d_routes = solve_vrp(inst, is_pickup=False)
                
                # 事業所名を付与してリスト結合
                for r in p_routes:
                    r[0]["fac_name"] = fac_name
                    all_pickup_routes.append(r)
                for r in d_routes:
                    r[0]["fac_name"] = fac_name
                    all_dropoff_routes.append(r)

            st.session_state["p_routes"] = all_pickup_routes
            st.session_state["d_routes"] = all_dropoff_routes

    # 結果表示
    if "p_routes" in st.session_state:
        st.markdown("---")
        st.subheader("📋 計算結果")
        
        tab1, tab2 = st.tabs(["迎え便 ルート", "送り便 ルート"])
        
        def render_routes(routes):
            if not routes:
                st.warning("有効なルートが生成されませんでした。")
                return
            for i, r in enumerate(routes):
                fac_lbl = r[0].get("fac_name", "")
                st.markdown(f"**号車 {i+1} : {r[0]['staff']}** （{fac_lbl}）")
                for step in r:
                    n = step["node"]
                    arr = format_minutes(step["arr_min"])
                    if n.ntype == NodeType.DEPOT:
                        st.markdown(f"- 🏠 {arr} {n.name} (事業所)")
                    else:
                        wc = "🦽 " if getattr(n.user_ref, "wheelchair", False) else "👤 "
                        st.markdown(f"- {wc} {arr} {n.name} 様")
                st.markdown("---")

        with tab1:
            render_routes(st.session_state["p_routes"])
        with tab2:
            render_routes(st.session_state["d_routes"])

if __name__ == "__main__":
    main()
