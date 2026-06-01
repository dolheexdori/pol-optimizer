import math
import os
import random
from typing import Any, Dict, List, Optional, Tuple

import folium
import altair as alt
import pandas as pd
import streamlit as st
from folium.features import DivIcon
from streamlit_folium import st_folium


st.set_page_config(
    page_title="POL 수송 최적화",
    page_icon="⛽",
    layout="wide",
)

DEPOT_ID = "DEPOT"
APP_VERSION = "작전거리 기반 UI v8.0"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BASE_XLSX_PATH = os.path.join(BASE_DIR, "기지_기본정보.xlsx")
STATUS_XLSX_PATH = os.path.join(BASE_DIR, "시나리오별_유류상태.xlsx")
SCENARIO_XLSX_PATH = os.path.join(BASE_DIR, "시나리오_설명.xlsx")

OPERATIONAL_DISTANCE_FILE_CANDIDATES = [
    "operational_distance_info.xlsx",
    "기지간_작전거리정보.xlsx",
    "기지간_작전거리정보_재생성.xlsx",
]

DEFAULT_OPERATION_SETTINGS = {
    "시간가중치": 0.3,
    "위험도패널티": 20.0,
    "통행불가패널티": 100000.0,
    "기본평균속도kmh": 55.0,
}


# =========================================================
# 데이터 로드 / 전처리
# =========================================================
def get_operational_distance_path() -> Optional[str]:
    for file_name in OPERATIONAL_DISTANCE_FILE_CANDIDATES:
        path = os.path.join(BASE_DIR, file_name)
        if os.path.exists(path):
            return path
    return None


def file_signature() -> Tuple[float, ...]:
    required_paths = [BASE_XLSX_PATH, STATUS_XLSX_PATH, SCENARIO_XLSX_PATH]
    operational_path = get_operational_distance_path()
    if operational_path:
        required_paths.append(operational_path)

    return tuple(os.path.getmtime(path) if os.path.exists(path) else 0 for path in required_paths)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all").copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace(r"\s+", "", regex=True)
        .str.replace("(L)", "L", regex=False)
        .str.strip()
    )

    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].astype(str).str.strip()

    return df


def rename_aliases(base_df: pd.DataFrame, status_df: pd.DataFrame, scenario_df: pd.DataFrame):
    base_aliases = {
        "최대저장량L": "최대저장량",
        "최대저장량리터": "최대저장량",
        "최대유류량": "최대저장량",
        "저장용량": "최대저장량",
        "위도좌표": "위도",
        "경도좌표": "경도",
    }

    status_aliases = {
        "유조차수": "가용유조차수",
        "가용수송자산": "가용유조차수",
        "가용유조차": "가용유조차수",
        "유조차적재량L": "유조차적재량",
        "유조차적재량리터": "유조차적재량",
    }

    scenario_aliases = {
        "상황가정": "상황설명",
        "상황설명": "상황설명",
        "시나리오설명": "상황설명",
        "분석목적": "분석목표",
        "분석목표": "분석목표",
    }

    return (
        base_df.rename(columns=base_aliases),
        status_df.rename(columns=status_aliases),
        scenario_df.rename(columns=scenario_aliases),
    )


def require_columns(df: pd.DataFrame, required: List[str], file_name: str):
    missing = [col for col in required if col not in df.columns]
    if missing:
        st.error(f"{file_name} 파일에 필요한 열이 없습니다.")
        st.write("누락된 열:", missing)
        st.write("현재 인식된 열:", df.columns.tolist())
        st.stop()


def to_number(df: pd.DataFrame, columns: List[str], file_name: str) -> pd.DataFrame:
    df = df.copy()
    for col in columns:
        if col not in df.columns:
            continue
        df[col] = pd.to_numeric(df[col], errors="coerce")
        if df[col].isna().any():
            st.error(f"{file_name} 파일의 '{col}' 열에 숫자로 바꿀 수 없는 값이 있습니다.")
            st.dataframe(df[df[col].isna()], use_container_width=True)
            st.stop()
    return df


@st.cache_data(show_spinner="엑셀 데이터를 불러오는 중입니다.")
def load_data(_signature: Tuple[float, ...]):
    try:
        base_df = pd.read_excel(BASE_XLSX_PATH, engine="openpyxl")
        status_df = pd.read_excel(STATUS_XLSX_PATH, engine="openpyxl")
        scenario_df = pd.read_excel(SCENARIO_XLSX_PATH, engine="openpyxl")
    except FileNotFoundError as e:
        st.error("엑셀 파일을 찾을 수 없습니다. 코드 파일과 엑셀 3개를 같은 폴더에 두어야 합니다.")
        st.write("필요 파일: 기지_기본정보.xlsx / 시나리오별_유류상태.xlsx / 시나리오_설명.xlsx")
        st.exception(e)
        st.stop()

    base_df = normalize_columns(base_df)
    status_df = normalize_columns(status_df)
    scenario_df = normalize_columns(scenario_df)
    base_df, status_df, scenario_df = rename_aliases(base_df, status_df, scenario_df)

    require_columns(base_df, ["기지ID", "기지명", "위도", "경도", "최대저장량"], "기지_기본정보.xlsx")
    require_columns(
        status_df,
        ["시나리오", "기지ID", "현재유류량", "일일소모량", "가용유조차수", "유조차적재량"],
        "시나리오별_유류상태.xlsx",
    )
    require_columns(scenario_df, ["시나리오"], "시나리오_설명.xlsx")

    for df in [base_df, status_df, scenario_df]:
        if "기지ID" in df.columns:
            df["기지ID"] = df["기지ID"].astype(str).str.strip()
        if "시나리오" in df.columns:
            df["시나리오"] = df["시나리오"].astype(str).str.strip()

    base_df = to_number(base_df, ["위도", "경도", "최대저장량"], "기지_기본정보.xlsx")
    status_df = to_number(
        status_df,
        ["현재유류량", "일일소모량", "가용유조차수", "유조차적재량"],
        "시나리오별_유류상태.xlsx",
    )

    if DEPOT_ID not in base_df["기지ID"].tolist():
        st.error("기지_기본정보.xlsx 파일에 DEPOT 행이 필요합니다.")
        st.stop()

    return base_df, status_df, scenario_df


def read_operation_settings(file_path: Optional[str]) -> Dict[str, float]:
    settings = DEFAULT_OPERATION_SETTINGS.copy()
    if not file_path:
        return settings

    try:
        sheet_names = pd.ExcelFile(file_path, engine="openpyxl").sheet_names
        if "설정" not in sheet_names:
            return settings
        setting_df = pd.read_excel(file_path, sheet_name="설정", engine="openpyxl")
        setting_df = normalize_columns(setting_df)
        if "항목" not in setting_df.columns or "값" not in setting_df.columns:
            return settings

        for _, row in setting_df.iterrows():
            key = str(row["항목"]).strip()
            value = pd.to_numeric(row["값"], errors="coerce")
            if pd.notna(value):
                if key in settings:
                    settings[key] = float(value)
                elif key == "평균주행속도kmh":
                    settings["기본평균속도kmh"] = float(value)
    except Exception:
        return settings

    return settings


@st.cache_data(show_spinner=False)
def load_operational_distance_info(_signature: Tuple[float, ...]):
    file_path = get_operational_distance_path()
    settings = read_operation_settings(file_path)

    if not file_path:
        return pd.DataFrame(), settings, None

    try:
        xls = pd.ExcelFile(file_path, engine="openpyxl")
        sheet_name = "기지간_작전거리정보" if "기지간_작전거리정보" in xls.sheet_names else xls.sheet_names[0]
        op_df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
        op_df = normalize_columns(op_df)
    except Exception as e:
        st.warning(f"작전거리정보 파일을 읽지 못했습니다. 직선거리 기반 추정값으로 대체합니다. 오류: {e}")
        return pd.DataFrame(), settings, None

    aliases = {
        "출발ID": "출발기지ID",
        "도착ID": "도착기지ID",
        "거리km": "도로거리km",
        "작전거리km": "도로거리km",
        "소요시간분": "예상소요시간분",
        "예상시간분": "예상소요시간분",
        "통행가능": "통행가능여부",
        "위험도": "도로위험도",
        "비용점수": "작전비용점수",
    }
    op_df = op_df.rename(columns=aliases)

    required = ["출발기지ID", "도착기지ID", "도로거리km", "예상소요시간분", "통행가능여부", "도로위험도"]
    missing = [col for col in required if col not in op_df.columns]
    if missing:
        st.warning(
            "작전거리정보 파일의 필수 열이 부족합니다. 직선거리 기반 추정값으로 대체합니다. "
            f"누락 열: {missing}"
        )
        return pd.DataFrame(), settings, file_path

    op_df["출발기지ID"] = op_df["출발기지ID"].astype(str).str.strip()
    op_df["도착기지ID"] = op_df["도착기지ID"].astype(str).str.strip()
    op_df["통행가능여부"] = op_df["통행가능여부"].astype(str).str.upper().str.strip()

    numeric_cols = ["도로거리km", "예상소요시간분", "도로위험도"]
    if "작전비용점수" in op_df.columns:
        numeric_cols.append("작전비용점수")

    for col in numeric_cols:
        op_df[col] = pd.to_numeric(op_df[col], errors="coerce")

    op_df = op_df.dropna(subset=["출발기지ID", "도착기지ID", "도로거리km", "예상소요시간분", "도로위험도"])

    if "작전비용점수" not in op_df.columns or op_df["작전비용점수"].isna().any():
        blocked_penalty = op_df["통행가능여부"].ne("Y").astype(float) * settings["통행불가패널티"]
        op_df["작전비용점수"] = (
            op_df["도로거리km"]
            + op_df["예상소요시간분"] * settings["시간가중치"]
            + op_df["도로위험도"] * settings["위험도패널티"]
            + blocked_penalty
        )

    return op_df, settings, file_path


# =========================================================
# 작전거리 / 경로 계산
# =========================================================
def haversine_km(lat1, lon1, lat2, lon2) -> float:
    earth_radius_km = 6371.0
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat / 2) ** 2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon / 2) ** 2
    return earth_radius_km * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


@st.cache_data(show_spinner=False)
def build_metric_matrix(
    base_records: Tuple[Tuple[str, float, float], ...],
    operational_records: Tuple[Tuple[str, str, float, float, str, float, float], ...],
    settings: Dict[str, float],
):
    matrix: Dict[str, Dict[str, Dict[str, Any]]] = {}

    for from_id, from_lat, from_lon in base_records:
        matrix[from_id] = {}
        for to_id, to_lat, to_lon in base_records:
            if from_id == to_id:
                distance = 0.0
            else:
                distance = haversine_km(from_lat, from_lon, to_lat, to_lon) * 1.25
            time_min = distance / max(settings.get("기본평균속도kmh", 55.0), 1.0) * 60
            risk = 1.0 if from_id != to_id else 0.0
            cost = distance + time_min * settings["시간가중치"] + risk * settings["위험도패널티"]
            matrix[from_id][to_id] = {
                "distance_km": distance,
                "time_min": time_min,
                "risk": risk,
                "passable": True,
                "cost": cost,
                "source": "직선거리추정",
            }

    has_operational_data = len(operational_records) > 0

    for start_id, end_id, road_distance, time_min, passable_value, road_risk, operation_cost in operational_records:
        start_id = str(start_id).strip()
        end_id = str(end_id).strip()
        passable = str(passable_value).upper().strip() == "Y"
        cost = float(operation_cost)

        if start_id in matrix and end_id in matrix[start_id]:
            matrix[start_id][end_id] = {
                "distance_km": float(road_distance),
                "time_min": float(time_min),
                "risk": float(road_risk),
                "passable": passable,
                "cost": cost,
                "source": "작전거리정보",
            }

        if end_id in matrix and start_id in matrix[end_id] and matrix[end_id][start_id]["source"] == "직선거리추정":
            matrix[end_id][start_id] = {
                "distance_km": float(road_distance),
                "time_min": float(time_min),
                "risk": float(road_risk),
                "passable": passable,
                "cost": cost,
                "source": "작전거리정보",
            }

    matrix_source = "작전거리정보 파일 사용" if has_operational_data else "직선거리 기반 추정값 사용"
    return matrix, matrix_source


def route_segments(route: List[str]) -> List[Tuple[str, str]]:
    if not route:
        return []
    path = [DEPOT_ID] + route + [DEPOT_ID]
    return [(path[i], path[i + 1]) for i in range(len(path) - 1)]


def route_distance(route: List[str], matrix: Dict[str, Dict[str, Dict[str, Any]]]) -> float:
    return sum(matrix[start][end]["distance_km"] for start, end in route_segments(route))


def route_time(route: List[str], matrix: Dict[str, Dict[str, Dict[str, Any]]]) -> float:
    return sum(matrix[start][end]["time_min"] for start, end in route_segments(route))


def route_operation_cost(route: List[str], matrix: Dict[str, Dict[str, Dict[str, Any]]]) -> float:
    return sum(matrix[start][end]["cost"] for start, end in route_segments(route))


def route_blocked_count(route: List[str], matrix: Dict[str, Dict[str, Dict[str, Any]]]) -> int:
    return sum(0 if matrix[start][end]["passable"] else 1 for start, end in route_segments(route))


def route_average_risk(route: List[str], matrix: Dict[str, Dict[str, Dict[str, Any]]]) -> float:
    segments = route_segments(route)
    if not segments:
        return 0.0
    return sum(matrix[start][end]["risk"] for start, end in segments) / len(segments)


def total_distance(routes: List[List[str]], matrix: Dict[str, Dict[str, Dict[str, Any]]]) -> float:
    return sum(route_distance(route, matrix) for route in routes)


def total_time(routes: List[List[str]], matrix: Dict[str, Dict[str, Dict[str, Any]]]) -> float:
    return sum(route_time(route, matrix) for route in routes)


def total_operation_cost(routes: List[List[str]], matrix: Dict[str, Dict[str, Dict[str, Any]]]) -> float:
    return sum(route_operation_cost(route, matrix) for route in routes)


def total_blocked_count(routes: List[List[str]], matrix: Dict[str, Dict[str, Dict[str, Any]]]) -> int:
    return sum(route_blocked_count(route, matrix) for route in routes)


# =========================================================
# 우선순위 / 작전위험점수 / 시나리오 데이터
# =========================================================
def priority_label(days: float, urgent_threshold: float, caution_threshold: float) -> str:
    if days < urgent_threshold:
        return "긴급"
    if days < caution_threshold:
        return "주의"
    return "안정"


def priority_weight(days: float, urgent_threshold: float, caution_threshold: float) -> int:
    if days < urgent_threshold:
        return 3
    if days < caution_threshold:
        return 2
    return 1


def calculate_operation_risk_score(row: pd.Series, max_demand: float, max_consumption: float, caution_threshold: float) -> float:
    days = float(row["잔여가능일수"])
    demand = float(row["보급필요량"])
    consumption = float(row["일일소모량"])

    days_risk = max(0.0, min(1.0, (caution_threshold - days) / caution_threshold))
    demand_risk = demand / max_demand if max_demand > 0 else 0.0
    consumption_risk = consumption / max_consumption if max_consumption > 0 else 0.0

    score = days_risk * 50 + demand_risk * 30 + consumption_risk * 20
    return round(min(score, 100), 1)


def prepare_scenario(
    base_df: pd.DataFrame,
    status_df: pd.DataFrame,
    scenario: str,
    target_ratio: float,
    urgent_threshold: float,
    caution_threshold: float,
):
    selected = status_df[status_df["시나리오"] == scenario].copy()
    merged = selected.merge(
        base_df[["기지ID", "기지명", "위도", "경도", "최대저장량"]],
        on="기지ID",
        how="left",
    )

    if merged["기지명"].isna().any():
        st.error("시나리오별_유류상태.xlsx의 기지ID가 기지_기본정보.xlsx와 매칭되지 않습니다.")
        st.dataframe(merged[merged["기지명"].isna()], use_container_width=True)
        st.stop()

    if (merged["일일소모량"] <= 0).any():
        st.error("일일소모량은 0보다 커야 합니다.")
        st.dataframe(merged[merged["일일소모량"] <= 0], use_container_width=True)
        st.stop()

    merged["잔여가능일수"] = merged["현재유류량"] / merged["일일소모량"]
    merged["우선순위"] = merged["잔여가능일수"].apply(
        lambda days: priority_label(days, urgent_threshold, caution_threshold)
    )
    merged["보급필요량"] = (merged["최대저장량"] * target_ratio - merged["현재유류량"]).clip(lower=0)

    max_demand = float(merged["보급필요량"].max()) if not merged.empty else 0.0
    max_consumption = float(merged["일일소모량"].max()) if not merged.empty else 0.0
    merged["작전위험점수"] = merged.apply(
        lambda row: calculate_operation_risk_score(row, max_demand, max_consumption, caution_threshold),
        axis=1,
    )

    targets = merged[merged["보급필요량"] > 0].copy()
    return merged, targets


def status_dict(targets: pd.DataFrame) -> Dict[str, Dict[str, float]]:
    return {
        str(row["기지ID"]): {
            "잔여가능일수": float(row["잔여가능일수"]),
            "보급필요량": float(row["보급필요량"]),
            "우선순위": row["우선순위"],
            "작전위험점수": float(row["작전위험점수"]),
        }
        for _, row in targets.iterrows()
    }


# =========================================================
# 유전 알고리즘
# =========================================================
def fitness(
    route: List[str],
    matrix: Dict[str, Dict[str, Dict[str, Any]]],
    info: Dict[str, Dict[str, float]],
    urgent_threshold: float,
    caution_threshold: float,
    objective_mode: str,
) -> float:
    distance_score = route_distance(route, matrix)
    cost_score = route_operation_cost(route, matrix)
    order_penalty = 0.0
    n = max(len(route), 1)

    for idx, base_id in enumerate(route):
        days = info[base_id]["잔여가능일수"]
        demand = info[base_id]["보급필요량"]
        risk = info[base_id]["작전위험점수"]
        order_ratio = (idx + 1) / n
        order_penalty += priority_weight(days, urgent_threshold, caution_threshold) * order_ratio * 100
        order_penalty += risk * order_ratio * 8
        order_penalty += demand / 100000 * order_ratio * 20

    if objective_mode == "거리 최소화":
        return distance_score + order_penalty * 0.25
    if objective_mode == "긴급기지 우선":
        return cost_score + order_penalty * 1.8
    if objective_mode == "차량 부하 균등":
        return cost_score + order_penalty * 0.9
    return cost_score + order_penalty


def ordered_crossover(parent1: List[str], parent2: List[str]) -> List[str]:
    if len(parent1) <= 2:
        return parent1[:]

    start, end = sorted(random.sample(range(len(parent1)), 2))
    child = [None] * len(parent1)
    child[start : end + 1] = parent1[start : end + 1]
    rest = [gene for gene in parent2 if gene not in child]
    rest_idx = 0

    for i in range(len(child)):
        if child[i] is None:
            child[i] = rest[rest_idx]
            rest_idx += 1

    return child


def mutate(route: List[str], rate: float) -> List[str]:
    route = route[:]
    if len(route) >= 2 and random.random() < rate:
        i, j = random.sample(range(len(route)), 2)
        route[i], route[j] = route[j], route[i]
    return route


def genetic_algorithm(
    base_ids: List[str],
    matrix: Dict[str, Dict[str, Dict[str, Any]]],
    info: Dict[str, Dict[str, float]],
    population_size: int,
    generations: int,
    mutation_rate: float,
    urgent_threshold: float,
    caution_threshold: float,
    objective_mode: str,
    seed: int = 42,
):
    random.seed(seed)
    base_ids = [str(x) for x in base_ids]

    if len(base_ids) <= 1:
        return base_ids, [
            {
                "세대": 1,
                "최적작전거리(km)": route_distance(base_ids, matrix),
                "최적작전비용": fitness(base_ids, matrix, info, urgent_threshold, caution_threshold, objective_mode),
            }
        ]

    population_size = max(int(population_size), 20)
    generations = max(int(generations), 1)
    elite_size = max(2, min(5, population_size // 10))

    population = []
    for _ in range(population_size):
        route = base_ids[:]
        random.shuffle(route)
        population.append(route)

    history = []

    for gen in range(generations):
        scores = [
            fitness(route, matrix, info, urgent_threshold, caution_threshold, objective_mode)
            for route in population
        ]

        elite_idx = sorted(range(len(population)), key=lambda i: scores[i])[:elite_size]
        best = population[elite_idx[0]][:]

        history.append(
            {
                "세대": gen + 1,
                "최적작전거리(km)": route_distance(best, matrix),
                "최적예상소요시간(분)": route_time(best, matrix),
                "최적작전비용": scores[elite_idx[0]],
            }
        )

        next_population = [population[i][:] for i in elite_idx]
        while len(next_population) < population_size:
            parents = random.sample(range(len(population)), k=6)
            parent1 = population[min(parents[:3], key=lambda i: scores[i])]
            parent2 = population[min(parents[3:], key=lambda i: scores[i])]
            child = mutate(ordered_crossover(parent1, parent2), mutation_rate)
            next_population.append(child)

        population = next_population

    final_scores = [
        fitness(route, matrix, info, urgent_threshold, caution_threshold, objective_mode)
        for route in population
    ]
    best_idx = min(range(len(population)), key=lambda i: final_scores[i])
    return population[best_idx], history


# =========================================================
# 경로 대안 / 유조차 배정
# =========================================================
def nearest_neighbor_route(base_ids: List[str], matrix: Dict[str, Dict[str, Dict[str, Any]]], criterion: str = "cost") -> List[str]:
    unvisited = [str(x) for x in base_ids]
    route = []
    current = DEPOT_ID

    while unvisited:
        next_base = min(unvisited, key=lambda base_id: matrix[current][base_id][criterion])
        route.append(next_base)
        unvisited.remove(next_base)
        current = next_base

    return route


def risk_priority_route(targets: pd.DataFrame) -> List[str]:
    return (
        targets.sort_values(["작전위험점수", "잔여가능일수", "보급필요량"], ascending=[False, True, False])["기지ID"]
        .astype(str)
        .tolist()
    )


def split_routes(
    route: List[str],
    info: Dict[str, Dict[str, float]],
    truck_count: int,
    truck_capacity: float,
    assignment_mode: str,
):
    truck_count = max(int(truck_count), 1)
    routes = [[] for _ in range(truck_count)]
    loads = [0.0 for _ in range(truck_count)]

    if assignment_mode == "차량 부하 균등":
        for base_id in route:
            demand = info[base_id]["보급필요량"]
            idx = min(
                range(truck_count),
                key=lambda i: (loads[i] + demand > truck_capacity if truck_capacity > 0 else False, loads[i]),
            )
            routes[idx].append(base_id)
            loads[idx] += demand
        return [r for r in routes if r]

    current_truck = 0
    for base_id in route:
        demand = info[base_id]["보급필요량"]
        if (
            truck_capacity > 0
            and routes[current_truck]
            and loads[current_truck] + demand > truck_capacity
            and current_truck < truck_count - 1
        ):
            current_truck += 1
        routes[current_truck].append(base_id)
        loads[current_truck] += demand

    return [r for r in routes if r]


def route_table(routes, info, matrix, name_map):
    rows = []
    for i, route in enumerate(routes, start=1):
        rows.append(
            {
                "유조차": f"{i}호차",
                "방문기지": " → ".join(name_map.get(x, x) for x in route),
                "보급량(L)": sum(info[x]["보급필요량"] for x in route),
                "작전거리(km)": route_distance(route, matrix),
                "예상소요시간(분)": route_time(route, matrix),
                "평균도로위험도": route_average_risk(route, matrix),
                "통행제한구간": route_blocked_count(route, matrix),
                "작전비용점수": route_operation_cost(route, matrix),
            }
        )
    return pd.DataFrame(rows)


def route_text(route: List[str], name_map: Dict[str, str]) -> str:
    if not route:
        return "방문 기지 없음"
    names = [name_map.get(x, x) for x in route]
    return "중앙보급기지 → " + " → ".join(names) + " → 중앙보급기지"


# =========================================================
# 지도 / 차트 / 테이블
# =========================================================
def marker_color(priority: str) -> str:
    return {"긴급": "red", "주의": "orange", "안정": "blue"}.get(priority, "blue")


def create_map(
    base_df: pd.DataFrame,
    scenario_data: pd.DataFrame,
    routes: List[List[str]],
    best_route: List[str],
    show_route_lines: bool,
):
    route_map = folium.Map(
        location=[base_df["위도"].mean(), base_df["경도"].mean()],
        zoom_start=7,
        tiles="OpenStreetMap",
        control_scale=True,
    )

    base_lookup = base_df.set_index("기지ID").to_dict("index")
    scenario_lookup = scenario_data.set_index("기지ID").to_dict("index")
    order_map = {str(base_id): order for order, base_id in enumerate(best_route, start=1)}
    route_colors = ["blue", "red", "purple", "green", "cadetblue", "darkred", "black"]
    all_points = []

    for _, row in base_df.iterrows():
        base_id = str(row["기지ID"])
        name = row["기지명"]
        lat = float(row["위도"])
        lon = float(row["경도"])
        all_points.append([lat, lon])

        if base_id == DEPOT_ID:
            folium.CircleMarker(
                location=[lat, lon],
                radius=9,
                color="black",
                fill=True,
                fill_color="black",
                fill_opacity=0.95,
                popup=folium.Popup(f"<b>{name}</b><br>구분: 중앙보급기지", max_width=300),
            ).add_to(route_map)

            folium.Marker(
                location=[lat, lon],
                icon=DivIcon(
                    icon_size=(120, 24),
                    icon_anchor=(-10, 28),
                    html="""
                    <div style="font-size:12px;font-weight:700;color:#111;background:rgba(255,255,255,0.88);border:1px solid #444;border-radius:5px;padding:2px 5px;white-space:nowrap;">
                        ★ 중앙보급기지
                    </div>
                    """,
                ),
            ).add_to(route_map)
            continue

        info = scenario_lookup.get(base_id, {})
        priority = info.get("우선순위", "안정")
        order = order_map.get(base_id)
        label = f"[{order}] {name}" if order else name

        popup = f"""
        <b>{label}</b><br>
        기지ID: {base_id}<br>
        우선순위: {priority}<br>
        작전위험점수: {info.get('작전위험점수', 0):.1f}점<br>
        잔여 가능 일수: {info.get('잔여가능일수', 0):.2f}일<br>
        현재 유류량: {info.get('현재유류량', 0):,.0f} L<br>
        일일 소모량: {info.get('일일소모량', 0):,.0f} L<br>
        최대 저장량: {info.get('최대저장량', 0):,.0f} L<br>
        보급 필요량: {info.get('보급필요량', 0):,.0f} L
        """

        folium.CircleMarker(
            location=[lat, lon],
            radius=7,
            color=marker_color(priority),
            fill=True,
            fill_color=marker_color(priority),
            fill_opacity=0.9,
            popup=folium.Popup(popup, max_width=350),
        ).add_to(route_map)

        folium.Marker(
            location=[lat, lon],
            icon=DivIcon(
                icon_size=(130, 24),
                icon_anchor=(-8, 24),
                html=f"""
                <div style="font-size:12px;font-weight:700;color:#111;background:rgba(255,255,255,0.86);border:1px solid #999;border-radius:5px;padding:2px 5px;white-space:nowrap;box-shadow:0 1px 3px rgba(0,0,0,0.18);">
                    {label}
                </div>
                """,
            ),
        ).add_to(route_map)

    if show_route_lines:
        depot = base_lookup[DEPOT_ID]
        for route_idx, route in enumerate(routes):
            if not route:
                continue
            color = route_colors[route_idx % len(route_colors)]
            points = [[depot["위도"], depot["경도"]]]
            points += [[base_lookup[base_id]["위도"], base_lookup[base_id]["경도"]] for base_id in route]
            points += [[depot["위도"], depot["경도"]]]
            folium.PolyLine(points, color=color, weight=3, opacity=0.65, tooltip=f"유조차 {route_idx + 1} 경로").add_to(route_map)

    if all_points:
        route_map.fit_bounds(all_points, padding=(35, 35))

    return route_map


def styled_priority_table(df: pd.DataFrame):
    table = df.copy().sort_values(["작전위험점수", "잔여가능일수", "보급필요량"], ascending=[False, True, False])

    def priority_row_style(row):
        priority = str(row.get("우선순위", ""))
        if priority == "긴급":
            background = "#ffe1e1"
        elif priority == "주의":
            background = "#fff4cc"
        else:
            background = "#eaf7ee"
        return [f"background-color: {background}; color: #111827; border-color: #d1d5db" for _ in row]

    return (
        table.style
        .format(
            {
                "현재유류량": "{:,.0f}",
                "최대저장량": "{:,.0f}",
                "일일소모량": "{:,.0f}",
                "잔여가능일수": "{:.2f}",
                "보급필요량": "{:,.0f}",
                "작전위험점수": "{:.1f}",
            }
        )
        .apply(priority_row_style, axis=1)
        .set_table_styles(
            [
                {"selector": "th", "props": [("background-color", "#f3f4f6"), ("color", "#111827"), ("border-color", "#d1d5db")]},
                {"selector": "td", "props": [("color", "#111827"), ("border-color", "#d1d5db")]},
            ]
        )
    )


def build_chart_data(scenario_data: pd.DataFrame):
    chart_df = scenario_data.copy().sort_values("잔여가능일수", ascending=True)
    chart_df["기지표시"] = chart_df["기지명"] + " (" + chart_df["기지ID"].astype(str) + ")"
    chart_df["기지"] = chart_df["기지ID"].astype(str) + " | " + chart_df["기지명"].astype(str)
    return chart_df


def render_fuel_status_charts(chart_df: pd.DataFrame):
    plot_df = chart_df.copy().sort_values("잔여가능일수", ascending=True)
    base_order = plot_df["기지"].tolist()
    color = alt.Color("기지:N", legend=None, scale=alt.Scale(scheme="tableau20"))
    tooltip = [
        alt.Tooltip("기지:N", title="기지"),
        alt.Tooltip("우선순위:N", title="우선순위"),
        alt.Tooltip("잔여가능일수:Q", title="잔여 가능 일수", format=".2f"),
        alt.Tooltip("보급필요량:Q", title="보급 필요량(L)", format=",.0f"),
        alt.Tooltip("작전위험점수:Q", title="작전위험점수", format=".1f"),
    ]

    chart_col1, chart_col2 = st.columns(2)
    with chart_col1:
        st.markdown("#### 잔여 가능 일수")
        days_chart = (
            alt.Chart(plot_df)
            .mark_bar(cornerRadiusEnd=4)
            .encode(
                x=alt.X("잔여가능일수:Q", title="일"),
                y=alt.Y("기지:N", sort=base_order, title="기지", axis=alt.Axis(labelLimit=180)),
                color=color,
                tooltip=tooltip,
            )
            .properties(height=280)
        )
        st.altair_chart(days_chart, use_container_width=True)

    with chart_col2:
        st.markdown("#### 보급 필요량")
        demand_chart = (
            alt.Chart(plot_df)
            .mark_bar(cornerRadiusEnd=4)
            .encode(
                x=alt.X("보급필요량:Q", title="L"),
                y=alt.Y("기지:N", sort=base_order, title="기지", axis=alt.Axis(labelLimit=180)),
                color=color,
                tooltip=tooltip,
            )
            .properties(height=280)
        )
        st.altair_chart(demand_chart, use_container_width=True)

    st.markdown("#### 작전위험점수 추이")
    risk_df = plot_df.sort_values("작전위험점수", ascending=False).reset_index(drop=True)
    risk_df["순위"] = risk_df.index + 1
    risk_chart = (
        alt.Chart(risk_df)
        .mark_line(point=alt.OverlayMarkDef(size=90), strokeWidth=3)
        .encode(
            x=alt.X("순위:O", title="위험도 순위", axis=alt.Axis(labelAngle=0)),
            y=alt.Y("작전위험점수:Q", title="점수"),
            color=alt.Color("기지:N", legend=alt.Legend(title="기지", orient="bottom", columns=4), scale=alt.Scale(scheme="tableau20")),
            tooltip=tooltip,
        )
        .properties(height=320)
    )
    st.altair_chart(risk_chart, use_container_width=True)


# =========================================================
# 작전 판단 / 시나리오 비교 / 보고서
# =========================================================
def get_transport_defaults(status_df: pd.DataFrame, selected_scenario: str):
    selected = status_df[status_df["시나리오"] == selected_scenario]
    if selected.empty:
        return 1, 25000.0
    return int(selected["가용유조차수"].iloc[0]), float(selected["유조차적재량"].iloc[0])


def total_demand_at_ratio(scenario_data: pd.DataFrame, ratio: float) -> float:
    return float((scenario_data["최대저장량"] * ratio - scenario_data["현재유류량"]).clip(lower=0).sum())


def find_feasible_ratio(scenario_data: pd.DataFrame, total_capacity: float, target_ratio: float) -> float:
    low = 0.0
    high = target_ratio
    for _ in range(40):
        mid = (low + high) / 2
        demand = total_demand_at_ratio(scenario_data, mid)
        if demand <= total_capacity:
            low = mid
        else:
            high = mid
    return low


def make_first_wave_recommendation(targets: pd.DataFrame, total_capacity: float):
    selected_rows = []
    used_capacity = 0.0
    sorted_targets = targets.sort_values(["작전위험점수", "잔여가능일수", "보급필요량"], ascending=[False, True, False])

    for _, row in sorted_targets.iterrows():
        demand = float(row["보급필요량"])
        if used_capacity + demand <= total_capacity:
            selected_rows.append(row)
            used_capacity += demand

    if not selected_rows:
        return pd.DataFrame(), 0.0
    return pd.DataFrame(selected_rows), used_capacity


def assess_operation_status(
    scenario_data: pd.DataFrame,
    targets: pd.DataFrame,
    truck_count: int,
    truck_capacity: float,
    urgent_threshold: float,
):
    total_demand = float(targets["보급필요량"].sum()) if not targets.empty else 0.0
    total_capacity = float(truck_count * truck_capacity)
    urgent_targets = targets[targets["잔여가능일수"] < urgent_threshold]
    urgent_demand = float(urgent_targets["보급필요량"].sum()) if not urgent_targets.empty else 0.0

    if total_demand <= total_capacity:
        status = "수행 가능"
        status_type = "success"
    elif urgent_demand <= total_capacity:
        status = "제한적 가능"
        status_type = "warning"
    else:
        status = "수행 곤란"
        status_type = "error"

    shortage = max(0.0, total_demand - total_capacity)
    required_trucks = math.ceil(total_demand / truck_capacity) if truck_capacity > 0 and total_demand > 0 else 0
    additional_trucks = max(0, required_trucks - truck_count)
    feasible_ratio = find_feasible_ratio(scenario_data, total_capacity, 1.0) if total_capacity > 0 else 0.0

    return {
        "status": status,
        "status_type": status_type,
        "total_demand": total_demand,
        "total_capacity": total_capacity,
        "urgent_demand": urgent_demand,
        "shortage": shortage,
        "required_trucks": required_trucks,
        "additional_trucks": additional_trucks,
        "feasible_ratio": feasible_ratio,
    }


def build_alternative_routes(
    target_ids: List[str],
    best_route: List[str],
    targets: pd.DataFrame,
    matrix: Dict[str, Dict[str, Dict[str, Any]]],
    info: Dict[str, Dict[str, float]],
    truck_count: int,
    truck_capacity: float,
    assignment_mode: str,
):
    alternatives = []
    candidate_routes = [
        ("AI 종합 추천안", "작전비용, 거리, 긴급도를 함께 고려한 유전 알고리즘 결과입니다.", best_route),
        ("긴급기지 우선안", "작전위험점수와 잔여가능일수를 기준으로 긴급 기지를 먼저 방문합니다.", risk_priority_route(targets)),
        ("작전비용 최저 탐욕안", "현재 지점에서 작전비용이 가장 낮은 구간을 순차 선택합니다.", nearest_neighbor_route(target_ids, matrix, criterion="cost")),
    ]

    for name, purpose, route in candidate_routes:
        routes = split_routes(route, info, truck_count, truck_capacity, assignment_mode)
        alternatives.append(
            {
                "대안": name,
                "목적": purpose,
                "방문순서": route,
                "유조차별경로": routes,
                "총작전거리(km)": total_distance(routes, matrix),
                "총예상소요시간(분)": total_time(routes, matrix),
                "총작전비용점수": total_operation_cost(routes, matrix),
                "통행제한구간": total_blocked_count(routes, matrix),
                "총보급량(L)": sum(info[x]["보급필요량"] for x in route),
            }
        )
    return alternatives


def alternative_summary_table(alternatives: List[Dict], name_map: Dict[str, str]):
    rows = []
    for alt in alternatives:
        rows.append(
            {
                "대안": alt["대안"],
                "목적": alt["목적"],
                "총작전거리(km)": alt["총작전거리(km)"],
                "총예상소요시간(분)": alt["총예상소요시간(분)"],
                "총작전비용점수": alt["총작전비용점수"],
                "통행제한구간": alt["통행제한구간"],
                "총보급량(L)": alt["총보급량(L)"],
                "방문순서": route_text(alt["방문순서"], name_map),
            }
        )
    return pd.DataFrame(rows)


def build_scenario_comparison(
    base_df: pd.DataFrame,
    status_df: pd.DataFrame,
    scenario_list: List[str],
    target_ratio: float,
    urgent_threshold: float,
    caution_threshold: float,
    truck_count: int,
    truck_capacity: float,
):
    rows = []
    for scenario in scenario_list:
        data, targets = prepare_scenario(base_df, status_df, scenario, target_ratio, urgent_threshold, caution_threshold)
        total_demand = float(targets["보급필요량"].sum()) if not targets.empty else 0.0
        total_capacity = truck_count * truck_capacity
        required_trucks = math.ceil(total_demand / truck_capacity) if truck_capacity > 0 and total_demand > 0 else 0
        rows.append(
            {
                "시나리오": scenario,
                "보급대상기지수": len(targets),
                "긴급기지수": int((data["우선순위"] == "긴급").sum()),
                "총보급필요량(L)": total_demand,
                "현재수송가능량(L)": total_capacity,
                "필요유조차수": required_trucks,
                "부족유조차수": max(0, required_trucks - truck_count),
            }
        )
    return pd.DataFrame(rows)


def make_final_recommendation_text(
    selected_scenario: str,
    operation_status: Dict,
    targets: pd.DataFrame,
    first_wave_df: pd.DataFrame,
    first_wave_load: float,
    target_ratio: float,
    truck_count: int,
    truck_capacity: float,
    matrix_source: str,
    optimized_distance: float,
    optimized_time: float,
    optimized_cost: float,
):
    top_targets = targets.sort_values(["작전위험점수", "잔여가능일수", "보급필요량"], ascending=[False, True, False]).head(3)
    top_names = ", ".join(top_targets["기지명"].astype(str).tolist()) if not top_targets.empty else "없음"
    first_wave_names = ", ".join(first_wave_df["기지명"].astype(str).tolist()) if not first_wave_df.empty else "없음"

    if operation_status["status"] == "수행 가능":
        decision = "현재 수송자산으로 목표 보급률 기준의 전체 보급 수행이 가능합니다. AI 종합 추천안을 기준으로 수송 경로를 편성하는 것이 적절합니다."
    elif operation_status["status"] == "제한적 가능":
        decision = "현재 수송자산으로 전체 보급은 제한되지만, 긴급 기지 중심의 1차 보급은 가능합니다. 작전위험점수가 높은 기지를 우선 보급하고, 추가 수송자산 투입을 검토해야 합니다."
    else:
        decision = "현재 수송자산만으로는 긴급 기지 보급도 제한될 수 있습니다. 유조차 추가 투입, 목표 보급률 하향, 또는 긴급 기지 한정 보급으로 계획을 조정해야 합니다."

    text = f"""
[POL 수송 최적화 작전 판단 보고]

1. 선택 시나리오
- 시나리오: {selected_scenario}
- 거리 기준: {matrix_source}
- 목표 보급률: {target_ratio * 100:.0f}%
- 가용 유조차 수: {truck_count}대
- 유조차 1대 적재량: {truck_capacity:,.0f} L

2. 작전 가능 여부
- 판정: {operation_status['status']}
- 총 보급 필요량: {operation_status['total_demand']:,.0f} L
- 총 수송 가능량: {operation_status['total_capacity']:,.0f} L
- 부족 유류량: {operation_status['shortage']:,.0f} L
- 필요 유조차 수: {operation_status['required_trucks']}대
- 추가 필요 유조차 수: {operation_status['additional_trucks']}대

3. 최적 경로 성과
- 최적 경로 작전거리: {optimized_distance:,.1f} km
- 최적 경로 예상 소요시간: {optimized_time:,.0f}분
- 최적 경로 작전비용점수: {optimized_cost:,.1f}점

4. 우선 보급 판단
- 최우선 검토 기지: {top_names}
- 1차 수송 권고 기지: {first_wave_names}
- 1차 수송 예상 보급량: {first_wave_load:,.0f} L

5. 권고 조치
- {decision}
"""
    if operation_status["status"] != "수행 가능":
        text += f"\n- 현재 수송 가능량 기준으로 달성 가능한 추정 보급률은 최대 약 {operation_status['feasible_ratio'] * 100:.1f}%입니다."
    return text.strip()


# =========================================================
# Streamlit 화면
# =========================================================
base_df, status_df, scenario_df = load_data(file_signature())
op_distance_df, op_settings, op_file_path = load_operational_distance_info(file_signature())

if "final_result" not in st.session_state:
    st.session_state.final_result = None
if "last_setting" not in st.session_state:
    st.session_state.last_setting = None

st.title("⛽ 항공유류(POL) 수송 최적화 의사결정 지원 시스템")
st.caption("기지별 유류 상태, 수송자산 제약, 작전거리정보, 경로 최적화를 통합한 군수 의사결정 지원 화면입니다.")

scenario_list = scenario_df["시나리오"].dropna().unique().tolist() or status_df["시나리오"].dropna().unique().tolist()

with st.sidebar:
    st.header("🪖 작전 조건 입력")

    selected_scenario = st.selectbox("시나리오", scenario_list)
    default_truck_count, default_truck_capacity = get_transport_defaults(status_df, selected_scenario)

    st.caption("핵심 조건만 조정한 뒤 바로 계산할 수 있습니다.")
    target_ratio = st.slider("목표 보급률", 0.50, 1.00, 0.80, 0.05)
    truck_count_input = st.number_input(
        "가용 유조차 수",
        min_value=1,
        max_value=30,
        value=default_truck_count,
        step=1,
        key=f"truck_count_{selected_scenario}",
    )
    truck_capacity_input = st.number_input(
        "1대 적재량(L)",
        min_value=1000,
        max_value=200000,
        value=int(default_truck_capacity),
        step=1000,
        key=f"truck_capacity_{selected_scenario}",
    )

    objective_mode = st.selectbox("최적화 목적", ["종합 판단", "거리 최소화", "긴급기지 우선", "차량 부하 균등"])
    assignment_mode = "순차 적재"
    show_route_lines = True
    urgent_threshold = 3.0
    caution_threshold = 5.0
    population_size = 80
    generations = 200
    mutation_rate = 0.10

    with st.expander("고급 설정"):
        assignment_mode = st.selectbox("유조차 배정 방식", ["순차 적재", "차량 부하 균등"])
        show_route_lines = st.checkbox("지도에 경로선 표시", value=True)
        urgent_threshold = st.number_input("긴급 기준 일수", min_value=1.0, max_value=10.0, value=3.0, step=0.5)
        caution_threshold = st.number_input("주의 기준 일수", min_value=2.0, max_value=15.0, value=5.0, step=0.5)
        st.markdown("##### 유전 알고리즘")
        population_size = st.number_input("개체 수", 30, 300, 80, 10)
        generations = st.number_input("세대 수", 50, 700, 200, 50)
        mutation_rate = st.slider("돌연변이율", 0.01, 0.30, 0.10, 0.01)

    run_button = st.button("🚀 최적해 계산하기", type="primary", use_container_width=True)

    if st.button("🔄 캐시 초기화", use_container_width=True):
        st.cache_data.clear()
        st.session_state.clear()
        st.rerun()

if caution_threshold <= urgent_threshold:
    st.error("주의 기준 일수는 긴급 기준 일수보다 커야 합니다.")
    st.stop()

current_setting = {
    "시나리오": selected_scenario,
    "목표보급률": float(target_ratio),
    "가용유조차수": int(truck_count_input),
    "유조차적재량": float(truck_capacity_input),
    "긴급기준": float(urgent_threshold),
    "주의기준": float(caution_threshold),
    "최적화목적": objective_mode,
    "배정방식": assignment_mode,
    "개체수": int(population_size),
    "세대수": int(generations),
    "돌연변이율": float(mutation_rate),
    "경로선표시": bool(show_route_lines),
    "작전거리파일": os.path.basename(op_file_path) if op_file_path else "없음",
}

if st.session_state.last_setting != current_setting:
    st.session_state.final_result = None
    st.session_state.last_setting = current_setting

scenario_data, targets = prepare_scenario(base_df, status_df, selected_scenario, target_ratio, urgent_threshold, caution_threshold)
truck_count = int(truck_count_input)
truck_capacity = float(truck_capacity_input)

base_records = tuple((str(row["기지ID"]), float(row["위도"]), float(row["경도"])) for _, row in base_df.iterrows())
operational_records = tuple(
    (
        str(row["출발기지ID"]),
        str(row["도착기지ID"]),
        float(row["도로거리km"]),
        float(row["예상소요시간분"]),
        str(row["통행가능여부"]),
        float(row["도로위험도"]),
        float(row["작전비용점수"]),
    )
    for _, row in op_distance_df.iterrows()
) if not op_distance_df.empty else tuple()

metric_matrix, matrix_source = build_metric_matrix(base_records, operational_records, op_settings)
name_map = base_df.set_index("기지ID")["기지명"].to_dict()

operation_status_preview = assess_operation_status(scenario_data, targets, truck_count, truck_capacity, urgent_threshold)

st.subheader(f"선택 시나리오: {selected_scenario}")
scenario_row = scenario_df[scenario_df["시나리오"] == selected_scenario]
if not scenario_row.empty:
    row = scenario_row.iloc[0]
    if "시나리오유형" in row.index:
        st.write(f"**시나리오 유형:** {row['시나리오유형']}")
    if "상황설명" in row.index:
        st.write(f"**상황 설명:** {row['상황설명']}")
    if "분석목표" in row.index:
        st.write(f"**분석 목표:** {row['분석목표']}")

if matrix_source == "작전거리정보 파일 사용":
    st.success(f"거리 기준: {matrix_source} ({os.path.basename(op_file_path)})")
else:
    st.warning("거리 기준: 직선거리 기반 추정값 사용입니다. operational_distance_info.xlsx 파일을 같은 폴더에 넣으면 작전거리정보를 사용합니다.")

metric_col1, metric_col2, metric_col3, metric_col4, metric_col5 = st.columns(5)
metric_col1.metric("보급 대상 기지 수", f"{len(targets)}개")
metric_col2.metric("긴급 기지 수", f"{(scenario_data['우선순위'] == '긴급').sum()}개")
metric_col3.metric("가용 유조차 수", f"{truck_count}대")
metric_col4.metric("총 수송 가능량", f"{operation_status_preview['total_capacity']:,.0f} L")
metric_col5.metric("부족 유류량", f"{operation_status_preview['shortage']:,.0f} L")

if operation_status_preview["status_type"] == "success":
    st.success(f"작전 가능 여부 사전판정: {operation_status_preview['status']}입니다.")
elif operation_status_preview["status_type"] == "warning":
    st.warning(f"작전 가능 여부 사전판정: {operation_status_preview['status']}입니다. 긴급 기지 중심 1차 보급을 우선 검토해야 합니다.")
else:
    st.error(f"작전 가능 여부 사전판정: {operation_status_preview['status']}입니다. 추가 수송자산 또는 목표 보급률 조정이 필요합니다.")

if run_button:
    if targets.empty:
        st.session_state.final_result = "empty"
    else:
        info = status_dict(targets)
        target_ids = (
            targets.sort_values(["작전위험점수", "잔여가능일수", "보급필요량"], ascending=[False, True, False])["기지ID"]
            .astype(str)
            .tolist()
        )

        best_route, history = genetic_algorithm(
            target_ids,
            metric_matrix,
            info,
            int(population_size),
            int(generations),
            float(mutation_rate),
            urgent_threshold,
            caution_threshold,
            objective_mode,
        )

        existing_route = target_ids[:]
        existing_routes = split_routes(existing_route, info, truck_count, truck_capacity, assignment_mode)
        optimized_routes = split_routes(best_route, info, truck_count, truck_capacity, assignment_mode)

        existing_distance = total_distance(existing_routes, metric_matrix)
        optimized_distance = total_distance(optimized_routes, metric_matrix)
        existing_time = total_time(existing_routes, metric_matrix)
        optimized_time = total_time(optimized_routes, metric_matrix)
        existing_cost = total_operation_cost(existing_routes, metric_matrix)
        optimized_cost = total_operation_cost(optimized_routes, metric_matrix)
        reduction_rate = (existing_cost - optimized_cost) / existing_cost * 100 if existing_cost else 0.0

        operation_status = assess_operation_status(scenario_data, targets, truck_count, truck_capacity, urgent_threshold)
        first_wave_df, first_wave_load = make_first_wave_recommendation(targets, operation_status["total_capacity"])

        alternatives = build_alternative_routes(
            target_ids,
            best_route,
            targets,
            metric_matrix,
            info,
            truck_count,
            truck_capacity,
            assignment_mode,
        )

        scenario_comparison_df = build_scenario_comparison(
            base_df,
            status_df,
            scenario_list,
            target_ratio,
            urgent_threshold,
            caution_threshold,
            truck_count,
            truck_capacity,
        )

        final_report = make_final_recommendation_text(
            selected_scenario,
            operation_status,
            targets,
            first_wave_df,
            first_wave_load,
            target_ratio,
            truck_count,
            truck_capacity,
            matrix_source,
            optimized_distance,
            optimized_time,
            optimized_cost,
        )

        st.session_state.final_result = {
            "best_route": best_route,
            "existing_route": existing_route,
            "existing_routes": existing_routes,
            "optimized_routes": optimized_routes,
            "existing_distance": existing_distance,
            "optimized_distance": optimized_distance,
            "existing_time": existing_time,
            "optimized_time": optimized_time,
            "existing_cost": existing_cost,
            "optimized_cost": optimized_cost,
            "reduction_rate": reduction_rate,
            "table": route_table(optimized_routes, info, metric_matrix, name_map),
            "total_demand": sum(info[x]["보급필요량"] for x in best_route),
            "total_capacity": truck_count * truck_capacity,
            "history": history,
            "scenario_data": scenario_data,
            "targets": targets,
            "operation_status": operation_status,
            "first_wave_df": first_wave_df,
            "first_wave_load": first_wave_load,
            "alternatives": alternatives,
            "scenario_comparison_df": scenario_comparison_df,
            "final_report": final_report,
            "show_route_lines": show_route_lines,
            "matrix_source": matrix_source,
        }

result = st.session_state.final_result

if result == "empty":
    st.success("현재 설정에서는 보급이 필요한 기지가 없습니다.")

elif isinstance(result, dict):
    st.subheader("📌 최적화 결과 상황판")
    operation_status = result["operation_status"]

    result_col1, result_col2, result_col3, result_col4, result_col5 = st.columns(5)
    result_col1.metric("작전 판정", operation_status["status"])
    result_col2.metric("기존 작전거리", f"{result['existing_distance']:,.1f} km")
    result_col3.metric("최적 작전거리", f"{result['optimized_distance']:,.1f} km")
    result_col4.metric("작전비용 감소율", f"{result['reduction_rate']:.1f}%")
    result_col5.metric("추가 필요 유조차", f"{operation_status['additional_trucks']}대")

    st.caption(f"거리 기준: {result['matrix_source']} / 최적 예상 소요시간: {result['optimized_time']:,.0f}분 / 최적 작전비용점수: {result['optimized_cost']:,.1f}점")

    if operation_status["status_type"] == "success":
        st.success("✅ 현재 수송자산으로 목표 보급률 기준의 전체 보급 수행이 가능합니다.")
    elif operation_status["status_type"] == "warning":
        st.warning("⚠️ 현재 수송자산으로 전체 보급은 제한되지만, 긴급 기지 중심의 1차 보급은 가능합니다.")
    else:
        st.error("🚨 현재 수송자산만으로는 긴급 기지 보급도 제한될 수 있습니다. 추가 유조차 투입 또는 목표 보급률 하향이 필요합니다.")

    if "통행제한구간" in result["table"].columns and result["table"]["통행제한구간"].sum() > 0:
        st.warning("일부 경로에 통행 제한 구간이 포함되어 있습니다. 작전거리정보 파일의 통행가능여부를 확인해야 합니다.")

    if truck_capacity > 0 and not result["table"].empty and (result["table"]["보급량(L)"] > truck_capacity).any():
        st.warning("일부 유조차 배정량이 차량별 적재량을 초과합니다. 차량 수를 늘리거나 목표 보급률을 낮춰야 합니다.")

    st.divider()
    result_menu = st.radio(
        "결과 카테고리",
        [
            "작전상황판",
            "수송경로지도",
            "유류상태분석",
            "차량배정결과",
            "경로대안비교",
            "상세데이터",
            "알고리즘검증",
            "최종보고서",
        ],
        horizontal=True,
        key="result_menu_v8",
    )
    st.divider()

    if result_menu == "작전상황판":
        st.markdown("### 🧭 작전 판단 요약")
        summary_col1, summary_col2, summary_col3 = st.columns(3)
        with summary_col1:
            st.info(f"""**총 보급 필요량**\n\n{operation_status['total_demand']:,.0f} L""")
        with summary_col2:
            st.info(f"""**총 수송 가능량**\n\n{operation_status['total_capacity']:,.0f} L""")
        with summary_col3:
            st.info(f"""**부족 유류량**\n\n{operation_status['shortage']:,.0f} L""")

        st.markdown("### 우선 보급 권고")
        first_wave_df = result["first_wave_df"]
        if not first_wave_df.empty:
            show_cols = ["기지ID", "기지명", "잔여가능일수", "우선순위", "작전위험점수", "보급필요량"]
            st.dataframe(
                first_wave_df[show_cols]
                .sort_values(["작전위험점수", "잔여가능일수"], ascending=[False, True])
                .style.format({"잔여가능일수": "{:.2f}", "작전위험점수": "{:.1f}", "보급필요량": "{:,.0f}"}),
                use_container_width=True,
            )
            st.success(f"1차 수송 권고 보급량은 {result['first_wave_load']:,.0f} L입니다.")
        else:
            st.warning("현재 수송 가능량 안에서 완전 보급 가능한 기지가 없습니다. 목표 보급률 조정 또는 부분 보급 기준이 필요합니다.")

        st.markdown("### 핵심 해석")
        st.write(
            f"""
            선택된 **{selected_scenario}** 시나리오에서 기존 경로의 작전거리는 
            **{result['existing_distance']:,.1f} km**이고, 최적화 후 작전거리는 
            **{result['optimized_distance']:,.1f} km**입니다.

            작전비용점수는 기존 **{result['existing_cost']:,.1f}점**에서 최적화 후 
            **{result['optimized_cost']:,.1f}점**으로 바뀌었고, 감소율은 **{result['reduction_rate']:.1f}%**입니다.
            """
        )

    elif result_menu == "수송경로지도":
        st.markdown("### 🗺️ 최적 수송 경로 지도")
        st.caption("지도에는 번호와 기지명만 상시 표시하고, 우선순위와 유류 수치는 마커 클릭 시 확인하도록 단순화했습니다.")
        st_folium(
            create_map(base_df, result["scenario_data"], result["optimized_routes"], result["best_route"], result["show_route_lines"]),
            height=680,
            use_container_width=True,
            key="optimized_route_map_v8",
        )

    elif result_menu == "유류상태분석":
        st.markdown("### 📊 유류 상태 분석")
        chart_df = build_chart_data(result["scenario_data"])
        render_fuel_status_charts(chart_df)

    elif result_menu == "차량배정결과":
        st.markdown("### 🚚 유조차별 최적 배정 결과")
        st.dataframe(
            result["table"].style.format(
                {
                    "보급량(L)": "{:,.0f}",
                    "작전거리(km)": "{:,.1f}",
                    "예상소요시간(분)": "{:,.0f}",
                    "평균도로위험도": "{:.1f}",
                    "작전비용점수": "{:,.1f}",
                }
            ),
            use_container_width=True,
        )

        download_df = result["table"].copy()
        download_df.insert(0, "시나리오", selected_scenario)
        download_df["기존작전거리(km)"] = result["existing_distance"]
        download_df["최적작전거리(km)"] = result["optimized_distance"]
        download_df["기존작전비용점수"] = result["existing_cost"]
        download_df["최적작전비용점수"] = result["optimized_cost"]
        download_df["작전비용감소율(%)"] = result["reduction_rate"]
        st.download_button(
            "최적화 결과 CSV 다운로드",
            data=download_df.to_csv(index=False, encoding="utf-8-sig"),
            file_name=f"{selected_scenario}_POL_최적화결과.csv",
            mime="text/csv",
        )

    elif result_menu == "경로대안비교":
        st.markdown("### 🔀 경로 대안 비교")
        alt_df = alternative_summary_table(result["alternatives"], name_map)
        st.dataframe(
            alt_df.style.format(
                {
                    "총작전거리(km)": "{:,.1f}",
                    "총예상소요시간(분)": "{:,.0f}",
                    "총작전비용점수": "{:,.1f}",
                    "총보급량(L)": "{:,.0f}",
                }
            ),
            use_container_width=True,
        )
        st.info("경로 대안은 작전비용, 긴급도, 수송거리 기준을 비교하기 위한 참고안입니다.")

    elif result_menu == "상세데이터":
        st.markdown("### 📋 전체 기지 상세 데이터")
        show_cols = ["기지ID", "기지명", "현재유류량", "최대저장량", "일일소모량", "잔여가능일수", "우선순위", "작전위험점수", "보급필요량"]
        st.dataframe(styled_priority_table(result["scenario_data"][show_cols]), use_container_width=True)

    elif result_menu == "알고리즘검증":
        st.markdown("### 🧬 유전 알고리즘 수렴 과정")
        if result["history"]:
            history_df = pd.DataFrame(result["history"])
            st.line_chart(history_df.set_index("세대")[["최적작전비용"]], use_container_width=True)
            st.dataframe(history_df, use_container_width=True)
        else:
            st.info("알고리즘 수렴 기록이 없습니다.")

    elif result_menu == "최종보고서":
        st.markdown("### 📝 최종 권고문")
        st.text_area("작전 판단 보고서", result["final_report"], height=440)
        st.download_button(
            "작전 판단 보고서 TXT 다운로드",
            data=result["final_report"].encode("utf-8-sig"),
            file_name=f"{selected_scenario}_POL_작전판단보고서.txt",
            mime="text/plain",
        )

else:
    st.info("왼쪽 사이드바에서 작전 조건을 입력한 뒤, **최적해 계산하기** 버튼을 누르시면 됩니다.")
    st.markdown("### 실행 전 현황 미리보기")

    preview_menu = st.radio("미리보기 카테고리", ["유류상태분석", "시나리오비교", "상세데이터"], horizontal=True, key="preview_menu_v8")

    if preview_menu == "유류상태분석":
        chart_df = build_chart_data(scenario_data)
        render_fuel_status_charts(chart_df)

    elif preview_menu == "시나리오비교":
        scenario_comparison_df = build_scenario_comparison(base_df, status_df, scenario_list, target_ratio, urgent_threshold, caution_threshold, truck_count, truck_capacity)
        st.dataframe(
            scenario_comparison_df.style.format({"총보급필요량(L)": "{:,.0f}", "현재수송가능량(L)": "{:,.0f}"}),
            use_container_width=True,
        )

    elif preview_menu == "상세데이터":
        show_cols = ["기지ID", "기지명", "현재유류량", "최대저장량", "일일소모량", "잔여가능일수", "우선순위", "작전위험점수", "보급필요량"]
        st.dataframe(styled_priority_table(scenario_data[show_cols]), use_container_width=True)
