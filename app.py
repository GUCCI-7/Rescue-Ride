import sqlite3
from datetime import datetime
from math import radians, sin, cos, sqrt, atan2
from pathlib import Path
from io import BytesIO
import re

import folium
import pandas as pd
import requests
import streamlit as st
import qrcode
import streamlit.components.v1 as components
from streamlit_folium import st_folium

try:
    from streamlit_geolocation import streamlit_geolocation
    GPS_AVAILABLE = True
except ImportError:
    GPS_AVAILABLE = False


# =========================================================
# 1. 基本設定
# =========================================================

st.set_page_config(
    page_title="レスキューライド",
    page_icon="🚲",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DB_FILE = "rescue_ride.db"

DEFAULT_LATITUDE = 35.681236
DEFAULT_LONGITUDE = 139.767125

ACCIDENT_FILE_CANDIDATES = [
    Path("honhyo_2024.xlsx"),
    Path("honhyo_2024.xls"),
    Path("honhyo_2024.csv"),
]

MAX_ACCIDENT_MAP_POINTS = 1200


# =========================================================
# 2. デザイン
# =========================================================

st.markdown(
    """
    <style>
    .block-container {
        padding-top: 1.1rem;
        padding-bottom: 3rem;
        max-width: 1450px;
    }

    [data-testid="stSidebar"] {
        display: none;
    }

    .app-title {
        font-size: 44px;
        font-weight: 900;
        color: #1d3557;
        margin-bottom: 0;
    }

    .app-subtitle {
        font-size: 17px;
        color: #666;
        margin-top: 0;
        margin-bottom: 16px;
    }

    .mode-select-wrap {
        min-height: 78vh;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    .mode-select-card {
        width: 100%;
        max-width: 1100px;
        background: linear-gradient(180deg, #ffffff, #f4fbff);
        border: 1px solid #d7e8ff;
        border-radius: 30px;
        padding: 44px 40px;
        box-shadow: 0 20px 50px rgba(20, 80, 140, 0.16);
        text-align: center;
    }

    .mode-select-title {
        font-size: 48px;
        font-weight: 950;
        color: #1d3557;
        margin-bottom: 8px;
    }

    .mode-select-subtitle {
        font-size: 20px;
        color: #555;
        margin-bottom: 28px;
    }

    .mode-box {
        background: #f7f9fb;
        border: 1px solid #e1e5e9;
        border-radius: 22px;
        padding: 24px;
        min-height: 185px;
        margin-bottom: 14px;
    }

    .mode-box-title {
        font-size: 30px;
        font-weight: 900;
        margin-bottom: 10px;
    }

    .mode-box-text {
        font-size: 17px;
        color: #555;
        line-height: 1.75;
    }

    .top-mode-bar {
        background: #f7fbff;
        border: 1px solid #d7e8ff;
        border-radius: 16px;
        padding: 16px 20px;
        margin: 10px 0 18px 0;
    }

    .info-card {
        background-color: #f7f9fb;
        border: 1px solid #e1e5e9;
        border-radius: 15px;
        padding: 20px;
        margin-bottom: 14px;
    }

    .danger-card {
        background-color: #fff3f3;
        border-left: 6px solid #e63946;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 12px;
    }

    .warning-card {
        background-color: #fff9e8;
        border-left: 6px solid #f4a261;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 12px;
    }

    .safe-card {
        background-color: #effaf3;
        border-left: 6px solid #2a9d8f;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 12px;
    }

    .bike-hero {
        background: linear-gradient(90deg, #0077b6, #00b4d8);
        color: white;
        border-radius: 24px;
        padding: 28px 26px;
        margin: 16px 0 22px 0;
        box-shadow: 0 14px 35px rgba(0, 119, 182, 0.25);
    }

    .bike-hero-title {
        font-size: 32px;
        font-weight: 900;
    }

    .bike-hero-text {
        font-size: 18px;
        margin-top: 8px;
        line-height: 1.7;
    }

    .car-hero {
        border-radius: 28px;
        padding: 38px 30px;
        margin: 16px 0 26px 0;
        color: white;
        text-align: center;
        font-weight: 900;
        box-shadow: 0 18px 45px rgba(0, 0, 0, 0.16);
    }

    .car-danger {
        background: linear-gradient(90deg, #7f0000, #ff1f1f, #7f0000);
        animation: redPulse 0.8s infinite alternate;
    }

    .car-warning {
        background: linear-gradient(90deg, #d65a00, #ff9f1c, #d65a00);
    }

    .car-safe {
        background: linear-gradient(90deg, #087f5b, #20b486, #087f5b);
    }

    .car-title {
        font-size: 40px;
        line-height: 1.35;
    }

    .car-message {
        font-size: 24px;
        margin-top: 12px;
        line-height: 1.55;
    }

    .car-card {
        background: linear-gradient(180deg, #ffffff, #f8fbff);
        border: 1px solid #dfe9f5;
        border-radius: 18px;
        padding: 20px;
        margin: 10px 0;
        min-height: 135px;
    }

    .car-card-title {
        font-size: 22px;
        font-weight: 900;
        color: #1d3557;
        margin-bottom: 8px;
    }

    .car-card-text {
        font-size: 16px;
        color: #555;
        line-height: 1.7;
    }

    .car-emergency {
        background: #fff3f3;
        border: 2px solid #ffb3b3;
        border-radius: 18px;
        padding: 18px 22px;
        margin: 14px 0;
        color: #9b0000;
        font-weight: 850;
    }

    @keyframes redPulse {
        from {
            filter: brightness(1);
            transform: scale(1.000);
        }
        to {
            filter: brightness(1.22);
            transform: scale(1.012);
        }
    }

    .stButton > button {
        border-radius: 12px;
        font-weight: 750;
        min-height: 46px;
    }

    @media screen and (max-width: 700px) {
        .mode-select-title {
            font-size: 34px;
        }
        .mode-select-subtitle {
            font-size: 16px;
        }
        .app-title {
            font-size: 32px;
        }
        .car-title {
            font-size: 28px;
        }
        .car-message {
            font-size: 18px;
        }
        .bike-hero-title {
            font-size: 24px;
        }
    }

    .safety-aura-fixed-panel {
        border-radius: 28px;
        padding: 32px 26px;
        margin: 16px 0 26px 0;
        color: white;
        text-align: center;
        font-weight: 900;
        box-shadow: 0 16px 38px rgba(0,0,0,0.18);
    }

    .safety-aura-fixed-waiting {
        background: linear-gradient(90deg, #457b9d, #00b4d8, #457b9d);
    }

    .safety-aura-fixed-safe {
        background: linear-gradient(90deg, #087f5b, #20b486, #087f5b);
    }

    .safety-aura-fixed-warning {
        background: linear-gradient(90deg, #e66b00, #ff9f1c, #e66b00);
    }

    .safety-aura-fixed-danger {
        background: linear-gradient(90deg, #b00020, #ff2d2d, #b00020);
        animation: redPulse 1s infinite alternate;
    }

    .safety-aura-fixed-title {
        font-size: 34px;
        line-height: 1.35;
    }

    .safety-aura-fixed-text {
        font-size: 21px;
        margin-top: 10px;
        line-height: 1.65;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 3. セッション状態
# =========================================================

def initialize_session_state():
    defaults = {
        "mode_selected": False,
        "app_mode": None,
        "current_latitude": None,
        "current_longitude": None,
        "start_location": None,
        "goal_location": None,
        "safety_aura_enabled": True,
        "safety_aura_radius": 200,
        "safety_aura_min_level": 3,
        "car_voice_enabled": True,
        "car_red_flash_enabled": True,
        "car_danger_radius": 250,
        "car_warning_radius": 500,
        "car_bicycle_alert_radius": 180,
        "last_voice_key": "",
        "last_voice_time": None,
        "share_app_url": "",
        "bike_aura_voice_enabled": True,
        "last_bike_aura_voice_key": "",
        "last_bike_aura_voice_time": None,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


initialize_session_state()


# =========================================================
# 4. 共通関数
# =========================================================

def haversine_distance(lat1, lon1, lat2, lon2):
    earth_radius = 6371000

    lat1_radian = radians(lat1)
    lat2_radian = radians(lat2)
    lat_difference = radians(lat2 - lat1)
    lon_difference = radians(lon2 - lon1)

    a = (
        sin(lat_difference / 2) ** 2
        + cos(lat1_radian)
        * cos(lat2_radian)
        * sin(lon_difference / 2) ** 2
    )

    c = 2 * atan2(sqrt(a), sqrt(1 - a))

    return earth_radius * c


def danger_color(level):
    if level >= 5:
        return "darkred"
    if level == 4:
        return "red"
    if level == 3:
        return "orange"
    if level == 2:
        return "cadetblue"
    return "green"


def danger_text(level):
    if level >= 5:
        return "非常に危険"
    if level == 4:
        return "危険"
    if level == 3:
        return "注意"
    if level == 2:
        return "やや注意"
    return "比較的安全"


def get_current_location():
    if (
        st.session_state.current_latitude is None
        or st.session_state.current_longitude is None
    ):
        return None

    return {
        "latitude": st.session_state.current_latitude,
        "longitude": st.session_state.current_longitude,
    }


# =========================================================
# 5. 公式事故データ
# =========================================================

def normalize_column_name(value):
    return (
        str(value)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("　", "")
        .replace("_", "")
        .replace("-", "")
        .replace("（", "(")
        .replace("）", ")")
    )


def find_first_column(columns, exact_names, partial_words):
    normalized_map = {
        normalize_column_name(col): col
        for col in columns
    }

    for name in exact_names:
        normalized_name = normalize_column_name(name)
        if normalized_name in normalized_map:
            return normalized_map[normalized_name]

    for col in columns:
        normalized_col = normalize_column_name(col)
        if any(normalize_column_name(word) in normalized_col for word in partial_words):
            return col

    return None


def decimal_from_dms(degree, minute=0, second=0):
    try:
        degree_value = float(degree)
        minute_value = float(minute or 0)
        second_value = float(second or 0)
    except (TypeError, ValueError):
        return None

    if not (0 <= minute_value < 60 and 0 <= second_value < 60):
        return None

    sign = -1 if degree_value < 0 else 1
    degree_value = abs(degree_value)

    return sign * (degree_value + minute_value / 60 + second_value / 3600)


def coordinate_is_valid(value, coordinate_type):
    if value is None:
        return False

    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return False

    if coordinate_type == "lat":
        return 20 <= numeric <= 50

    return 120 <= numeric <= 155


def packed_dms_to_decimal(number, coordinate_type):
    try:
        value = abs(float(number))
    except (TypeError, ValueError):
        return None

    degree_digits = 2 if coordinate_type == "lat" else 3
    raw = f"{value:.8f}".rstrip("0").rstrip(".")
    integer_part, _, decimal_part = raw.partition(".")

    if len(integer_part) < degree_digits + 4:
        return None

    degrees_text = integer_part[:degree_digits]
    minutes_text = integer_part[degree_digits:degree_digits + 2]
    seconds_text = integer_part[degree_digits + 2:]

    if decimal_part:
        seconds_text = f"{seconds_text}.{decimal_part}"

    try:
        degrees = float(degrees_text)
        minutes = float(minutes_text)
        seconds = float(seconds_text)
    except ValueError:
        return None

    decimal = decimal_from_dms(degrees, minutes, seconds)

    if float(number) < 0 and decimal is not None:
        decimal = -decimal

    return decimal


def choose_coordinate_candidate(number, coordinate_type):
    try:
        numeric = float(number)
    except (TypeError, ValueError):
        return None

    candidates = [numeric]

    for divisor in [10, 100, 1000, 10000, 100000, 1000000, 10000000, 100000000]:
        candidates.append(numeric / divisor)

    packed_candidate = packed_dms_to_decimal(numeric, coordinate_type)

    if packed_candidate is not None:
        candidates.append(packed_candidate)

    for candidate in candidates:
        if coordinate_is_valid(candidate, coordinate_type):
            return float(candidate)

    return None


def parse_coordinate_value(value, coordinate_type):
    if pd.isna(value):
        return None

    try:
        return choose_coordinate_candidate(float(value), coordinate_type)
    except (TypeError, ValueError):
        pass

    text_value = str(value).strip()

    if not text_value:
        return None

    negative = "南緯" in text_value or "西経" in text_value or text_value.startswith("-")

    cleaned = (
        text_value
        .replace(",", "")
        .replace("北緯", "")
        .replace("東経", "")
        .replace("南緯", "")
        .replace("西経", "")
        .strip()
    )

    dms_pattern = re.compile(
        r"(-?\d+(?:\.\d+)?)\s*"
        r"(?:度|°)\s*"
        r"(\d+(?:\.\d+)?)?\s*"
        r"(?:分|'|′)?\s*"
        r"(\d+(?:\.\d+)?)?\s*"
        r"(?:秒|\"|″)?"
    )

    match = dms_pattern.search(cleaned)

    if match:
        decimal = decimal_from_dms(
            match.group(1),
            match.group(2) or 0,
            match.group(3) or 0,
        )

        if decimal is not None and negative:
            decimal = -abs(decimal)

        if coordinate_is_valid(decimal, coordinate_type):
            return decimal

    number_match = re.search(r"-?\d+(?:\.\d+)?", cleaned)

    if number_match:
        numeric = float(number_match.group())

        if negative:
            numeric = -abs(numeric)

        return choose_coordinate_candidate(numeric, coordinate_type)

    return None


def convert_coordinate_series(series, coordinate_type):
    return series.apply(lambda value: parse_coordinate_value(value, coordinate_type))


def find_header_row(raw_preview):
    max_rows = min(20, len(raw_preview))

    for row_index in range(max_rows):
        row_values = [
            normalize_column_name(value)
            for value in raw_preview.iloc[row_index].tolist()
        ]

        has_latitude = any("緯度" in value or "latitude" in value for value in row_values)
        has_longitude = any("経度" in value or "longitude" in value for value in row_values)

        if has_latitude and has_longitude:
            return row_index

    return 0


def read_excel_sheets_robustly(source_file):
    excel_file = pd.ExcelFile(source_file)
    candidates = []

    for sheet_name in excel_file.sheet_names:
        try:
            preview = pd.read_excel(
                source_file,
                sheet_name=sheet_name,
                header=None,
                nrows=20,
            )

            header_row = find_header_row(preview)

            sheet_df = pd.read_excel(
                source_file,
                sheet_name=sheet_name,
                header=header_row,
            )

            candidates.append((sheet_name, sheet_df))
        except Exception:
            continue

    return candidates


def prepare_accident_dataframe(raw_df):
    if raw_df is None or raw_df.empty:
        return pd.DataFrame(), "データが空です"

    latitude_column = find_first_column(
        raw_df.columns,
        ["緯度", "地点緯度", "地点_緯度", "事故地点緯度", "発生地点緯度", "latitude", "Latitude", "LAT", "lat"],
        ["緯度", "latitude", "地点緯度"],
    )

    longitude_column = find_first_column(
        raw_df.columns,
        ["経度", "地点経度", "地点_経度", "事故地点経度", "発生地点経度", "longitude", "Longitude", "LON", "lon", "lng"],
        ["経度", "longitude", "地点経度"],
    )

    result = raw_df.copy()

    if latitude_column is not None:
        result["latitude"] = convert_coordinate_series(result[latitude_column], "lat")

    if longitude_column is not None:
        result["longitude"] = convert_coordinate_series(result[longitude_column], "lon")

    if "latitude" not in result.columns or "longitude" not in result.columns:
        return pd.DataFrame(), "緯度・経度列を見つけられません"

    result = result.dropna(subset=["latitude", "longitude"])

    result = result[
        result["latitude"].between(20, 50)
        & result["longitude"].between(120, 155)
    ].copy()

    if result.empty:
        return pd.DataFrame(), "座標変換後に有効な日本国内の緯度・経度がありません"

    return result, None


@st.cache_data(show_spinner=False)
def load_official_accident_data():
    source_file = next((path for path in ACCIDENT_FILE_CANDIDATES if path.exists()), None)

    if source_file is None:
        return pd.DataFrame(), None, "ファイルが見つかりません", None

    try:
        if source_file.suffix.lower() in [".xlsx", ".xls"]:
            errors = []

            for sheet_name, raw_df in read_excel_sheets_robustly(source_file):
                prepared_df, prepare_error = prepare_accident_dataframe(raw_df)

                if not prepared_df.empty:
                    return prepared_df, source_file.name, None, sheet_name

                errors.append(f"{sheet_name}: {prepare_error}")

            return (
                pd.DataFrame(),
                source_file.name,
                " / ".join(errors) if errors else "読み取れるシートがありません",
                None,
            )

        raw_df = None

        for encoding in ["utf-8-sig", "cp932", "shift_jis", "utf-8"]:
            try:
                preview = pd.read_csv(
                    source_file,
                    encoding=encoding,
                    header=None,
                    nrows=20,
                    low_memory=False,
                )

                header_row = find_header_row(preview)

                raw_df = pd.read_csv(
                    source_file,
                    encoding=encoding,
                    header=header_row,
                    low_memory=False,
                )

                break
            except UnicodeDecodeError:
                continue

        if raw_df is None:
            return pd.DataFrame(), source_file.name, "CSVの文字コードを判定できません", None

        prepared_df, prepare_error = prepare_accident_dataframe(raw_df)

        return prepared_df, source_file.name, prepare_error, "CSV"

    except Exception as error:
        return pd.DataFrame(), source_file.name, f"読み込みエラー：{error}", None


official_accident_df, official_accident_file, official_accident_error, official_accident_sheet = load_official_accident_data()


def nearest_accident_information(current_latitude, current_longitude, accident_df):
    if accident_df.empty:
        return None

    working = accident_df.copy()

    working["distance_m"] = working.apply(
        lambda row: haversine_distance(
            current_latitude,
            current_longitude,
            row["latitude"],
            row["longitude"],
        ),
        axis=1,
    )

    nearest_row = working.loc[working["distance_m"].idxmin()]

    return {
        "nearest_distance_m": float(nearest_row["distance_m"]),
        "danger_count_200m": int((working["distance_m"] <= 200).sum()),
        "danger_count_500m": int((working["distance_m"] <= 500).sum()),
        "danger_count_1000m": int((working["distance_m"] <= 1000).sum()),
    }


# =========================================================
# 6. DB
# =========================================================

def get_connection():
    return sqlite3.connect(DB_FILE, check_same_thread=False)


def initialize_database():
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS danger_points (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            place TEXT NOT NULL,
            latitude REAL NOT NULL,
            longitude REAL NOT NULL,
            danger_level INTEGER NOT NULL,
            category TEXT NOT NULL,
            comment TEXT,
            image_name TEXT,
            status TEXT DEFAULT '公開',
            created_at TEXT NOT NULL
        )
        """
    )

    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS favorites (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            danger_point_id INTEGER UNIQUE,
            created_at TEXT NOT NULL
        )
        """
    )

    conn.commit()

    cursor.execute("SELECT COUNT(*) FROM danger_points")
    count = cursor.fetchone()[0]

    if count == 0:
        sample_data = [
            (
                "テスト危険地点：自転車の飛び出し",
                DEFAULT_LATITUDE,
                DEFAULT_LONGITUDE,
                5,
                "自転車の飛び出し",
                "歩道から自転車が車道へ急に出てくる可能性があります。",
                "",
                "公開",
                datetime.now().strftime("%Y-%m-%d %H:%M"),
            ),
            (
                "見通しの悪い交差点",
                DEFAULT_LATITUDE + 0.003,
                DEFAULT_LONGITUDE + 0.003,
                4,
                "見通しが悪い",
                "建物で見通しが悪く、自転車や歩行者に気づきにくい場所です。",
                "",
                "公開",
                datetime.now().strftime("%Y-%m-%d %H:%M"),
            ),
            (
                "交通量が多い道路",
                DEFAULT_LATITUDE - 0.003,
                DEFAULT_LONGITUDE - 0.003,
                3,
                "交通量が多い",
                "車と自転車が混在しやすい道路です。",
                "",
                "公開",
                datetime.now().strftime("%Y-%m-%d %H:%M"),
            ),
        ]

        cursor.executemany(
            """
            INSERT INTO danger_points (
                place, latitude, longitude, danger_level, category,
                comment, image_name, status, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            sample_data,
        )

        conn.commit()

    conn.close()


def load_danger_points(public_only=True):
    conn = get_connection()

    if public_only:
        query = """
            SELECT *
            FROM danger_points
            WHERE status = '公開'
            ORDER BY id DESC
        """
    else:
        query = """
            SELECT *
            FROM danger_points
            ORDER BY id DESC
        """

    df = pd.read_sql_query(query, conn)
    conn.close()

    return df


def insert_danger_point(place, latitude, longitude, danger_level, category, comment, image_name):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        INSERT INTO danger_points (
            place, latitude, longitude, danger_level, category,
            comment, image_name, status, created_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            place,
            latitude,
            longitude,
            danger_level,
            category,
            comment,
            image_name,
            "公開",
            datetime.now().strftime("%Y-%m-%d %H:%M"),
        ),
    )

    conn.commit()
    conn.close()


def delete_danger_point(point_id):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM favorites WHERE danger_point_id = ?", (point_id,))
    cursor.execute("DELETE FROM danger_points WHERE id = ?", (point_id,))

    conn.commit()
    conn.close()


def update_point_status(point_id, status):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        UPDATE danger_points
        SET status = ?
        WHERE id = ?
        """,
        (status, point_id),
    )

    conn.commit()
    conn.close()


def add_favorite(point_id):
    conn = get_connection()
    cursor = conn.cursor()

    try:
        cursor.execute(
            """
            INSERT INTO favorites (
                danger_point_id,
                created_at
            )
            VALUES (?, ?)
            """,
            (point_id, datetime.now().strftime("%Y-%m-%d %H:%M")),
        )

        conn.commit()
        result = True
    except sqlite3.IntegrityError:
        result = False

    conn.close()
    return result


def remove_favorite(point_id):
    conn = get_connection()
    cursor = conn.cursor()

    cursor.execute(
        """
        DELETE FROM favorites
        WHERE danger_point_id = ?
        """,
        (point_id,),
    )

    conn.commit()
    conn.close()


def load_favorites():
    conn = get_connection()

    query = """
        SELECT d.*
        FROM danger_points d
        INNER JOIN favorites f
            ON d.id = f.danger_point_id
        ORDER BY f.id DESC
    """

    df = pd.read_sql_query(query, conn)
    conn.close()

    return df


initialize_database()


# =========================================================
# 7. 判定・地図・音声
# =========================================================

def get_safety_aura_dangers(danger_df, current_latitude, current_longitude, radius_m, minimum_level):
    if danger_df.empty:
        return danger_df.copy()

    result = danger_df.copy()

    result["distance_m"] = result.apply(
        lambda row: haversine_distance(
            current_latitude,
            current_longitude,
            row["latitude"],
            row["longitude"],
        ),
        axis=1,
    )

    result = result[
        (result["distance_m"] <= radius_m)
        & (result["danger_level"] >= minimum_level)
    ]

    return result.sort_values(["danger_level", "distance_m"], ascending=[False, True])


def get_user_danger_candidates(current_latitude, current_longitude):
    danger_df = load_danger_points()

    if danger_df.empty:
        return danger_df.copy()

    result = danger_df.copy()

    result["distance_m"] = result.apply(
        lambda row: haversine_distance(
            current_latitude,
            current_longitude,
            row["latitude"],
            row["longitude"],
        ),
        axis=1,
    )

    return result.sort_values("distance_m")


def get_car_mode_status():
    if (
        st.session_state.current_latitude is None
        or st.session_state.current_longitude is None
    ):
        return {
            "level": "waiting",
            "message": "GPSで現在地を取得してください",
            "nearest_distance": None,
            "nearest_name": None,
            "bicycle_alerts": [],
            "voice_text": "",
        }

    current_latitude = st.session_state.current_latitude
    current_longitude = st.session_state.current_longitude

    nearest_distance = None
    nearest_name = None

    if not official_accident_df.empty:
        info = nearest_accident_information(
            current_latitude,
            current_longitude,
            official_accident_df,
        )

        if info:
            nearest_distance = info["nearest_distance_m"]
            nearest_name = "公式事故地点"

    user_candidates = get_user_danger_candidates(current_latitude, current_longitude)

    if not user_candidates.empty:
        first_user = user_candidates.iloc[0]
        user_distance = float(first_user["distance_m"])

        if nearest_distance is None or user_distance < nearest_distance:
            nearest_distance = user_distance
            nearest_name = str(first_user["place"])

    bicycle_alerts = []

    if not user_candidates.empty:
        bicycle_keywords = ["自転車", "飛び出し", "カットイン", "交通量が多い", "見通しが悪い"]

        for _, row in user_candidates.iterrows():
            text_all = (
                str(row.get("category", ""))
                + str(row.get("comment", ""))
                + str(row.get("place", ""))
            )

            keyword_hit = any(keyword in text_all for keyword in bicycle_keywords)

            if keyword_hit and float(row["distance_m"]) <= st.session_state.car_bicycle_alert_radius:
                bicycle_alerts.append(row)

    if nearest_distance is None:
        return {
            "level": "safe",
            "message": "周辺の危険地点データがありません",
            "nearest_distance": None,
            "nearest_name": None,
            "bicycle_alerts": bicycle_alerts,
            "voice_text": "",
        }

    if nearest_distance <= st.session_state.car_danger_radius or len(bicycle_alerts) >= 1:
        return {
            "level": "danger",
            "message": "危険地点が近くにあります。速度を落として周囲を確認してください。",
            "nearest_distance": nearest_distance,
            "nearest_name": nearest_name,
            "bicycle_alerts": bicycle_alerts,
            "voice_text": "警告。危険地点が近くにあります。速度を落として、歩行者と自転車に注意してください。",
        }

    if nearest_distance <= st.session_state.car_warning_radius:
        return {
            "level": "warning",
            "message": "この先に危険地点があります。注意して走行してください。",
            "nearest_distance": nearest_distance,
            "nearest_name": nearest_name,
            "bicycle_alerts": bicycle_alerts,
            "voice_text": "この先に危険地点があります。注意して走行してください。",
        }

    return {
        "level": "safe",
        "message": "現在地周辺に近い危険地点はありません",
        "nearest_distance": nearest_distance,
        "nearest_name": nearest_name,
        "bicycle_alerts": bicycle_alerts,
        "voice_text": "",
    }


def speak_warning_once(text_to_speak, key):
    if not text_to_speak or not st.session_state.car_voice_enabled:
        return

    now = datetime.now()
    last_time = st.session_state.last_voice_time
    last_key = st.session_state.last_voice_key

    if last_key == key and last_time is not None:
        if (now - last_time).total_seconds() < 60:
            return

    st.session_state.last_voice_key = key
    st.session_state.last_voice_time = now

    safe_text = (
        text_to_speak
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", " ")
    )

    components.html(
        f"""
        <script>
        const msg = new SpeechSynthesisUtterance("{safe_text}");
        msg.lang = "ja-JP";
        msg.rate = 1.0;
        msg.pitch = 1.0;
        window.speechSynthesis.cancel();
        window.speechSynthesis.speak(msg);
        </script>
        """,
        height=0,
    )



def speak_bicycle_aura_warning_once(text_to_speak, key):
    """自転車モードのセーフティ・オーラ用音声警告"""
    if not text_to_speak:
        return

    if not st.session_state.bike_aura_voice_enabled:
        return

    now = datetime.now()
    last_time = st.session_state.last_bike_aura_voice_time
    last_key = st.session_state.last_bike_aura_voice_key

    if last_key == key and last_time is not None:
        # 同じ警告は60秒間くり返さない
        if (now - last_time).total_seconds() < 60:
            return

    st.session_state.last_bike_aura_voice_key = key
    st.session_state.last_bike_aura_voice_time = now

    safe_text = (
        text_to_speak
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\n", " ")
    )

    components.html(
        f"""
        <script>
        const msg = new SpeechSynthesisUtterance("{safe_text}");
        msg.lang = "ja-JP";
        msg.rate = 1.0;
        msg.pitch = 1.0;
        window.speechSynthesis.cancel();
        window.speechSynthesis.speak(msg);
        </script>
        """,
        height=0,
    )


def create_danger_map(
    danger_df,
    center_lat=DEFAULT_LATITUDE,
    center_lon=DEFAULT_LONGITUDE,
    zoom=13,
    current_location=None,
    safety_aura_enabled=False,
    safety_aura_radius=200,
    safety_aura_min_level=3,
    accident_df=None,
    show_official_accidents=False,
):
    map_object = folium.Map(
        location=[center_lat, center_lon],
        zoom_start=zoom,
        control_scale=True,
        tiles="OpenStreetMap",
    )

    for _, row in danger_df.iterrows():
        popup_html = f"""
        <div style="width:260px;">
            <h4>{row['place']}</h4>
            <p><b>危険度：</b>{row['danger_level']}／5</p>
            <p><b>分類：</b>{row['category']}</p>
            <p>{row['comment']}</p>
            <small>{row['created_at']}</small>
        </div>
        """

        folium.Marker(
            location=[row["latitude"], row["longitude"]],
            tooltip=f"{row['place']}｜危険度 {row['danger_level']}",
            popup=folium.Popup(popup_html, max_width=320),
            icon=folium.Icon(color=danger_color(int(row["danger_level"])), icon="warning-sign"),
        ).add_to(map_object)

        folium.Circle(
            location=[row["latitude"], row["longitude"]],
            radius=40 + int(row["danger_level"]) * 15,
            color=danger_color(int(row["danger_level"])),
            fill=True,
            fill_opacity=0.12,
            weight=1,
        ).add_to(map_object)

    if show_official_accidents and accident_df is not None and not accident_df.empty:
        accident_sample = accident_df

        if current_location:
            lat = current_location["latitude"]
            lon = current_location["longitude"]
        else:
            lat = center_lat
            lon = center_lon

        # 現在地または地図中心から約8km以内の事故地点を優先表示
        lat_range = 8 / 111
        lon_range = 8 / max(111 * cos(radians(lat)), 1)

        nearby_accidents = accident_df[
            accident_df["latitude"].between(lat - lat_range, lat + lat_range)
            & accident_df["longitude"].between(lon - lon_range, lon + lon_range)
        ]

        if not nearby_accidents.empty:
            accident_sample = nearby_accidents

        if len(accident_sample) > MAX_ACCIDENT_MAP_POINTS:
            accident_sample = accident_sample.sample(MAX_ACCIDENT_MAP_POINTS, random_state=42)

        accident_layer = folium.FeatureGroup(name="公式事故地点", show=True)

        for _, accident_row in accident_sample.iterrows():
            folium.CircleMarker(
                location=[accident_row["latitude"], accident_row["longitude"]],
                radius=6,
                color="#d00000",
                fill=True,
                fill_color="#ff2d2d",
                fill_opacity=0.9,
                weight=2,
                tooltip="公式事故地点",
            ).add_to(accident_layer)

        accident_layer.add_to(map_object)

    if current_location:
        folium.Marker(
            location=[current_location["latitude"], current_location["longitude"]],
            tooltip="現在地",
            popup="あなたの現在地",
            icon=folium.Icon(color="blue", icon="user"),
        ).add_to(map_object)

        if safety_aura_enabled:
            aura_color = "blue"
            aura_danger_count = 0

            for _, row in danger_df.iterrows():
                distance = haversine_distance(
                    current_location["latitude"],
                    current_location["longitude"],
                    row["latitude"],
                    row["longitude"],
                )

                if distance <= safety_aura_radius and int(row["danger_level"]) >= safety_aura_min_level:
                    aura_danger_count += 1

            if aura_danger_count >= 3:
                aura_color = "red"
            elif aura_danger_count >= 1:
                aura_color = "orange"

            folium.Circle(
                location=[current_location["latitude"], current_location["longitude"]],
                radius=safety_aura_radius,
                color=aura_color,
                fill=True,
                fill_opacity=0.08,
                weight=3,
                tooltip=f"警告範囲 {safety_aura_radius}m",
            ).add_to(map_object)

    folium.LayerControl().add_to(map_object)

    return map_object


def display_point_card(row, show_favorite=True, key_prefix="point"):
    level = int(row["danger_level"])

    if level >= 4:
        card_class = "danger-card"
    elif level == 3:
        card_class = "warning-card"
    else:
        card_class = "safe-card"

    st.markdown(
        f"""
        <div class="{card_class}">
            <h4 style="margin:0 0 7px 0;">📍 {row["place"]}</h4>
            <b>危険度：</b>{level}／5（{danger_text(level)}）<br>
            <b>分類：</b>{row["category"]}<br>
            <b>内容：</b>{row["comment"] or "説明なし"}<br>
            <small>登録日時：{row["created_at"]}</small>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if show_favorite:
        if st.button(
            "⭐ お気に入りに追加",
            key=f"{key_prefix}_favorite_{row['id']}",
            use_container_width=True,
        ):
            if add_favorite(int(row["id"])):
                st.success("お気に入りに追加しました。")
            else:
                st.info("すでにお気に入りへ登録されています。")


def geocode_address(address):
    if not address:
        return None

    url = "https://nominatim.openstreetmap.org/search"

    params = {
        "q": address,
        "format": "json",
        "limit": 1,
        "countrycodes": "jp",
    }

    headers = {
        "User-Agent": "RescueRide-Bicycle-Safety-App/1.0"
    }

    try:
        response = requests.get(url, params=params, headers=headers, timeout=15)
        response.raise_for_status()

        results = response.json()

        if not results:
            return None

        return {
            "latitude": float(results[0]["lat"]),
            "longitude": float(results[0]["lon"]),
            "display_name": results[0]["display_name"],
        }

    except requests.RequestException:
        return None



def calculate_route_danger_score(route_coordinates, danger_df, threshold_m=120):
    """
    ルート周辺の危険地点を数える。
    route_coordinates は [[lat, lon], ...] の形式。
    """
    if danger_df.empty or not route_coordinates:
        return {
            "danger_count": 0,
            "high_danger_count": 0,
            "nearby_points": pd.DataFrame(),
        }

    danger_rows = []

    # 全座標を使うと重くなるので、最大120点に間引く
    step = max(1, len(route_coordinates) // 120)
    sampled_route = route_coordinates[::step]

    work_df = danger_df.copy()

    for _, danger_row in work_df.iterrows():
        min_distance = None

        for route_lat, route_lon in sampled_route:
            distance = haversine_distance(
                route_lat,
                route_lon,
                danger_row["latitude"],
                danger_row["longitude"],
            )

            if min_distance is None or distance < min_distance:
                min_distance = distance

        if min_distance is not None and min_distance <= threshold_m:
            row_dict = danger_row.to_dict()
            row_dict["route_distance_m"] = min_distance
            danger_rows.append(row_dict)

    if not danger_rows:
        nearby_points = pd.DataFrame()
    else:
        nearby_points = pd.DataFrame(danger_rows).sort_values("route_distance_m")

    high_danger_count = (
        len(nearby_points[nearby_points["danger_level"] >= 4])
        if not nearby_points.empty
        else 0
    )

    return {
        "danger_count": len(nearby_points),
        "high_danger_count": high_danger_count,
        "nearby_points": nearby_points,
    }


def get_route_type_message(route_type):
    if route_type == "最短ルート":
        return {
            "title": "最短ルート",
            "message": "距離と時間を優先して、目的地までの基本ルートを表示します。",
        }

    if route_type == "危険地点を避けるルート":
        return {
            "title": "危険地点を避けるルート",
            "message": "ルート周辺の危険地点を確認し、危険地点が近い場合は警告します。",
        }

    return {
        "title": "上り坂少ないルート",
        "message": "坂道負担を減らすルートとして選択できます。現在は地形データ未連携のため、参考表示です。",
    }



def get_bicycle_route(start_lat, start_lon, end_lat, end_lon):
    url = (
        "https://router.project-osrm.org/route/v1/driving/"
        f"{start_lon},{start_lat};{end_lon},{end_lat}"
    )

    params = {
        "overview": "full",
        "geometries": "geojson",
        "steps": "true",
    }

    try:
        response = requests.get(url, params=params, timeout=20)
        response.raise_for_status()

        data = response.json()

        if data.get("code") != "Ok":
            return None

        route = data["routes"][0]

        return {
            "coordinates": route["geometry"]["coordinates"],
            "distance": route["distance"],
            "duration": route["duration"],
        }

    except (requests.RequestException, KeyError, IndexError):
        return None


# =========================================================
# 8. 画面共通
# =========================================================


def make_qr_code_image(url):
    """URLからQRコード画像を作成する"""
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=4,
    )

    qr.add_data(url)
    qr.make(fit=True)

    image = qr.make_image(
        fill_color="black",
        back_color="white",
    )

    buffer = BytesIO()
    image.save(buffer, format="PNG")
    buffer.seek(0)

    return buffer


def render_smartphone_qr_panel(location_key="common"):
    """スマホでアプリを開くためのQRコード表示"""
    st.markdown(
        """
        <div class="info-card">
            <h3>📱 スマホで開くQRコード</h3>
            PCで起動しているアプリのURLを入力すると、
            スマホで読み取れるQRコードを表示します。
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.info(
        "Streamlit Cloudで公開している場合は、その公開URLを入力してください。"
        "ローカルPCで動かしている場合は、スマホとPCを同じWi-Fiに接続し、"
        "PCのIPアドレスを使ったURLを入力します。例：http://192.168.1.10:8501"
    )

    url = st.text_input(
        "スマホで開きたいアプリURL",
        value=st.session_state.get("share_app_url", ""),
        placeholder="例：https://your-app.streamlit.app または http://192.168.1.10:8501",
        key=f"share_app_url_input_{location_key}",
    )

    if url:
        st.session_state.share_app_url = url

        qr_buffer = make_qr_code_image(url)

        left, right = st.columns([1, 1.4])

        with left:
            st.image(
                qr_buffer,
                caption="このQRコードをスマホで読み取ってください",
                width=260,
            )

        with right:
            st.success("QRコードを作成しました。")
            st.write("スマホのカメラで読み取ると、このアプリを開けます。")
            st.code(url, language="text")

            st.download_button(
                "QRコード画像をダウンロード",
                data=qr_buffer.getvalue(),
                file_name="rescue_ride_qr.png",
                mime="image/png",
                use_container_width=True,
                key=f"qr_download_{location_key}",
            )
    else:
        st.warning("URLを入力するとQRコードが表示されます。")



def render_header(title, subtitle, mode_label):
    top_left, top_right = st.columns([4, 1.4])

    with top_left:
        st.markdown(f'<div class="app-title">{title}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="app-subtitle">{subtitle}</div>', unsafe_allow_html=True)

    with top_right:
        st.success(f"現在：{mode_label}")

    st.markdown(
        """
        <div class="top-mode-bar">
            モード変更：下のボタンで別モードへ移動できます。
        </div>
        """,
        unsafe_allow_html=True,
    )

    b1, b2, b3 = st.columns([1, 1, 1.2])

    with b1:
        if st.button("🚲 自転車モードへ", use_container_width=True, key=f"header_bike_{mode_label}"):
            st.session_state.app_mode = "自転車モード"
            st.session_state.mode_selected = True
            st.rerun()

    with b2:
        if st.button("🚗 自動車モードへ", use_container_width=True, key=f"header_car_{mode_label}"):
            st.session_state.app_mode = "自動車モード"
            st.session_state.mode_selected = True
            st.rerun()

    with b3:
        if st.button("🔁 最初の選択画面へ戻る", use_container_width=True, key=f"header_back_{mode_label}"):
            st.session_state.mode_selected = False
            st.session_state.app_mode = None
            st.rerun()


def render_first_mode_select_screen():
    st.markdown(
        """
        <div class="mode-select-wrap">
            <div class="mode-select-card">
                <div class="mode-select-title">🚲 レスキューライド</div>
                <div class="mode-select-subtitle">
                    使うモードを選んでください。選んだモードによって、次の画面が完全に変わります。
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns(2)

    with left:
        st.markdown(
            """
            <div class="mode-box">
                <div class="mode-box-title">🚲 自転車モード</div>
                <div class="mode-box-text">
                    自転車で走る人向け。<br>
                    ホーム、危険マップ、GPS確認、危険地点投稿、
                    ルート検索、お気に入りなどを使います。
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("🚲 自転車モードで始める", type="primary", use_container_width=True):
            st.session_state.app_mode = "自転車モード"
            st.session_state.mode_selected = True
            st.rerun()

    with right:
        st.markdown(
            """
            <div class="mode-box">
                <div class="mode-box-title">🚗 自動車モード</div>
                <div class="mode-box-text">
                    車を運転する人向け。<br>
                    運転メイン、危険マップ、GPS警告、
                    自転車接近、車モード設定だけを表示します。
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        if st.button("🚗 自動車モードで始める", type="primary", use_container_width=True):
            st.session_state.app_mode = "自動車モード"
            st.session_state.mode_selected = True
            st.rerun()

    st.divider()

    with st.expander("📱 スマホで開くQRコードを表示", expanded=False):
        render_smartphone_qr_panel("first_screen")


def render_gps_input_block(prefix):
    if GPS_AVAILABLE:
        gps_result = streamlit_geolocation()

        if gps_result:
            latitude = gps_result.get("latitude")
            longitude = gps_result.get("longitude")

            if latitude is not None and longitude is not None:
                st.session_state.current_latitude = float(latitude)
                st.session_state.current_longitude = float(longitude)
                st.success("GPSで現在地を取得しました。")
    else:
        st.warning("GPS機能のライブラリが入っていません。requirements.txtを確認してください。")

    st.write("#### 緯度・経度を直接入力")

    manual_latitude = st.number_input(
        "現在地の緯度",
        min_value=-90.0,
        max_value=90.0,
        value=float(st.session_state.current_latitude if st.session_state.current_latitude is not None else DEFAULT_LATITUDE),
        format="%.7f",
        key=f"{prefix}_manual_latitude",
    )

    manual_longitude = st.number_input(
        "現在地の経度",
        min_value=-180.0,
        max_value=180.0,
        value=float(st.session_state.current_longitude if st.session_state.current_longitude is not None else DEFAULT_LONGITUDE),
        format="%.7f",
        key=f"{prefix}_manual_longitude",
    )

    if st.button("この位置を現在地に設定", use_container_width=True, key=f"{prefix}_set_location"):
        st.session_state.current_latitude = manual_latitude
        st.session_state.current_longitude = manual_longitude
        st.success("現在地を設定しました。")
        st.rerun()


# =========================================================
# 9. 自転車モード専用アプリ
# =========================================================


def render_safety_aura_fixed_top_panel():
    """自転車モードの一番上に必ず出すセーフティ・オーラ"""

    # GPS未取得
    if (
        st.session_state.current_latitude is None
        or st.session_state.current_longitude is None
    ):
        st.markdown(
            """
            <div class="safety-aura-fixed-panel safety-aura-fixed-waiting">
                <div class="safety-aura-fixed-title">🛡️ SAFETY AURA 待機中</div>
                <div class="safety-aura-fixed-text">
                    GPS確認タブで現在地を取得すると、周辺の危険地点を自動判定します。
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        c1, c2, c3 = st.columns(3)
        c1.metric("オーラ半径", f"{st.session_state.safety_aura_radius}m")
        c2.metric("警告対象", f"危険度{st.session_state.safety_aura_min_level}以上")
        c3.metric("現在地", "未取得")
        return

    danger_df = load_danger_points()

    # 危険地点データなし
    if danger_df.empty:
        st.markdown(
            """
            <div class="safety-aura-fixed-panel safety-aura-fixed-safe">
                <div class="safety-aura-fixed-title">✅ SAFETY AURA 正常</div>
                <div class="safety-aura-fixed-text">
                    登録済みの危険地点はまだありません。
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        return

    # 距離計算
    work_df = danger_df.copy()
    work_df["distance_m"] = work_df.apply(
        lambda row: haversine_distance(
            st.session_state.current_latitude,
            st.session_state.current_longitude,
            row["latitude"],
            row["longitude"],
        ),
        axis=1,
    )

    nearest_distance = float(work_df["distance_m"].min())

    aura_df = work_df[
        (work_df["distance_m"] <= st.session_state.safety_aura_radius)
        & (work_df["danger_level"] >= st.session_state.safety_aura_min_level)
    ].copy()

    if aura_df.empty:
        panel_class = "safety-aura-fixed-safe"
        title = "✅ SAFETY AURA 正常"
        message = (
            f"半径{st.session_state.safety_aura_radius}m以内に、"
            f"危険度{st.session_state.safety_aura_min_level}以上の危険地点はありません。"
        )
    else:
        severe_count = len(aura_df[aura_df["danger_level"] >= 4])

        if severe_count >= 1:
            panel_class = "safety-aura-fixed-danger"
            title = "⚠️ SAFETY AURA 発動中"
            message = (
                f"オーラ内に危険地点が{len(aura_df)}件あります。"
                f"危険度4以上は{severe_count}件です。"
            )
        else:
            panel_class = "safety-aura-fixed-warning"
            title = "⚠️ SAFETY AURA 注意中"
            message = f"オーラ内に危険地点が{len(aura_df)}件あります。注意して走行してください。"

    st.markdown(
        f"""
        <div class="safety-aura-fixed-panel {panel_class}">
            <div class="safety-aura-fixed-title">{title}</div>
            <div class="safety-aura-fixed-text">
                {message}<br>
                最寄り登録危険地点まで約{nearest_distance:.0f}m
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # セーフティ・オーラ音声警告
    if panel_class == "safety-aura-fixed-danger":
        speak_bicycle_aura_warning_once(
            "警告。セーフティオーラが発動しました。危険地点が近くにあります。速度を落として周囲を確認してください。",
            f"bike_aura_danger_{int(nearest_distance)}",
        )
    elif panel_class == "safety-aura-fixed-warning":
        speak_bicycle_aura_warning_once(
            "注意。セーフティオーラ内に危険地点があります。注意して走行してください。",
            f"bike_aura_warning_{int(nearest_distance)}",
        )

    c1, c2, c3 = st.columns(3)
    c1.metric("オーラ半径", f"{st.session_state.safety_aura_radius}m")
    c2.metric("警告対象", f"危険度{st.session_state.safety_aura_min_level}以上")
    c3.metric("オーラ内危険地点", f"{len(aura_df)}件")

    st.caption(
        "🔊 セーフティ・オーラ音声警告："
        + ("ON" if st.session_state.bike_aura_voice_enabled else "OFF")
        + "（同じ警告は60秒間くり返しません）"
    )



def render_bicycle_app():
    render_header(
        "🚲 レスキューライド 自転車モード",
        "自転車利用者向け：危険地点共有・セーフティオーラ・ルート確認",
        "🚲 自転車モード",
    )

    # ここで必ずセーフティ・オーラを表示する
    render_safety_aura_fixed_top_panel()

    st.markdown(
        """
        <div class="bike-hero">
            <div class="bike-hero-title">🚲 自転車モード専用画面</div>
            <div class="bike-hero-text">
                ここでは自転車で走る人向けに、危険マップ、GPS確認、危険地点投稿、ルート検索を使えます。
                自動車モードの運転警告画面は表示しません。
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    tabs = st.tabs(
        [
            "🏠 ホーム",
            "🗺️ 危険マップ",
            "📍 GPS確認",
            "⚠️ 危険地点投稿",
            "🧭 ルート検索",
            "⭐ お気に入り",
            "📊 データ確認",
            "🔐 管理者",
            "📱 スマホ共有",
        ]
    )

    with tabs[0]:
        danger_df = load_danger_points()

        st.header("ホーム")

        ride_kun_image = Path("ride_kun.png")

        left, right = st.columns([1.5, 1])

        with left:
            st.markdown(
                """
                <div class="info-card">
                    <h2 style="margin-top:0;">こんにちは！ライド君です 🚲</h2>
                    <p style="font-size:18px; line-height:1.8;">
                        自転車モードでは、現在地周辺の危険地点やセーフティ・オーラを確認できます。
                        危険地点の投稿やルート検索もできます。
                    </p>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with right:
            if ride_kun_image.exists():
                st.image(str(ride_kun_image), caption="ライド君", width=300)

        st.divider()

        c1, c2, c3, c4 = st.columns(4)

        c1.metric("危険地点", f"{len(danger_df)}件")
        c2.metric("危険度4以上", f"{len(danger_df[danger_df['danger_level'] >= 4])}件")
        c3.metric("公式事故データ", f"{len(official_accident_df)}件" if not official_accident_df.empty else "未読込")
        c4.metric("現在地", "取得済み" if st.session_state.current_latitude is not None else "未取得")

        st.subheader("最近登録された危険地点")

        if danger_df.empty:
            st.info("危険地点はまだ登録されていません。")
        else:
            for _, row in danger_df.head(3).iterrows():
                display_point_card(row, show_favorite=False, key_prefix="bike_home")

    with tabs[1]:
        st.header("危険マップ")

        danger_df = load_danger_points()

        show_official_accidents = st.checkbox("公式事故データを表示", value=False)

        if show_official_accidents:
            if official_accident_df.empty:
                st.error(
                    f"公式事故データを読み込めていません。"
                    f"app.pyと同じフォルダに honhyo_2024.xlsx を置いてください。"
                    f" 詳細：{official_accident_error}"
                )
            else:
                st.success(
                    f"公式事故データを読み込み済み：{len(official_accident_df):,}件。"
                    f"赤い点で地図上に表示します。"
                )

        current_location = get_current_location()

        center_lat = st.session_state.current_latitude if current_location else DEFAULT_LATITUDE
        center_lon = st.session_state.current_longitude if current_location else DEFAULT_LONGITUDE

        danger_map = create_danger_map(
            danger_df=danger_df,
            center_lat=center_lat,
            center_lon=center_lon,
            zoom=15 if current_location else 12,
            current_location=current_location,
            safety_aura_enabled=st.session_state.safety_aura_enabled and current_location is not None,
            safety_aura_radius=st.session_state.safety_aura_radius,
            safety_aura_min_level=st.session_state.safety_aura_min_level,
            accident_df=official_accident_df,
            show_official_accidents=show_official_accidents,
        )

        st_folium(danger_map, width=None, height=590, returned_objects=[], key="bike_map")

        with st.expander("危険地点一覧"):
            if danger_df.empty:
                st.info("危険地点はありません。")
            else:
                for _, row in danger_df.iterrows():
                    display_point_card(row, show_favorite=True, key_prefix="bike_map")

    with tabs[2]:
        st.header("GPS確認・セーフティオーラ")

        col1, col2 = st.columns([1, 1.4])

        with col1:
            render_gps_input_block("bike_gps")

            st.divider()

            st.session_state.safety_aura_enabled = st.toggle(
                "セーフティ・オーラを表示",
                value=st.session_state.safety_aura_enabled,
            )

            st.session_state.bike_aura_voice_enabled = st.toggle(
                "セーフティ・オーラ音声警告",
                value=st.session_state.bike_aura_voice_enabled,
            )

            st.session_state.safety_aura_radius = st.slider(
                "オーラ半径",
                50,
                1000,
                int(st.session_state.safety_aura_radius),
                step=50,
                format="%dメートル",
            )

            st.session_state.safety_aura_min_level = st.select_slider(
                "警告対象の最低危険度",
                options=[1, 2, 3, 4, 5],
                value=int(st.session_state.safety_aura_min_level),
            )

            if st.button(
                "🔊 セーフティ・オーラ音声テスト",
                use_container_width=True,
                key="bike_aura_voice_test",
            ):
                speak_bicycle_aura_warning_once(
                    "セーフティオーラの音声警告テストです。危険地点に注意してください。",
                    f"bike_aura_test_{datetime.now().timestamp()}",
                )
                st.success("音声警告テストを実行しました。音が出ない場合は、ブラウザ画面を一度クリックしてから試してください。")

        with col2:
            current_location = get_current_location()

            if current_location is None:
                st.warning("現在地を取得または入力してください。")
            else:
                danger_df = load_danger_points()

                gps_map = create_danger_map(
                    danger_df=danger_df,
                    center_lat=st.session_state.current_latitude,
                    center_lon=st.session_state.current_longitude,
                    zoom=15,
                    current_location=current_location,
                    safety_aura_enabled=st.session_state.safety_aura_enabled,
                    safety_aura_radius=st.session_state.safety_aura_radius,
                    safety_aura_min_level=st.session_state.safety_aura_min_level,
                    accident_df=official_accident_df,
                    show_official_accidents=True,
                )

                st_folium(gps_map, width=None, height=520, returned_objects=[], key="bike_gps_map")

                aura_df = get_safety_aura_dangers(
                    danger_df,
                    st.session_state.current_latitude,
                    st.session_state.current_longitude,
                    st.session_state.safety_aura_radius,
                    st.session_state.safety_aura_min_level,
                )

                if aura_df.empty:
                    st.success("オーラ内に警告対象の危険地点はありません。")
                else:
                    st.error(f"オーラ内に危険地点が {len(aura_df)}件あります。")
                    for _, row in aura_df.head(5).iterrows():
                        st.write(f"⚠️ {row['place']}｜約{row['distance_m']:.0f}m")

    with tabs[3]:
        st.header("危険地点を投稿")

        with st.form("bike_danger_point_form", clear_on_submit=True):
            left, right = st.columns(2)

            with left:
                place = st.text_input("場所・施設名")
                address = st.text_input("住所から検索する場合")
                latitude = st.number_input(
                    "緯度",
                    min_value=-90.0,
                    max_value=90.0,
                    value=float(st.session_state.current_latitude if st.session_state.current_latitude is not None else DEFAULT_LATITUDE),
                    format="%.7f",
                )
                longitude = st.number_input(
                    "経度",
                    min_value=-180.0,
                    max_value=180.0,
                    value=float(st.session_state.current_longitude if st.session_state.current_longitude is not None else DEFAULT_LONGITUDE),
                    format="%.7f",
                )

            with right:
                danger_level = st.slider("危険度", 1, 5, 3)
                category = st.selectbox(
                    "危険の分類",
                    [
                        "交通量が多い",
                        "見通しが悪い",
                        "歩行者が多い",
                        "車の飛び出し",
                        "自転車の飛び出し",
                        "急な飛び出し・カットイン",
                        "道路が狭い",
                        "路面状態が悪い",
                        "信号・標識が分かりにくい",
                        "夜間が暗い",
                        "工事中",
                        "その他",
                    ],
                )
                comment = st.text_area("危険な状況の説明", height=130)
                uploaded_image = st.file_uploader("現地写真（任意）", type=["png", "jpg", "jpeg"])

            submitted = st.form_submit_button("危険地点を登録する", use_container_width=True)

        if submitted:
            final_latitude = latitude
            final_longitude = longitude

            if address.strip():
                with st.spinner("住所から位置を検索しています..."):
                    geocoded = geocode_address(address)

                if geocoded:
                    final_latitude = geocoded["latitude"]
                    final_longitude = geocoded["longitude"]
                    st.info(f"検索された住所：{geocoded['display_name']}")

            if not place.strip():
                st.error("場所・施設名を入力してください。")
            elif not comment.strip():
                st.error("危険な状況の説明を入力してください。")
            else:
                insert_danger_point(
                    place=place.strip(),
                    latitude=float(final_latitude),
                    longitude=float(final_longitude),
                    danger_level=int(danger_level),
                    category=category,
                    comment=comment.strip(),
                    image_name=uploaded_image.name if uploaded_image is not None else "",
                )
                st.success("危険地点を登録しました。")
                st.balloons()

    with tabs[4]:
        st.header("ルート検索")

        left, right = st.columns([1, 1.6])

        with left:
            start_address = st.text_input("出発地", placeholder="例：八王子駅")
            goal_address = st.text_input("目的地", placeholder="例：創価大学")

            route_type = st.selectbox(
                "ルートの出し方",
                [
                    "最短ルート",
                    "危険地点を避けるルート",
                    "上り坂少ないルート",
                ],
                index=0,
                help="目的に合わせてルート検索の方針を選べます。",
            )

            route_type_info = get_route_type_message(route_type)

            st.info(
                f"選択中：{route_type_info['title']}｜"
                f"{route_type_info['message']}"
            )

            use_current_location = st.checkbox("現在地を出発地として使用", value=False)

            if st.button("ルートを検索", type="primary", use_container_width=True):
                st.session_state.selected_route_type = route_type
                start_result = None
                goal_result = None

                if use_current_location and st.session_state.current_latitude is not None:
                    start_result = {
                        "latitude": st.session_state.current_latitude,
                        "longitude": st.session_state.current_longitude,
                        "display_name": "現在地",
                    }
                elif start_address.strip():
                    with st.spinner("出発地を検索しています..."):
                        start_result = geocode_address(start_address)

                if goal_address.strip():
                    with st.spinner("目的地を検索しています..."):
                        goal_result = geocode_address(goal_address)

                if start_result is None:
                    st.error("出発地を確認できませんでした。")
                elif goal_result is None:
                    st.error("目的地を確認できませんでした。")
                else:
                    st.session_state.start_location = start_result
                    st.session_state.goal_location = goal_result

        with right:
            start_location = st.session_state.start_location
            goal_location = st.session_state.goal_location

            if start_location and goal_location:
                route_result = get_bicycle_route(
                    start_location["latitude"],
                    start_location["longitude"],
                    goal_location["latitude"],
                    goal_location["longitude"],
                )

                center_latitude = (start_location["latitude"] + goal_location["latitude"]) / 2
                center_longitude = (start_location["longitude"] + goal_location["longitude"]) / 2

                route_map = folium.Map(location=[center_latitude, center_longitude], zoom_start=13, control_scale=True)

                folium.Marker(
                    [start_location["latitude"], start_location["longitude"]],
                    tooltip="出発地",
                    popup=start_location["display_name"],
                    icon=folium.Icon(color="green", icon="play"),
                ).add_to(route_map)

                folium.Marker(
                    [goal_location["latitude"], goal_location["longitude"]],
                    tooltip="目的地",
                    popup=goal_location["display_name"],
                    icon=folium.Icon(color="red", icon="flag"),
                ).add_to(route_map)

                if route_result:
                    route_coordinates = [
                        [coordinate[1], coordinate[0]]
                        for coordinate in route_result["coordinates"]
                    ]

                    folium.PolyLine(route_coordinates, weight=6, opacity=0.8, tooltip="検索ルート").add_to(route_map)
                    route_map.fit_bounds(route_coordinates)

                    c1, c2, c3 = st.columns(3)
                    c1.metric("ルート距離", f"{route_result['distance'] / 1000:.1f} km")
                    c2.metric("推定移動時間", f"{route_result['duration'] / 60:.0f} 分")

                    selected_route_type = st.session_state.get(
                        "selected_route_type",
                        "最短ルート",
                    )

                    c3.metric("ルート種類", selected_route_type)

                    danger_df_for_route = load_danger_points()
                    route_danger_info = calculate_route_danger_score(
                        route_coordinates,
                        danger_df_for_route,
                        threshold_m=120,
                    )

                    if selected_route_type == "最短ルート":
                        st.success(
                            "最短ルートを表示しています。距離と時間を優先した基本ルートです。"
                        )

                    elif selected_route_type == "危険地点を避けるルート":
                        if route_danger_info["danger_count"] == 0:
                            st.success(
                                "このルートの近くに登録済みの危険地点は見つかりませんでした。"
                            )
                        else:
                            st.warning(
                                f"このルートの周辺120m以内に危険地点が"
                                f"{route_danger_info['danger_count']}件あります。"
                                f"うち危険度4以上は"
                                f"{route_danger_info['high_danger_count']}件です。"
                            )

                            st.write("#### ルート周辺の危険地点")

                            for _, danger_row in route_danger_info["nearby_points"].head(5).iterrows():
                                st.write(
                                    f"⚠️ **{danger_row['place']}**｜"
                                    f"危険度{danger_row['danger_level']}｜"
                                    f"ルートから約{danger_row['route_distance_m']:.0f}m"
                                )

                            st.info(
                                "現在のStreamlit版では、危険地点を完全に避ける自動再計算ではなく、"
                                "危険地点が近いルートかどうかを判定して警告します。"
                                "本格実装では、危険地点を通らないように中継地点を自動生成します。"
                            )

                    elif selected_route_type == "上り坂少ないルート":
                        st.info(
                            "上り坂少ないルートが選択されています。"
                            "現在のコードでは標高データをまだ連携していないため、"
                            "通常ルートを表示しつつ、坂道考慮ルートの選択状態を表示しています。"
                            "今後、標高APIを接続すると上り坂の少ないルートを自動計算できます。"
                        )
                else:
                    st.warning(
                        "道路ルートを取得できなかったため、直線で表示します。"
                    )
                    st.info(
                        f"選択中のルート種類："
                        f"{st.session_state.get('selected_route_type', '最短ルート')}"
                    )
                    folium.PolyLine(
                        [
                            [start_location["latitude"], start_location["longitude"]],
                            [goal_location["latitude"], goal_location["longitude"]],
                        ],
                        weight=4,
                        dash_array="10",
                    ).add_to(route_map)

                st_folium(route_map, width=None, height=540, returned_objects=[], key="bike_route_map")
            else:
                st.info("出発地と目的地を入力してください。")

    with tabs[5]:
        st.header("お気に入り")

        favorite_df = load_favorites()

        if favorite_df.empty:
            st.info("お気に入りはまだありません。")
        else:
            for _, row in favorite_df.iterrows():
                left, right = st.columns([5, 1])

                with left:
                    display_point_card(row, show_favorite=False, key_prefix="bike_favorite")

                with right:
                    if st.button("削除", key=f"remove_favorite_{row['id']}", use_container_width=True):
                        remove_favorite(int(row["id"]))
                        st.success("削除しました。")
                        st.rerun()

    with tabs[6]:
        st.header("データ確認")

        danger_df = load_danger_points()

        if danger_df.empty:
            st.info("登録されているデータがありません。")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("公開中の危険地点", f"{len(danger_df)}件")
            c2.metric("平均危険度", f"{danger_df['danger_level'].mean():.1f}")
            c3.metric("危険度5", f"{len(danger_df[danger_df['danger_level'] == 5])}件")

            st.subheader("危険度別件数")
            level_count = danger_df.groupby("danger_level").size().reindex([1, 2, 3, 4, 5], fill_value=0)
            st.bar_chart(level_count)

            st.subheader("登録データ一覧")
            display_columns = ["id", "place", "danger_level", "category", "comment", "latitude", "longitude", "created_at"]
            st.dataframe(danger_df[display_columns], use_container_width=True, hide_index=True)

            csv_data = danger_df.to_csv(index=False).encode("utf-8-sig")

            st.download_button(
                "CSVデータをダウンロード",
                data=csv_data,
                file_name="rescue_ride_danger_points.csv",
                mime="text/csv",
                use_container_width=True,
            )

    with tabs[7]:
        st.header("管理者設定")

        admin_password = st.text_input("管理者パスワード", type="password")
        correct_password = "ama1234"

        if admin_password:
            if admin_password != correct_password:
                st.error("パスワードが違います。")
            else:
                st.success("管理者としてログインしました。")

                all_points_df = load_danger_points(public_only=False)

                if all_points_df.empty:
                    st.info("管理対象のデータはありません。")
                else:
                    for _, row in all_points_df.iterrows():
                        with st.container(border=True):
                            st.write(f"### ID {row['id']}｜{row['place']}")
                            st.write(f"危険度：{row['danger_level']}／5")
                            st.write(f"分類：{row['category']}")
                            st.write(f"説明：{row['comment']}")
                            st.write(f"公開状態：{row['status']}")

                            c1, c2 = st.columns(2)

                            with c1:
                                new_status = st.selectbox(
                                    "公開状態",
                                    ["公開", "非公開"],
                                    index=0 if row["status"] == "公開" else 1,
                                    key=f"bike_status_{row['id']}",
                                )

                                if st.button("状態を変更", key=f"bike_update_{row['id']}", use_container_width=True):
                                    update_point_status(int(row["id"]), new_status)
                                    st.success("変更しました。")
                                    st.rerun()

                            with c2:
                                if st.button("この地点を削除", key=f"bike_delete_{row['id']}", use_container_width=True):
                                    delete_danger_point(int(row["id"]))
                                    st.success("削除しました。")
                                    st.rerun()

    with tabs[8]:
        st.header("📱 スマホ共有")

        render_smartphone_qr_panel("bike_share_tab")

        st.divider()

        st.write("### ローカルでスマホ表示する時のURL例")

        st.code(
            "http://あなたのPCのIPアドレス:8501\n"
            "例：http://192.168.1.10:8501",
            language="text",
        )

        st.caption(
            "スマホとPCが同じWi-Fiに接続されている必要があります。"
            "Streamlit Cloudで公開した場合は、公開URLを入力してください。"
        )



# =========================================================
# 10. 自動車モード専用アプリ
# =========================================================

def render_car_app():
    render_header(
        "🚗 レスキューライド 自動車モード",
        "自動車運転者向け：危険接近警告・音声警告・自転車接近注意",
        "🚗 自動車モード",
    )

    status = get_car_mode_status()

    if status["level"] == "danger":
        hero_class = "car-danger"
        title = "🚨 危険接近中"
        message = status["message"]

        if status.get("nearest_distance") is not None:
            message += f"<br>最寄り危険地点まで約{status['nearest_distance']:.0f}m"

        speak_warning_once(
            status.get("voice_text", ""),
            f"car_danger_{int(status.get('nearest_distance') or 0)}",
        )

    elif status["level"] == "warning":
        hero_class = "car-warning"
        title = "⚠️ この先注意"
        message = status["message"]

        if status.get("nearest_distance") is not None:
            message += f"<br>最寄り危険地点まで約{status['nearest_distance']:.0f}m"

        speak_warning_once(
            status.get("voice_text", ""),
            f"car_warning_{int(status.get('nearest_distance') or 0)}",
        )

    elif status["level"] == "waiting":
        hero_class = "car-warning"
        title = "📍 GPS待機中"
        message = "GPS警告タブで現在地を取得してください"

    else:
        hero_class = "car-safe"
        title = "✅ 安全確認中"
        message = status["message"]

    st.markdown(
        f"""
        <div class="car-hero {hero_class}">
            <div class="car-title">{title}</div>
            <div class="car-message">{message}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        f"""
        <div class="car-emergency">
            🛡️ 車用セーフティ・オーラ：
            現在地から半径{st.session_state.car_warning_radius}mを警告範囲として、
            危険地点・公式事故地点・自転車接近注意地点を確認します。
        </div>
        """,
        unsafe_allow_html=True,
    )

    car_tabs = st.tabs(
        [
            "🚘 運転メイン",
            "🗺️ 危険マップ",
            "📍 GPS警告",
            "🚲 自転車接近",
            "⚙️ 車モード設定",
            "📱 スマホ共有",
        ]
    )

    with car_tabs[0]:
        st.header("🚘 運転メイン")

        c1, c2, c3, c4 = st.columns(4)

        nearest_display = (
            f"約{status['nearest_distance']:.0f}m"
            if status.get("nearest_distance") is not None
            else "未取得"
        )

        c1.metric("最寄り危険地点", nearest_display)
        c2.metric("強警告距離", f"{st.session_state.car_danger_radius}m")
        c3.metric("注意警告距離", f"{st.session_state.car_warning_radius}m")
        c4.metric("自転車接近警告", f"{st.session_state.car_bicycle_alert_radius}m")

        st.divider()

        f1, f2, f3 = st.columns(3)

        with f1:
            st.markdown(
                """
                <div class="car-card">
                    <div class="car-card-title">🔊 音声警告</div>
                    <div class="car-card-text">
                        危険地点に近づくと、日本語音声で注意を促します。
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with f2:
            st.markdown(
                """
                <div class="car-card">
                    <div class="car-card-title">🔴 赤色点滅警告</div>
                    <div class="car-card-text">
                        強警告距離に入ると、赤い点滅画面で危険を知らせます。
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        with f3:
            st.markdown(
                """
                <div class="car-card">
                    <div class="car-card-title">🚲 自転車接近注意</div>
                    <div class="car-card-text">
                        自転車の飛び出し・カットインが起きやすい地点を検知します。
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

        if st.button("🔊 音声警告テスト", use_container_width=True):
            speak_warning_once(
                "警告テストです。危険地点に注意してください。",
                f"test_{datetime.now().timestamp()}",
            )
            st.success("音声警告テストを実行しました。音が出ない場合は、ブラウザ画面を一度クリックしてください。")

    with car_tabs[1]:
        st.header("🗺️ 自動車モード用 危険マップ")

        current_location = get_current_location()

        center_lat = st.session_state.current_latitude if current_location else DEFAULT_LATITUDE
        center_lon = st.session_state.current_longitude if current_location else DEFAULT_LONGITUDE

        danger_df = load_danger_points()

        car_map = create_danger_map(
            danger_df=danger_df,
            center_lat=center_lat,
            center_lon=center_lon,
            zoom=15 if current_location else 12,
            current_location=current_location,
            safety_aura_enabled=current_location is not None,
            safety_aura_radius=st.session_state.car_warning_radius,
            safety_aura_min_level=3,
            accident_df=official_accident_df,
            show_official_accidents=True,
        )

        st_folium(car_map, width=None, height=620, returned_objects=[], key="car_map")

        st.caption("車モードでは、公式事故地点・ユーザー投稿危険地点・現在地周辺の警告範囲を表示します。")

    with car_tabs[2]:
        st.header("📍 GPS警告")

        left, right = st.columns([1, 1.2])

        with left:
            render_gps_input_block("car_gps")

        with right:
            updated_status = get_car_mode_status()

            st.subheader("判定結果")

            nearest_display = (
                f"約{updated_status['nearest_distance']:.0f}m"
                if updated_status.get("nearest_distance") is not None
                else "未取得"
            )

            st.metric("最寄り危険地点まで", nearest_display)

            if updated_status["level"] == "danger":
                st.error(updated_status["message"])
            elif updated_status["level"] == "warning":
                st.warning(updated_status["message"])
            elif updated_status["level"] == "waiting":
                st.info(updated_status["message"])
            else:
                st.success(updated_status["message"])

            if st.session_state.current_latitude is not None:
                st.write(f"緯度：{st.session_state.current_latitude:.7f}")
                st.write(f"経度：{st.session_state.current_longitude:.7f}")

    with car_tabs[3]:
        st.header("🚲 自転車接近通知警告")

        updated_status = get_car_mode_status()
        bicycle_alerts = updated_status.get("bicycle_alerts", [])

        st.markdown(
            """
            <div class="car-emergency">
                自転車の飛び出し・カットイン・見通しの悪い交差点など、
                車から見て危険になりやすい地点を検知します。
            </div>
            """,
            unsafe_allow_html=True,
        )

        if not bicycle_alerts:
            st.success("現在の警告範囲内に、自転車接近注意地点はありません。")
        else:
            st.error(f"警告範囲内に自転車接近注意地点が {len(bicycle_alerts)}件 あります。")

            for row in bicycle_alerts:
                st.markdown(
                    f"""
                    <div class="danger-card">
                        <h4>🚲 {row['place']}</h4>
                        <b>分類：</b>{row['category']}<br>
                        <b>距離：</b>約{float(row['distance_m']):.0f}m<br>
                        <b>内容：</b>{row['comment']}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        st.session_state.car_bicycle_alert_radius = st.slider(
            "自転車接近通知警告距離",
            min_value=50,
            max_value=500,
            value=int(st.session_state.car_bicycle_alert_radius),
            step=10,
            format="%dメートル",
        )

    with car_tabs[4]:
        st.header("⚙️ 車モード設定")

        left, right = st.columns(2)

        with left:
            st.session_state.car_voice_enabled = st.toggle(
                "音声警告を使う",
                value=st.session_state.car_voice_enabled,
            )

            st.session_state.car_red_flash_enabled = st.toggle(
                "赤色点滅警告を使う",
                value=st.session_state.car_red_flash_enabled,
            )

        with right:
            st.session_state.car_danger_radius = st.slider(
                "強警告距離",
                min_value=50,
                max_value=500,
                value=int(st.session_state.car_danger_radius),
                step=10,
                format="%dメートル",
            )

            st.session_state.car_warning_radius = st.slider(
                "注意警告距離",
                min_value=200,
                max_value=1000,
                value=int(st.session_state.car_warning_radius),
                step=50,
                format="%dメートル",
            )

        st.info(
            "Streamlit版では、音声はブラウザ読み上げです。スマホのロック画面通知・バックグラウンドGPS・本体振動はFlutter版で実装します。"
        )

    with car_tabs[5]:
        st.header("📱 スマホ共有")

        render_smartphone_qr_panel("car_share_tab")

        st.divider()

        st.write("### ローカルでスマホ表示する時のURL例")

        st.code(
            "http://あなたのPCのIPアドレス:8501\n"
            "例：http://192.168.1.10:8501",
            language="text",
        )

        st.caption(
            "スマホとPCが同じWi-Fiに接続されている必要があります。"
            "Streamlit Cloudで公開した場合は、公開URLを入力してください。"
        )



# =========================================================
# 11. ルーティング
# =========================================================

if not st.session_state.mode_selected or st.session_state.app_mode is None:
    render_first_mode_select_screen()
    st.stop()

if st.session_state.app_mode == "自転車モード":
    render_bicycle_app()
    st.stop()

if st.session_state.app_mode == "自動車モード":
    render_car_app()
    st.stop()

render_first_mode_select_screen()
