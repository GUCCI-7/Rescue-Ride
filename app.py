import json
import math
import sqlite3
import time
import uuid
from datetime import datetime, timedelta
from pathlib import Path

import folium
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from streamlit_autorefresh import st_autorefresh
from streamlit_folium import st_folium
from streamlit_js_eval import streamlit_js_eval

APP_DIR = Path(__file__).resolve().parent
POLICE_XLSX = APP_DIR / "honhyo_2024.xlsx"
POLICE_CACHE = APP_DIR / "police_danger_points.csv"
REVIEWS_FILE = APP_DIR / "reviews.csv"
DB_FILE = APP_DIR / "rescue_ride.db"
RIDE_KUN_FILE = APP_DIR / "ride_kun.png"

DEFAULT_LAT = 35.681236
DEFAULT_LON = 139.767125
GPS_REFRESH_MS = 3500
LOCATION_EXPIRE_SECONDS = 15
ALERT_COOLDOWN_SECONDS = 45
MAP_SEARCH_RADIUS_KM = 3.0
MAP_MAX_POINTS = 450
LOCATION_WRITE_INTERVAL_SECONDS = 6
LOCATION_MIN_MOVE_METERS = 8

st.set_page_config(page_title="レスキューライド", page_icon="🚲", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
[data-testid="stSidebar"] {display:none;}
.block-container {max-width:1250px; padding-top:1rem; padding-bottom:2rem;}
.hero {background:linear-gradient(135deg,#075985,#0f766e);color:white;border-radius:25px;padding:24px 28px;margin-bottom:16px;box-shadow:0 8px 24px rgba(0,0,0,.15);}
.hero h1 {margin:0;font-size:2.35rem;}
.hero p {margin:.45rem 0 0;line-height:1.7;}
.mode-card {background:white;border:2px solid #e2e8f0;border-radius:22px;padding:20px;min-height:180px;box-shadow:0 4px 14px rgba(15,23,42,.07);}
.mode-selected {border:4px solid #0284c7;background:#eff6ff;}
.mode-title {font-size:1.55rem;font-weight:900;margin-bottom:8px;}
.metric {background:white;border:1px solid #e2e8f0;border-radius:18px;padding:14px;text-align:center;box-shadow:0 3px 10px rgba(15,23,42,.05);}
.metric strong {display:block;font-size:1.55rem;margin-top:4px;}
.small {font-size:.85rem;color:#64748b;}
.alert-danger {background:#fee2e2;color:#7f1d1d;border-left:10px solid #dc2626;border-radius:18px;padding:20px;font-size:1.15rem;font-weight:850;margin:12px 0;}
.alert-warning {background:#ffedd5;color:#7c2d12;border-left:10px solid #f97316;border-radius:18px;padding:20px;font-size:1.1rem;font-weight:850;margin:12px 0;}
.alert-safe {background:#dcfce7;color:#166534;border-left:10px solid #16a34a;border-radius:18px;padding:18px;font-weight:800;margin:12px 0;}
.flash-panel {animation:rescueFlash .65s infinite alternate;border-radius:22px;padding:26px;text-align:center;font-size:1.5rem;font-weight:950;margin:12px 0;border:5px solid #ef4444;}
@keyframes rescueFlash {from {background:#fff;color:#991b1b;box-shadow:0 0 8px #ef4444;} to {background:#ef4444;color:#fff;box-shadow:0 0 35px #ef4444;}}
.stButton button {min-height:48px;border-radius:14px;font-weight:850;}
</style>
""", unsafe_allow_html=True)


def haversine_m(lat1, lon1, lat2, lon2):
    r = 6_371_000
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = math.radians(lat2-lat1), math.radians(lon2-lon1)
    a = math.sin(dp/2)**2 + math.cos(p1)*math.cos(p2)*math.sin(dl/2)**2
    return r * 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))


def packed_dms_to_decimal(value, is_longitude=False):
    if pd.isna(value):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if not is_longitude and 20 <= number <= 46:
        return number
    if is_longitude and 122 <= number <= 154:
        return number
    digits = str(int(number))
    try:
        if is_longitude:
            digits = digits.zfill(10)
            deg, minute = int(digits[:3]), int(digits[3:5])
            sec = int(digits[5:7]) + int(digits[7:10]) / 1000
        else:
            digits = digits.zfill(9)
            deg, minute = int(digits[:2]), int(digits[2:4])
            sec = int(digits[4:6]) + int(digits[6:9]) / 1000
    except ValueError:
        return None
    if minute >= 60 or sec >= 60:
        return None
    return deg + minute/60 + sec/3600


def convert_police_data():
    columns = ["id","place_name","lat","lon","danger_level","category","comment","source"]
    empty = pd.DataFrame(columns=columns)
    if not POLICE_XLSX.exists():
        return empty
    wanted = ["本票番号","事故内容","死者数","負傷者数","道路形状","事故類型","地点　緯度（北緯）","地点　経度（東経）"]
    try:
        raw = pd.read_excel(POLICE_XLSX, usecols=lambda c: c in wanted, engine="openpyxl")
    except Exception as exc:
        st.error(f"警察庁オープンデータを読み込めませんでした：{exc}")
        return empty
    lat_col = next((c for c in raw.columns if "緯度" in str(c)), None)
    lon_col = next((c for c in raw.columns if "経度" in str(c)), None)
    if lat_col is None or lon_col is None:
        st.error("警察庁データに緯度・経度列がありません。")
        return empty
    raw["lat"] = raw[lat_col].apply(lambda v: packed_dms_to_decimal(v, False))
    raw["lon"] = raw[lon_col].apply(lambda v: packed_dms_to_decimal(v, True))
    raw = raw.dropna(subset=["lat","lon"])
    raw = raw[raw["lat"].between(20,46) & raw["lon"].between(122,154)].copy()
    deaths = pd.to_numeric(raw.get("死者数",0), errors="coerce").fillna(0)
    injuries = pd.to_numeric(raw.get("負傷者数",0), errors="coerce").fillna(0)
    raw["danger_level"] = 3
    raw.loc[injuries >= 2, "danger_level"] = 4
    raw.loc[deaths >= 1, "danger_level"] = 5
    ticket = raw.get("本票番号", pd.Series(range(len(raw)), index=raw.index)).astype(str)
    accident = raw.get("事故内容", pd.Series("", index=raw.index)).astype(str)
    road = raw.get("道路形状", pd.Series("", index=raw.index)).astype(str)
    accident_type = raw.get("事故類型", pd.Series("", index=raw.index)).astype(str)
    converted = pd.DataFrame({
        "id":"police_"+ticket,
        "place_name":"交通事故発生地点",
        "lat":raw["lat"],
        "lon":raw["lon"],
        "danger_level":raw["danger_level"],
        "category":"警察庁交通事故データ",
        "comment":"事故内容コード："+accident+"／道路形状コード："+road+"／事故類型コード："+accident_type,
        "source":"警察庁オープンデータ（2024年）",
    }).drop_duplicates(subset=["id","lat","lon"])
    converted.to_csv(POLICE_CACHE, index=False)
    return converted.reset_index(drop=True)


@st.cache_data(show_spinner="警察庁オープンデータを準備しています…")
def load_police_points():
    if POLICE_CACHE.exists():
        try:
            df = pd.read_csv(POLICE_CACHE)
        except Exception:
            df = pd.DataFrame()
    else:
        df = convert_police_data()
    required = ["id","place_name","lat","lon","danger_level","category","comment","source"]
    for c in required:
        if c not in df.columns:
            df[c] = ""
    if len(df):
        df["lat"] = pd.to_numeric(df["lat"], errors="coerce")
        df["lon"] = pd.to_numeric(df["lon"], errors="coerce")
        df["danger_level"] = pd.to_numeric(df["danger_level"], errors="coerce").fillna(3).clip(1,5)
        df = df.dropna(subset=["lat","lon"])
    return df.reset_index(drop=True)


REVIEW_COLUMNS = ["datetime","mode","place_name","lat","lon","danger_level","category","comment"]

@st.cache_data(ttl=5, show_spinner=False)
def load_reviews():
    if not REVIEWS_FILE.exists():
        pd.DataFrame(columns=REVIEW_COLUMNS).to_csv(REVIEWS_FILE, index=False)
    try:
        df = pd.read_csv(REVIEWS_FILE)
    except Exception:
        df = pd.DataFrame(columns=REVIEW_COLUMNS)
    for c in REVIEW_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    return df


def save_review(mode, place_name, lat, lon, danger_level, category, comment):
    df = load_reviews()
    new = pd.DataFrame([{
        "datetime":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "mode":mode,
        "place_name":place_name,
        "lat":lat,
        "lon":lon,
        "danger_level":danger_level,
        "category":category,
        "comment":comment,
    }])
    pd.concat([df,new], ignore_index=True).to_csv(REVIEWS_FILE, index=False)


def reviews_as_points(reviews):
    if len(reviews)==0:
        return pd.DataFrame(columns=["id","place_name","lat","lon","danger_level","category","comment","source"])
    points = reviews[["place_name","lat","lon","danger_level","category","comment"]].copy()
    points["lat"] = pd.to_numeric(points["lat"], errors="coerce")
    points["lon"] = pd.to_numeric(points["lon"], errors="coerce")
    points["danger_level"] = pd.to_numeric(points["danger_level"], errors="coerce").fillna(3).clip(1,5)
    points["id"] = [f"review_{i}" for i in range(len(points))]
    points["source"] = "利用者の口コミ"
    return points.dropna(subset=["lat","lon"])


def init_database():
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS live_locations (
            anonymous_id TEXT PRIMARY KEY,
            mode TEXT NOT NULL,
            lat REAL NOT NULL,
            lon REAL NOT NULL,
            accuracy REAL,
            updated_at TEXT NOT NULL
        )
        """)
        conn.commit()


def update_live_location(anonymous_id, mode, lat, lon, accuracy):
    now = datetime.now().isoformat(timespec="seconds")
    cutoff = (datetime.now()-timedelta(seconds=LOCATION_EXPIRE_SECONDS)).isoformat(timespec="seconds")
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute("""
        INSERT INTO live_locations (anonymous_id,mode,lat,lon,accuracy,updated_at)
        VALUES (?,?,?,?,?,?)
        ON CONFLICT(anonymous_id) DO UPDATE SET
            mode=excluded.mode, lat=excluded.lat, lon=excluded.lon,
            accuracy=excluded.accuracy, updated_at=excluded.updated_at
        """, (anonymous_id,mode,lat,lon,accuracy,now))
        conn.execute("DELETE FROM live_locations WHERE updated_at < ?", (cutoff,))
        conn.commit()


def remove_live_location(anonymous_id):
    with sqlite3.connect(DB_FILE) as conn:
        conn.execute("DELETE FROM live_locations WHERE anonymous_id = ?", (anonymous_id,))
        conn.commit()


def nearby_opposite_users(anonymous_id, mode, lat, lon, radius):
    opposite = "自動車モード" if mode == "自転車モード" else "自転車モード"
    with sqlite3.connect(DB_FILE) as conn:
        rows = conn.execute("SELECT anonymous_id,mode,lat,lon,accuracy,updated_at FROM live_locations WHERE anonymous_id != ? AND mode = ?", (anonymous_id,opposite)).fetchall()
    items = []
    for row in rows:
        distance = haversine_m(lat,lon,float(row[2]),float(row[3]))
        if distance <= radius:
            items.append({"anonymous_id":row[0],"mode":row[1],"lat":float(row[2]),"lon":float(row[3]),"accuracy":row[4],"updated_at":row[5],"distance_m":distance})
    return sorted(items, key=lambda x:x["distance_m"])


def browser_location(refresh_count):
    return streamlit_js_eval(js_expressions="""
    new Promise((resolve) => {
      if (!navigator.geolocation) {resolve({error:"このブラウザは位置情報に対応していません。"});return;}
      const isLocal=["localhost","127.0.0.1","::1"].includes(window.location.hostname);
      if (!window.isSecureContext && !isLocal) {resolve({error:"GPSはHTTPSまたはlocalhostで利用してください。"});return;}
      navigator.geolocation.getCurrentPosition(
        (p)=>resolve({latitude:p.coords.latitude,longitude:p.coords.longitude,accuracy:p.coords.accuracy,speed:p.coords.speed,heading:p.coords.heading,timestamp:p.timestamp}),
        (e)=>{const m={1:"位置情報が拒否されています。ブラウザ設定で許可してください。",2:"現在地を特定できません。屋外や窓際で再試行してください。",3:"GPS取得がタイムアウトしました。"};resolve({error:m[e.code]||"GPSを取得できませんでした。"});},
        {enableHighAccuracy:true,timeout:9000,maximumAge:0}
      );
    });
    """, key=f"gps_{refresh_count}")


def enable_alerts():
    components.html("""
    <script>
    (async()=>{try{window.speechSynthesis.cancel();const u=new SpeechSynthesisUtterance("レスキューライドの通知を有効にしました。");u.lang="ja-JP";window.speechSynthesis.speak(u);if("Notification" in window&&Notification.permission==="default"){await Notification.requestPermission();}if(navigator.vibrate)navigator.vibrate([150,80,150]);}catch(e){console.log(e);}})();
    </script>
    """, height=0)


def send_alert(title, message, urgent=False):
    t=json.dumps(title,ensure_ascii=False); m=json.dumps(message,ensure_ascii=False)
    vibration="[500,180,500,180,700]" if urgent else "[250,120,250]"
    components.html(f"""
    <script>
    (()=>{{const title={t};const message={m};try{{window.speechSynthesis.cancel();const u=new SpeechSynthesisUtterance(message);u.lang='ja-JP';u.rate=.95;u.volume=1;window.speechSynthesis.speak(u);}}catch(e){{}}try{{if(navigator.vibrate)navigator.vibrate({vibration});}}catch(e){{}}try{{if('Notification' in window&&Notification.permission==='granted'){{const n=new Notification(title,{{body:message,tag:'rescue-ride-alert',renotify:true,requireInteraction:{str(True).lower() if urgent else str(False).lower()}}});setTimeout(()=>n.close(),10000);}}}}catch(e){{}}}})();
    </script>
    """, height=0)


def danger_message(distance):
    if distance <= 20:
        return "緊急", "危険地点です。すぐに速度を落とし、周囲を十分確認してください。"
    if distance <= 50:
        return "警告", "まもなく危険地点です。速度を落として注意してください。"
    return "注意", "前方に危険地点があります。周囲を確認してください。"


def create_map(lat,lon,accuracy,mode,danger_points,nearby_users,danger_radius):
    m=folium.Map(location=[lat,lon],zoom_start=17,tiles="OpenStreetMap")
    color="blue" if mode=="自転車モード" else "green"; icon="bicycle" if mode=="自転車モード" else "car"
    folium.Marker([lat,lon],tooltip=f"現在地：{mode}",icon=folium.Icon(color=color,icon=icon,prefix="fa")).add_to(m)
    folium.Circle([lat,lon],radius=max(float(accuracy or 20),5),color="#0284c7",fill=True,fill_opacity=.08,tooltip=f"GPS推定誤差：約{round(float(accuracy or 0))}m").add_to(m)
    folium.Circle([lat,lon],radius=danger_radius,color="#f59e0b",fill=False,dash_array="7",tooltip=f"危険地点通知範囲：{danger_radius}m").add_to(m)
    for _,row in danger_points.iterrows():
        level=int(row["danger_level"]); c="red" if level>=5 else "orange" if level>=3 else "green"
        popup=f"<b>{row['place_name']}</b><br>危険度：{level}<br>種類：{row['category']}<br>情報源：{row.get('source','不明')}<br>{row['comment']}"
        folium.CircleMarker([float(row["lat"]),float(row["lon"])],radius=6+level,color=c,fill=True,fill_opacity=.75,tooltip=str(row["place_name"]),popup=popup).add_to(m)
    for user in nearby_users:
        i="car" if user["mode"]=="自動車モード" else "bicycle"; c="green" if user["mode"]=="自動車モード" else "blue"
        folium.Marker([user["lat"],user["lon"]],tooltip=f"{user['mode']}・約{user['distance_m']:.0f}m",icon=folium.Icon(color=c,icon=i,prefix="fa")).add_to(m)
    return m


init_database()
defaults={"mode":None,"monitoring":False,"alerts_enabled":False,"anonymous_id":str(uuid.uuid4()),"lat":DEFAULT_LAT,"lon":DEFAULT_LON,"accuracy":None,"speed":None,"heading":None,"gps_error":None,"last_danger_alert_key":None,"last_danger_alert_time":0.0,"last_user_alert_key":None,"last_user_alert_time":0.0,"alert_history":[],"last_location_write_time":0.0,"last_written_lat":None,"last_written_lon":None}
for k,v in defaults.items(): st.session_state.setdefault(k,v)

hl,hr=st.columns([5,1])
with hl:
    st.markdown("""<div class='hero'><h1>🚲 レスキューライド</h1><p>警察庁の事故データ、利用者の口コミ、匿名の位置情報を組み合わせ、自転車と自動車の双方へ危険を事前に知らせる交通安全支援アプリです。</p></div>""", unsafe_allow_html=True)
with hr:
    if RIDE_KUN_FILE.exists(): st.image(str(RIDE_KUN_FILE),use_container_width=True)

st.warning("走行中はスマホを安全なホルダーに固定し、画面を操作・注視しないでください。本アプリは安全運転を補助する試作版で、事故防止を保証するものではありません。")

st.subheader("① 利用モードを選択")
ml,mr=st.columns(2)
with ml:
    sel=st.session_state.mode=="自転車モード"
    st.markdown(f"<div class='mode-card {'mode-selected' if sel else ''}'><div class='mode-title'>🚲 自転車モード</div>危険地点に近づくと、画面を点滅させて注意を促します。<br>車の接近、見通しの悪さ、狭い歩道、危険な交差点などを投稿できます。</div>",unsafe_allow_html=True)
    if st.button("🚲 自転車モードを選択",use_container_width=True): st.session_state.mode="自転車モード"; st.rerun()
with mr:
    sel=st.session_state.mode=="自動車モード"
    st.markdown(f"<div class='mode-card {'mode-selected' if sel else ''}'><div class='mode-title'>🚗 自動車モード</div>危険地点に近づくと、画面と音声で警告します。<br>自転車の飛び出し、車道への急な進入、急ブレーキなどを投稿できます。</div>",unsafe_allow_html=True)
    if st.button("🚗 自動車モードを選択",use_container_width=True): st.session_state.mode="自動車モード"; st.rerun()
if st.session_state.mode is None: st.info("上のどちらかのモードを選択してください。"); st.stop()
mode=st.session_state.mode

st.subheader("② GPS・通知設定")
c1,c2,c3,c4,c5=st.columns([1.1,1.1,1.1,1.4,1.4])
with c1:
    if not st.session_state.monitoring:
        if st.button("▶ 走行モニター開始",type="primary",use_container_width=True): st.session_state.monitoring=True; st.session_state.last_danger_alert_key=None; st.session_state.last_user_alert_key=None; st.rerun()
    else:
        if st.button("■ モニター停止",use_container_width=True): st.session_state.monitoring=False; remove_live_location(st.session_state.anonymous_id); st.rerun()
with c2:
    label="✅ 通知有効" if st.session_state.alerts_enabled else "🔔 音声・通知を有効化"
    if st.button(label,disabled=st.session_state.alerts_enabled,use_container_width=True): st.session_state.alerts_enabled=True; enable_alerts()
with c3:
    if st.button("📍 GPS再取得",use_container_width=True): st.session_state.gps_error=None; st.rerun()
with c4:
    danger_radius=st.select_slider("危険地点の通知距離",options=[30,50,75,100,150,200,300],value=100)
with c5:
    user_radius=st.select_slider("車・自転車の接近通知距離",options=[20,30,50,75,100,150,200],value=50)

refresh_count=st_autorefresh(interval=GPS_REFRESH_MS,limit=None,key="gps_refresh") if st.session_state.monitoring else int(time.time()*1000)
gps=browser_location(refresh_count)
if isinstance(gps,dict):
    if gps.get("error"): st.session_state.gps_error=gps["error"]
    elif gps.get("latitude") is not None and gps.get("longitude") is not None:
        st.session_state.lat=float(gps["latitude"]); st.session_state.lon=float(gps["longitude"])
        st.session_state.accuracy=float(gps["accuracy"]) if gps.get("accuracy") is not None else None
        st.session_state.speed=gps.get("speed"); st.session_state.heading=gps.get("heading"); st.session_state.gps_error=None
lat,lon,accuracy=st.session_state.lat,st.session_state.lon,st.session_state.accuracy
if st.session_state.gps_error: st.error(st.session_state.gps_error)
if st.session_state.monitoring and accuracy is not None:
    now_location=time.time()
    last_lat,last_lon=st.session_state.last_written_lat,st.session_state.last_written_lon
    moved=None if last_lat is None or last_lon is None else haversine_m(last_lat,last_lon,lat,lon)
    if last_lat is None or last_lon is None or moved is None or moved>=LOCATION_MIN_MOVE_METERS or now_location-st.session_state.last_location_write_time>=LOCATION_WRITE_INTERVAL_SECONDS:
        update_live_location(st.session_state.anonymous_id,mode,lat,lon,accuracy)
        st.session_state.last_location_write_time=now_location
        st.session_state.last_written_lat=lat
        st.session_state.last_written_lon=lon

police_points=load_police_points(); reviews=load_reviews(); review_points=reviews_as_points(reviews); all_points=pd.concat([police_points,review_points],ignore_index=True)
nearby_dangers=[]
prefilter_m=max(danger_radius,600)
lat_delta=prefilter_m/111000
lon_delta=prefilter_m/(111000*max(math.cos(math.radians(lat)),0.2))
alert_candidates=all_points[
    all_points["lat"].between(lat-lat_delta,lat+lat_delta)
    & all_points["lon"].between(lon-lon_delta,lon+lon_delta)
]
for idx,row in alert_candidates.iterrows():
    d=haversine_m(lat,lon,float(row["lat"]),float(row["lon"]))
    if d<=danger_radius: nearby_dangers.append((idx,d,row))
nearby_dangers.sort(key=lambda x:x[1])
nearby_users=nearby_opposite_users(st.session_state.anonymous_id,mode,lat,lon,user_radius)

speed_kmh=None
if st.session_state.speed is not None:
    try: speed_kmh=max(0.0,float(st.session_state.speed)*3.6)
    except: pass
cols=st.columns(5)
metrics=[("現在のモード",mode.replace("モード","")),("GPS精度","未取得" if accuracy is None else f"約{accuracy:.0f}m"),("速度","不明" if speed_kmh is None else f"{speed_kmh:.1f}km/h"),("近くの危険地点",f"{len(nearby_dangers)}件"),("近くの相手",f"{len(nearby_users)}人")]
for col,(label,value) in zip(cols,metrics):
    with col: st.markdown(f"<div class='metric'><span class='small'>{label}</span><strong>{value}</strong></div>",unsafe_allow_html=True)
if len(police_points): st.success(f"警察庁オープンデータと連携中：{len(police_points):,}地点")
elif not POLICE_XLSX.exists(): st.warning("honhyo_2024.xlsxが見つかりません。このPythonファイルと同じフォルダに置いてください。")

if nearby_dangers:
    idx,distance,danger=nearby_dangers[0]; band,voice=danger_message(distance)
    if mode=="自転車モード":
        st.markdown(f"<div class='flash-panel'>⚠️ {band}<br>{danger['place_name']}まで約{distance:.0f}m<br>速度を落とし、周囲を確認してください</div>",unsafe_allow_html=True)
    else:
        css="alert-danger" if distance<=50 else "alert-warning"
        st.markdown(f"<div class='{css}'>⚠️ {band}：{danger['place_name']}まで約{distance:.0f}m<br>種類：{danger['category']}／危険度：{int(danger['danger_level'])}<br>{danger['comment']}</div>",unsafe_allow_html=True)
    danger_id=danger.get("id",idx); alert_key=f"{danger_id}:{band}:{mode}"; now=time.time()
    if st.session_state.monitoring and st.session_state.alerts_enabled and (st.session_state.last_danger_alert_key!=alert_key or now-st.session_state.last_danger_alert_time>=ALERT_COOLDOWN_SECONDS):
        text=f"{voice}{danger['place_name']}まで約{distance:.0f}メートルです。"
        if mode=="自動車モード": text += "自転車の飛び出しや死角に注意してください。"
        send_alert(f"危険地点まで約{distance:.0f}m",text,urgent=distance<=20)
        st.session_state.last_danger_alert_key=alert_key; st.session_state.last_danger_alert_time=now
        st.session_state.alert_history.insert(0,{"時刻":datetime.now().strftime("%H:%M:%S"),"種類":"危険地点","モード":mode,"内容":danger["place_name"],"距離":f"{distance:.0f}m"}); st.session_state.alert_history=st.session_state.alert_history[:30]
else:
    st.markdown("<div class='alert-safe'>✅ 現在、設定範囲内に危険地点はありません。</div>",unsafe_allow_html=True); st.session_state.last_danger_alert_key=None

if nearby_users:
    nearest=nearby_users[0]; distance=nearest["distance_m"]
    if mode=="自転車モード": title="🚗 車の接近"; message=f"近くに車が接近しています。距離は約{distance:.0f}メートルです。周囲を確認してください。"; target="車"
    else: title="🚲 自転車の接近"; message=f"近くに自転車が接近しています。距離は約{distance:.0f}メートルです。速度を落とし、飛び出しに注意してください。"; target="自転車"
    st.markdown(f"<div class='alert-warning'>{title}<br>近くに{target}ユーザーがいます。距離：約{distance:.0f}m</div>",unsafe_allow_html=True)
    band="緊急" if distance<=20 else "接近"; key=f"{nearest['anonymous_id']}:{band}:{mode}"; now=time.time()
    if st.session_state.monitoring and st.session_state.alerts_enabled and (st.session_state.last_user_alert_key!=key or now-st.session_state.last_user_alert_time>=ALERT_COOLDOWN_SECONDS):
        send_alert(title,message,urgent=distance<=20); st.session_state.last_user_alert_key=key; st.session_state.last_user_alert_time=now
        st.session_state.alert_history.insert(0,{"時刻":datetime.now().strftime("%H:%M:%S"),"種類":"利用者接近","モード":mode,"内容":f"{target}が接近","距離":f"{distance:.0f}m"}); st.session_state.alert_history=st.session_state.alert_history[:30]
else: st.session_state.last_user_alert_key=None

tabs=st.tabs(["🗺️ 安全マップ","⚠️ 近くの危険地点","💬 口コミ投稿","📝 口コミ一覧","🔔 通知履歴","📱 使い方"])
tab_map,tab_dangers,tab_post,tab_reviews,tab_history,tab_help=tabs
with tab_map:
    map_m=MAP_SEARCH_RADIUS_KM*1000
    map_lat_delta=map_m/111000
    map_lon_delta=map_m/(111000*max(math.cos(math.radians(lat)),0.2))
    display_points=all_points[
        all_points["lat"].between(lat-map_lat_delta,lat+map_lat_delta)
        & all_points["lon"].between(lon-map_lon_delta,lon+map_lon_delta)
    ].copy()
    if len(display_points)>MAP_MAX_POINTS:
        display_points["distance"]=[
            haversine_m(lat,lon,float(a),float(b))
            for a,b in zip(display_points["lat"],display_points["lon"])
        ]
        display_points=display_points.nsmallest(MAP_MAX_POINTS,"distance")
    st.caption(f"地図には現在地から約{MAP_SEARCH_RADIUS_KM:.0f}km以内を最大{MAP_MAX_POINTS}件表示します。")
    st_folium(create_map(lat,lon,accuracy,mode,display_points,nearby_users,danger_radius),height=520,use_container_width=True,returned_objects=[])
with tab_dangers:
    if not nearby_dangers: st.info("通知範囲内の危険地点はありません。")
    else:
        rows=[{"距離":f"{d:.0f}m","場所":r["place_name"],"危険度":int(r["danger_level"]),"種類":r["category"],"情報源":r.get("source","不明"),"内容":r["comment"]} for _,d,r in nearby_dangers[:50]]
        st.dataframe(pd.DataFrame(rows),use_container_width=True,hide_index=True)
with tab_post:
    if mode=="自転車モード": categories=["車の接近が怖い","見通しが悪い","歩道が狭い","交差点が危ない","路面状態が悪い","その他"]; st.info("自転車利用者の実感に基づく危険情報を投稿できます。")
    else: categories=["自転車の飛び出し","歩道から車道への急な進入","急ブレーキが必要だった","通学路で自転車が多い","路上駐車で死角がある","歩道がなく車道へ出る地点","その他"]; st.info("ドライバーが経験したヒヤリハット情報を投稿できます。")
    with st.form("review_form"):
        place=st.text_input("場所名",placeholder="例：〇〇交差点、〇〇通り"); category=st.selectbox("危険の種類",categories); level=st.slider("危険度",1,5,3); comment=st.text_area("詳しい内容",placeholder="いつ、どのような危険があったか入力してください。"); submitted=st.form_submit_button("現在地を危険地点として投稿")
        if submitted:
            if accuracy is None: st.error("先にGPSを取得してください。")
            elif not place.strip(): st.error("場所名を入力してください。")
            elif not comment.strip(): st.error("詳しい内容を入力してください。")
            else: save_review(mode,place.strip(),lat,lon,level,category,comment.strip()); st.success("口コミを投稿しました。"); st.rerun()
with tab_reviews:
    all_reviews=load_reviews()
    if len(all_reviews)==0: st.info("口コミはまだありません。")
    else:
        filter_mode=st.selectbox("表示する口コミ",["すべて","自転車モード","自動車モード"])
        shown=all_reviews if filter_mode=="すべて" else all_reviews[all_reviews["mode"]==filter_mode]
        st.dataframe(shown.sort_values("datetime",ascending=False).rename(columns={"datetime":"投稿日時","mode":"モード","place_name":"場所","lat":"緯度","lon":"経度","danger_level":"危険度","category":"種類","comment":"内容"}),use_container_width=True,hide_index=True)
with tab_history:
    if st.session_state.alert_history:
        st.dataframe(pd.DataFrame(st.session_state.alert_history),use_container_width=True,hide_index=True)
        if st.button("通知履歴を削除"): st.session_state.alert_history=[]; st.rerun()
    else: st.info("通知履歴はまだありません。")
with tab_help:
    st.markdown("""
### 基本的な使い方
1. 自転車モードまたは自動車モードを選択します。
2. 「音声・通知を有効化」を一度押します。
3. 「走行モニター開始」を押します。
4. ブラウザの位置情報を許可します。
5. スマホを安全なホルダーに固定して使用します。

### 自転車モード
- 危険地点に接近すると、画面が赤く点滅します。
- 車ユーザーが接近すると、画面・音声・振動で知らせます。
- 自転車利用者目線の危険な場所を投稿できます。

### 自動車モード
- 危険地点に接近すると、画面と音声で警告します。
- 自転車ユーザーが接近すると、画面・音声・振動で知らせます。
- 自転車の飛び出しや急ブレーキなどを投稿できます。

### 警察庁オープンデータ
- `honhyo_2024.xlsx`をこのファイルと同じフォルダに置きます。
- 初回起動時に`police_danger_points.csv`へ自動変換します。
- 警察の事故地点と利用者の口コミ地点を一緒に表示します。

### プライバシー
- 位置共有にはランダムに生成した匿名IDのみを使います。
- 氏名、メールアドレス、電話番号は保存しません。
- 停止後または一定時間更新がない位置情報は削除されます。

### 試作版における制限
- SQLiteによる位置共有は、同じサーバー上で動く試作版です。
- Streamlit Cloudではファイルが再起動時に消える場合があります。
- 本格運用にはSupabaseやFirebaseなどのリアルタイムデータベースが必要です。
- 画面ロック中や別アプリ表示中のGPS監視は保証されません。
- 「ライト点滅」はスマホ画面上の点滅です。端末の懐中電灯を直接点滅させる機能ではありません。
""")

st.caption("位置情報は匿名IDとともに一時的に利用します。正式運用には利用規約、プライバシーポリシー、認証、セキュリティ対策、安全性検証が必要です。")