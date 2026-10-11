# -*- coding: utf-8 -*-
# 道頓堀 焼肉 新恵 DX注文管理ツール
# 実行: python -m streamlit run app.py

import streamlit as st
import pandas as pd
import plotly.express as px
import os
import uuid
import io
import json
import re
from datetime import date, datetime
import time

import dx_addon as dx   # 追加モジュール（コード表・精度ログ・一括読み取り）

# Google Sheets連携（gspreadがインストール済みの場合のみ有効）
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False

# google-generativeai はオプション（インストール済みの場合のみOCR機能が有効）
try:
    import google.generativeai as genai
    from PIL import Image
    GEMINI_AVAILABLE = True
except ImportError:
    GEMINI_AVAILABLE = False

# ============================================================
# 定数・設定
# ============================================================

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
CSV_FILE   = os.path.join(BASE_DIR, "yakiniku_orders.csv")
EXCEL_FILE = os.path.join(BASE_DIR, "yakiniku_orders.xlsx")

# Googleスプレッドシート設定
SPREADSHEET_ID  = "1FQJT7zJE3-sQOdchxcrohmI6BXUZ-rcyzuIUWAVzYFU"
SHEET_NAME      = "注文データ"
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "service_account.json")
GSPREAD_SCOPES  = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# データの正式な列順（これに厳密に合わせることで列のずれを防ぐ）
COLUMNS_ORDER = [
    "日付", "テーブル番号", "国籍", "人数", "メニュー名", "数量", "金額",
    "来店経路", "顧客グループID", "登録日時"
]

# Gemini モデル名（モデル更新時はここだけ変更すればよい）
GEMINI_MODEL = "gemini-3.5-flash-lite"

NATIONALITY_CODES = ["J", "C", "K", "A", "W", "O"]

NATIONALITY_LABELS = {
    "J": "日本",
    "C": "中国系",
    "K": "韓国",
    "A": "他アジア",
    "W": "欧米",
    "O": "その他・不明",
}

_LEGACY_NATIONALITY_TO_CODE = {
    "日本": "J",
    "中国": "C", "台湾": "C", "香港": "C", "マカオ": "C", "中国系": "C",
    "韓国": "K",
    "タイ": "A", "シンガポール": "A", "ベトナム": "A", "フィリピン": "A",
    "インドネシア": "A", "マレーシア": "A", "インド": "A", "他アジア": "A",
    "アメリカ": "W", "カナダ": "W", "イギリス": "W", "フランス": "W",
    "ドイツ": "W", "オーストラリア": "W", "ニュージーランド": "W",
    "スイス": "W", "イタリア": "W", "スペイン": "W", "オランダ": "W", "欧米": "W",
    "その他": "O", "不明": "O", "": "O", "その他・不明": "O",
}

def to_nationality_code(value):
    if value in NATIONALITY_LABELS:
        return value
    return _LEGACY_NATIONALITY_TO_CODE.get(str(value).strip(), "O")

def nationality_display(code):
    c = to_nationality_code(code)
    return f"{c} {NATIONALITY_LABELS[c]}"

VISIT_ROUTES = [
    "Google Maps", "食べログ", "TripAdvisor", "Instagram",
    "口コミ（知人紹介）", "ホテルフロント", "旅行会社", "その他"
]

MENU_PRESETS = {
    "🥩 熟成牛肉": [
        ("特選ロース",    3720),
        ("上ロース",      2520),
        ("特選カルビ",    2380),
        ("上カルビ",      1580),
        ("特選ハラミ",    2380),
        ("上ハラミ",      1580),
        ("特選フィレ",    3380),
    ],
    "🫀 ホルモン": [
        ("タン",          3000),
        ("うす切りタン",  1200),
        ("ツラミ",         880),
        ("ミノ",           680),
        ("センマイ",       680),
        ("アカセンマイ",   680),
        ("シマチョウ",     680),
        ("ショウチョウ",   680),
        ("レバー",         680),
    ],
    "🌟 セット": [
        ("神戸ビーフセット",           21000),
        ("神戸ビーフセット スペシャル", 41000),
        ("特選三種",                    9000),
        ("上物三種",                    6000),
    ],
    "🥬 野菜・ナムル": [
        ("チシャ菜",     380),
        ("キャベツ",     380),
        ("サラダ",       600),
        ("ナムル盛合せ", 600),
        ("モヤシナムル", 380),
        ("焼野菜",       600),
    ],
    "🥣 キムチ": [
        ("キムチ盛合せ", 700),
        ("白菜キムチ",   500),
        ("大根キムチ",   500),
        ("胡瓜キムチ",   500),
    ],
    "🍚 ライス・スープ": [
        ("ライス",       380),
        ("ビビンバ",     800),
        ("石焼ビビンバ", 950),
        ("クッパ",       700),
        ("ワカメスープ", 500),
        ("玉子スープ",   500),
        ("冷麺",         800),
    ],
    "🍺 ビール": [
        ("生ビール 380ml",  600),
        ("生ビール 大",     950),
    ],
    "🍋 サワー・チューハイ": [
        ("レモンサワー",          550),
        ("ライムサワー",          550),
        ("ピーチサワー",          550),
        ("グレープフルーツサワー", 550),
        ("カルピスサワー",         550),
    ],
    "🥃 ハイボール": [
        ("ブラックニッカクリア", 600),
        ("ジャックダニエル",     700),
        ("山崎ハイボール",      1400),
        ("白州ハイボール",      1400),
        ("知多ハイボール",      1000),
    ],
    "🍷 ワイン": [
        ("グラス(赤)",              800),
        ("カベルネ(赤)ハーフボトル", 2200),
    ],
    "🌸 梅酒": [
        ("梅酒(ロック/ソーダ/水割)", 600),
    ],
    "🥃 焼酎": [
        ("二階堂(麦)",          600),
        ("佐藤(麦)",            900),
        ("黒霧島(芋)",          600),
        ("ジンロマッコリ 375ml", 1000),
    ],
    "🍶 日本酒": [
        ("白鶴 一合 180ml",     600),
        ("白鶴 二合 360ml",     800),
        ("冷酒 300ml",          900),
        ("長兵衛 大吟醸 120ml", 1300),
    ],
    "🥤 ソフトドリンク": [
        ("ウーロン茶",      400),
        ("オレンジジュース", 400),
        ("カルピス",        400),
        ("コーラ",          400),
        ("ジンジャーエール", 400),
        ("緑茶",            400),
    ],
}

ALL_MENUS = {name: price for items in MENU_PRESETS.values() for name, price in items}

# ============================================================
# Googleスプレッドシート連携関数（列ズレ防止・完全一括上書き）
# ============================================================

def get_gspread_client():
    if not GSPREAD_AVAILABLE:
        return None
    try:
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"], scopes=GSPREAD_SCOPES
            )
        elif os.path.exists(SERVICE_ACCOUNT_FILE):
            creds = Credentials.from_service_account_file(
                SERVICE_ACCOUNT_FILE, scopes=GSPREAD_SCOPES
            )
        else:
            return None
        return gspread.authorize(creds)
    except Exception:
        return None

def save_to_spreadsheet(df):
    """常に定義されたCOLUMNS_ORDERの通りに並べ替えて全件上書きする"""
    client = get_gspread_client()
    if client is None:
        return False
    try:
        sh = client.open_by_key(SPREADSHEET_ID)
        try:
            ws = sh.worksheet(SHEET_NAME)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=SHEET_NAME, rows=5000, cols=20)
        
        ws.clear()
        
        for col in COLUMNS_ORDER:
            if col not in df.columns:
                df[col] = ""
        df_ordered = df[COLUMNS_ORDER].fillna("").astype(str)
        
        ws.update([df_ordered.columns.tolist()] + df_ordered.values.tolist(), value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        st.error(f"❌ スプレッドシート保存エラー: {e}")
        return False

# ============================================================
# データ管理関数
# ============================================================

@st.cache_data
def load_data():
    if os.path.exists(CSV_FILE):
        df = pd.read_csv(CSV_FILE, encoding="utf-8-sig")
        for col in COLUMNS_ORDER:
            if col not in df.columns:
                if col == "顧客グループID":
                    df[col] = [str(uuid.uuid4())[:8] for _ in range(len(df))]
                elif col == "テーブル番号":
                    df[col] = ""
                else:
                    df[col] = ""
        return df[COLUMNS_ORDER]
    return pd.DataFrame(columns=COLUMNS_ORDER)

def save_data(df):
    for col in COLUMNS_ORDER:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS_ORDER]

    df.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")
    try:
        with open(EXCEL_FILE, "wb") as f:
            f.write(generate_excel(df))
    except Exception as e:
        st.warning(f"⚠️ Excelファイルへの同期保存に失敗しました: {e}")
    
    success = save_to_spreadsheet(df)
    if not success:
        st.warning("⚠️ Googleスプレッドシートへの保存に失敗しました（CSVには保存済み）")
    
    load_data.clear()

def add_order(visit_date, table_num, nationality, num_people, menu_items, visit_route, group_id):
    df = load_data()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [
        {
            "日付":           str(visit_date),
            "テーブル番号":   table_num,
            "国籍":           nationality,
            "人数":           num_people,
            "メニュー名":     menu_name,
            "数量":           quantity,
            "金額":           price,
            "来店経路":       visit_route,
            "顧客グループID": group_id,
            "登録日時":       now_str,
        }
        for menu_name, quantity, price in menu_items
    ]
    save_data(pd.concat([df, pd.DataFrame(rows)], ignore_index=True))

def delete_group(group_id):
    df = load_data()
    save_data(df[df["顧客グループID"] != group_id])

# ============================================================
# 分析・OCR・UI部分
# ============================================================

@st.cache_data
def calc_sales_by_nationality(df):
    if df.empty:
        return pd.DataFrame()
    return df.groupby("国籍")["金額"].sum().reset_index().rename(columns={"金額": "売上合計(円)"})

@st.cache_data
def calc_popular_menu_by_nationality(df, nationality):
    if df.empty:
        return pd.DataFrame()
    filtered = df[df["国籍"] == nationality]
    if filtered.empty:
        return pd.DataFrame()
    return filtered.groupby("メニュー名").agg(
        注文数=("数量", "sum"),
        売上合計=("金額", "sum")
    ).reset_index().sort_values("注文数", ascending=False)

@st.cache_data
def calc_avg_spend_per_person(df):
    if df.empty:
        return pd.DataFrame()
    visit_summary = df.groupby(["顧客グループID", "国籍", "人数"]).agg(
        合計金額=("金額", "sum")
    ).reset_index()
    visit_summary["客単価(円/人)"] = visit_summary["合計金額"] / visit_summary["人数"]
    avg_spend = visit_summary.groupby("国籍")["客単価(円/人)"].mean().reset_index()
    avg_spend["客単価(円/人)"] = avg_spend["客単価(円/人)"].round(0).astype(int)
    return avg_spend.sort_values("客単価(円/人)", ascending=False)

def generate_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="注文データ", index=False)
        sales_df = calc_sales_by_nationality(df)
        if not sales_df.empty:
            sales_df.to_excel(writer, sheet_name="国籍別売上", index=False)
        avg_df = calc_avg_spend_per_person(df)
        if not avg_df.empty:
            avg_df.to_excel(writer, sheet_name="平均客単価", index=False)
    return output.getvalue()

def ocr_receipt_with_gemini(image_bytes, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(GEMINI_MODEL)
    img = Image.open(io.BytesIO(image_bytes))
    response = model.generate_content([dx.build_ocr_prompt(), img])
    return dx.parse_ocr_response(response.text)

def match_menu_to_presets(ocr_result):
    items = ocr_result.get("items", []) if isinstance(ocr_result, dict) else ocr_result
    matched, unmatched = [], []
    for item in items:
        code = item.get("code", "")
        qty  = int(item.get("quantity", 1) or 1)
        if code in dx.MENU_CODES:
            name = dx.name_of(code)
            matched.append({
                "menu_name":  name,
                "quantity":   qty,
                "unit_price": ALL_MENUS.get(name, dx.price_of(code)),
                "code":       code,
                "confidence": item.get("confidence", ""),
            })
        else:
            unmatched.append({
                "menu_name": f"[読取失敗 {code}] {item.get('note', '')}",
                "quantity":  qty,
            })
    return matched, unmatched

if "cart" not in st.session_state:
    st.session_state.cart = []

if "ocr_results" not in st.session_state:
    st.session_state.ocr_results = None

if "ocr_unmatched" not in st.session_state:
    st.session_state.ocr_unmatched = []

for _k, _v in [("ocr_raw", []), ("ocr_image", ""), ("ocr_elapsed", None),
               ("ocr_header", {})]:
    if _k not in st.session_state:
        st.session_state[_k] = _v

st.set_page_config(page_title="焼肉 新恵 DXツール", page_icon="🥩", layout="wide")

st.title("🥩 道頓堀 焼肉 新恵｜DX注文管理ツール")
st.caption("大学ゼミ研究用 | 外国人観光客向け注文データ分析システム")

st.sidebar.header("📝 注文入力")
st.sidebar.markdown("---")

st.sidebar.subheader("来店情報")
with st.sidebar.form("sidebar_form"):
    visit_date  = st.date_input("📅 来店日", value=date.today())
    beer_large_price = 800 if visit_date <= date(2026, 6, 1) else 950
    table_num   = st.text_input("🪑 テーブル番号", placeholder="例：3、A-2")
    nationality = st.selectbox("🌍 国籍", options=NATIONALITY_CODES, format_func=lambda c: nationality_display(c))
    num_people  = st.number_input("👥 人数（名）", min_value=1, max_value=20, value=2, step=1)
    visit_route = st.selectbox("📍 来店経路", options=VISIT_ROUTES)
    st.form_submit_button("反映")

st.sidebar.markdown("---")

st.sidebar.subheader("🤖 AI/OCR 伝票読み取り")
st.sidebar.caption("伝票を撮影してアップロード → AIが自動でカートに追加")

if not GEMINI_AVAILABLE:
    st.sidebar.warning("OCR機能を使うには以下をインストールしてください：\n```\npip install google-generativeai Pillow\n```")
else:
    gemini_api_key = st.sidebar.text_input("Gemini API Key", type="password", placeholder="AIza...")
    uploaded_img = st.sidebar.file_uploader("伝票画像（JPG / PNG）", type=["jpg", "jpeg", "png"])

    if uploaded_img and gemini_api_key:
        st.sidebar.image(uploaded_img, caption="アップロード画像", use_container_width=True)
        if st.sidebar.button("🔍 AIで伝票を読み取る", use_container_width=True, type="primary"):
            _t0 = time.time()
            with st.spinner("AI読み取り中..."):
                try:
                    image_bytes = uploaded_img.read()
                    ocr_result = ocr_receipt_with_gemini(image_bytes, gemini_api_key)
                    matched, unmatched = match_menu_to_presets(ocr_result)
                    st.session_state.ocr_results   = matched
                    st.session_state.ocr_unmatched = unmatched
                    st.session_state.ocr_raw       = (ocr_result.get("items", []) if isinstance(ocr_result, dict) else ocr_result)
                    st.session_state.ocr_header    = (ocr_result if isinstance(ocr_result, dict) else {})
                    st.session_state.ocr_image     = uploaded_img.name
                    st.session_state.ocr_elapsed   = time.time() - _t0
                    st.sidebar.success(f"✅ {len(matched)} 品目を読み取りました")
                except Exception as e:
                    st.sidebar.error(f"❌ エラー: {e}")

    if st.session_state.ocr_results is not None:
        st.sidebar.markdown("---")
        st.sidebar.markdown("#### 📋 読み取り結果（確認）")
        confirmed_items = []
        for idx, item in enumerate(st.session_state.ocr_results):
            col_name, col_qty = st.sidebar.columns([3, 1])
            col_name.write(f"**{item['menu_name']}**  ¥{item['unit_price']:,}")
            new_qty = col_qty.number_input("数量", min_value=0, max_value=99, value=item["quantity"], key=f"ocr_qty_{idx}", label_visibility="collapsed")
            if new_qty > 0:
                confirmed_items.append({**item, "quantity": new_qty})

        oc1, oc2 = st.sidebar.columns(2)
        if oc1.button("🛒 カートに追加", key="ocr_add_to_cart", type="primary", use_container_width=True):
            for item in confirmed_items:
                for c in st.session_state.cart:
                    if c["menu_name"] == item["menu_name"]:
                        c["quantity"] += item["quantity"]
                        break
                else:
                    st.session_state.cart.append(item)
            st.session_state.ocr_results   = None
            st.session_state.ocr_unmatched = []
            st.rerun()
        if oc2.button("✕ キャンセル", key="ocr_cancel", use_container_width=True):
            st.session_state.ocr_results   = None
            st.session_state.ocr_unmatched = []
            st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🍖 メニュー選択")
for category, items in MENU_PRESETS.items():
    with st.sidebar.expander(category, expanded=False):
        for name, price in items:
            actual_price = beer_large_price if name == "生ビール 大" else price
            if st.button(f"{name} ¥{actual_price:,}", key=f"btn_{name}", use_container_width=True):
                for item in st.session_state.cart:
                    if item["menu_name"] == name:
                        item["quantity"] += 1
                        break
                else:
                    st.session_state.cart.append({"menu_name": name, "quantity": 1, "unit_price": actual_price})
                st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🧾 カート")
if st.session_state.cart:
    total = sum(i["quantity"] * i["unit_price"] for i in st.session_state.cart)
    if st.sidebar.button("💾 このグループを保存", type="primary", use_container_width=True):
        group_id   = str(uuid.uuid4())[:8]
        menu_items = [(i["menu_name"], i["quantity"], i["quantity"] * i["unit_price"]) for i in st.session_state.cart]
        add_order(visit_date, table_num, nationality, num_people, menu_items, visit_route, group_id)
        st.session_state.cart = []
        st.sidebar.success(f"✅ 保存しました！（ID: {group_id}）")
        st.rerun()

df = load_data()
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(["📋 データ一覧", "👥 顧客グループ管理", "📊 国籍別売上", "🏆 人気メニュー", "💰 平均客単価", "📦 一括読み取り", "🎯 読み取り精度"])

with tab1:
    st.subheader("📋 保存済みデータ一覧")
    if not df.empty:
        st.dataframe(df, use_container_width=True)

with tab2:
    st.subheader("👥 顧客グループ管理")
    if not df.empty and "顧客グループID" in df.columns:
        for gid in df["顧客グループID"].unique():
            st.write(f"Group ID: {gid}")

with tab3:
    st.subheader("📊 国籍別売上集計")
    if not df.empty:
        sales_df = calc_sales_by_nationality(df)
        st.dataframe(sales_df)

with tab4:
    st.subheader("🏆 人気メニュー")

with tab5:
    st.subheader("💰 平均客単価")

with tab6:
    dx.render_batch_tab(st, pd, add_order, ALL_MENUS, NATIONALITY_CODES, VISIT_ROUTES, ocr_receipt_with_gemini, uuid, date, time)

with tab7:
    st.subheader("🎯 読み取り精度")
