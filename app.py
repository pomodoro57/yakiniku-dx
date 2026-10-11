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
# 分析関数
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

# ============================================================
# AI/OCR 伝票読み取り関数
# ============================================================

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

# ============================================================
# セッション初期化
# ============================================================

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

# ============================================================
# ページ設定
# ============================================================

st.set_page_config(page_title="焼肉 新恵 DXツール", page_icon="🥩", layout="wide")

st.title("🥩 道頓堀 焼肉 新恵｜DX注文管理ツール")
st.caption("大学ゼミ研究用 | 外国人観光客向け注文データ分析システム")

# ============================================================
# サイドバー：注文入力
# ============================================================

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

# ============================================================
# サイドバー：AI/OCR 伝票読み取り
# ============================================================

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

# ============================================================
# サイドバー：メニュー選択（プリセット）
# ============================================================

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

# ============================================================
# サイドバー：カート
# ============================================================

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

# ============================================================
# メインエリア
# ============================================================

df = load_data()
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📋 データ一覧",
    "👥 顧客グループ管理",
    "📊 国籍別売上",
    "🏆 人気メニュー",
    "💰 平均客単価",
    "📦 一括読み取り",
    "🎯 読み取り精度"
])

# タブ1：データ一覧
with tab1:
    st.subheader("📋 保存済みデータ一覧")
    if df.empty:
        st.info("👈 左のメニューからデータを入力して保存してください")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("総レコード数", f"{len(df)} 件")
        c2.metric("来店区分数",   f"{df['国籍'].apply(to_nationality_code).nunique()} 区分")
        c3.metric("総売上",       f"¥{df['金額'].sum():,}")
        st.markdown("---")
        filt = st.multiselect("🔍 国籍フィルター", options=df["国籍"].unique(), default=[])
        disp = df[df["国籍"].isin(filt)] if filt else df
        st.dataframe(
            disp.sort_values("登録日時", ascending=False).reset_index(drop=True),
            use_container_width=True, height=400
        )
        dl_col1, dl_col2 = st.columns(2)
        with dl_col1:
            st.download_button(
                "📥 CSVダウンロード",
                data=disp.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig"),
                file_name=f"yakiniku_orders_{date.today()}.csv",
                mime="text/csv",
                use_container_width=True
            )
        with dl_col2:
            excel_data = generate_excel(df)
            st.download_button(
                "📊 Excelダウンロード（3シート）",
                data=excel_data,
                file_name=f"yakiniku_orders_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# タブ2：顧客グループ管理
with tab2:
    st.subheader("👥 顧客グループ管理")
    st.caption("来店グループごとに注文内容を確認・編集・削除できます")

    if "editing_group" not in st.session_state:
        st.session_state.editing_group = None

    if df.empty:
        st.info("データがありません。")
    else:
        group_summary = df.groupby("顧客グループID").agg(
            日付=("日付", "first"),
            テーブル番号=("テーブル番号", "first"),
            国籍=("国籍", "first"),
            人数=("人数", "first"),
            来店経路=("来店経路", "first"),
            合計金額=("金額", "sum"),
            登録日時=("登録日時", "first"),
        ).reset_index().sort_values("登録日時", ascending=False)

        st.markdown(f"**登録グループ数: {len(group_summary)} 件**")

        for _, row in group_summary.iterrows():
            gid = row["顧客グループID"]
            tbl = f"🪑{row['テーブル番号']} " if str(row.get("テーブル番号", "")).strip() else ""
            with st.expander(
                f"📅 {row['日付']} {tbl}🌍 {row['国籍']} 👥 {row['人数']}名 "
                f"💴 ¥{int(row['合計金額']):,} [ID: {gid}]"
            ):
                gdf = df[df["顧客グループID"] == gid].copy()

                if st.session_state.editing_group == gid:
                    st.markdown("#### ✏️ 編集中")
                    with st.form(f"group_form_{gid}"):
                        ec1, ec2, ec3, ec4, ec5 = st.columns(5)
                        edit_date  = ec1.date_input("来店日", value=pd.to_datetime(row["日付"]).date(), key=f"edate_{gid}")
                        edit_tbl   = ec2.text_input("テーブル番号", value=str(row.get("テーブル番号", "") or ""), key=f"etbl_{gid}")
                        _cur_code = to_nationality_code(row["国籍"])
                        edit_nat   = ec3.selectbox("国籍", NATIONALITY_CODES,
                                        index=NATIONALITY_CODES.index(_cur_code),
                                        format_func=lambda c: nationality_display(c),
                                        key=f"enat_{gid}")
                        edit_ppl   = ec4.number_input("人数", min_value=1, max_value=20,
                                        value=int(row["人数"]), key=f"eppl_{gid}")
                        edit_route = ec5.selectbox("来店経路", VISIT_ROUTES,
                                        index=VISIT_ROUTES.index(row["来店経路"]) if row["来店経路"] in VISIT_ROUTES else 0,
                                        key=f"eroute_{gid}")

                        st.markdown("**注文明細**（数量を変更できます）")
                        edited_rows = []
                        orig_indices = gdf.index.tolist()
                        for i, orig_idx in enumerate(orig_indices):
                            r = gdf.loc[orig_idx]
                            mc1, mc2, mc3 = st.columns([4, 2, 1])
                            mc1.write(f"**{r['メニュー名']}** ¥{int(r['金額']) // int(r['数量']):,}/個")
                            new_q = mc2.number_input("数量", min_value=0, max_value=99,
                                        value=int(r["数量"]), key=f"eq_{gid}_{i}",
                                        label_visibility="collapsed")
                            delete_row = mc3.checkbox("削除", key=f"edel_{gid}_{i}")
                            if not delete_row and new_q > 0:
                                unit_price = int(r["金額"]) // int(r["数量"])
                                edited_rows.append({"メニュー名": r["メニュー名"], "数量": new_q, "金額": unit_price * new_q})

                        st.markdown("**メニューを追加する（プリセット）**")
                        all_menus   = [(name, price) for items in MENU_PRESETS.values() for name, price in items]
                        menu_labels = [f"{n} ¥{p:,}" for n, p in all_menus]
                        add_sel = st.selectbox("追加メニュー", ["（追加しない）"] + menu_labels, key=f"eadd_{gid}")
                        add_qty = st.number_input("追加数量", min_value=1, max_value=99, value=1, key=f"eaddq_{gid}")

                        st.markdown("**手入力でメニューを追加する**")
                        hc1, hc2, hc3 = st.columns(3)
                        manual_add_name  = hc1.text_input("メニュー名", key=f"emanual_name_{gid}")
                        manual_add_price = hc2.number_input("単価（円）", min_value=0, max_value=1000000, value=0, step=10, key=f"emanual_price_{gid}")
                        manual_add_qty   = hc3.number_input("数量", min_value=1, max_value=99, value=1, key=f"emanual_qty_{gid}")

                        sc1, sc2 = st.columns(2)
                        submitted_save   = sc1.form_submit_button("💾 変更を保存", type="primary")
                        submitted_cancel = sc2.form_submit_button("❌ キャンセル")

                    if submitted_save:
                        if add_sel != "（追加しない）":
                            sel_idx  = menu_labels.index(add_sel)
                            add_name, add_price = all_menus[sel_idx]
                            edited_rows.append({"メニュー名": add_name, "数量": add_qty, "金額": add_price * add_qty})
                        if manual_add_name.strip():
                            edited_rows.append({"メニュー名": manual_add_name.strip(), "数量": manual_add_qty, "金額": manual_add_price * manual_add_qty})
                        full_df = load_data()
                        full_df = full_df[full_df["顧客グループID"] != gid]
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        new_rows = [{
                            "日付":           str(edit_date),
                            "テーブル番号":   edit_tbl,
                            "国籍":           edit_nat,
                            "人数":           edit_ppl,
                            "メニュー名":     r["メニュー名"],
                            "数量":           r["数量"],
                            "金額":           r["金額"],
                            "来店経路":       edit_route,
                            "顧客グループID": gid,
                            "登録日時":       now_str,
                        } for r in edited_rows]
                        save_data(pd.concat([full_df, pd.DataFrame(new_rows)], ignore_index=True))
                        st.session_state.editing_group = None
                        st.success("✅ 保存しました")
                        st.rerun()

                    if submitted_cancel:
                        st.session_state.editing_group = None
                        st.rerun()

                else:
                    show_df = gdf[["メニュー名", "数量", "金額"]].copy()
                    show_df["金額"] = show_df["金額"].apply(lambda x: f"¥{x:,}")
                    st.dataframe(show_df.reset_index(drop=True), use_container_width=True)
                    tbl_disp = f"テーブル: {row['テーブル番号']} ｜ " if str(row.get("テーブル番号", "")).strip() else ""
                    st.caption(f"{tbl_disp}来店経路: {row['来店経路']} ｜ 登録日時: {row['登録日時']}")

                    bc1, bc2 = st.columns(2)
                    if bc1.button("✏️ 編集", key=f"edit_group_{gid}"):
                        st.session_state.editing_group = gid
                        st.rerun()
                    if bc2.button("🗑 削除", key=f"del_group_{gid}"):
                        delete_group(gid)
                        st.success("削除しました")
                        st.rerun()

# タブ3：国籍別売上
with tab3:
    st.subheader("📊 国籍別売上集計")
    if df.empty:
        st.info("データがありません。")
    else:
        sales_df = calc_sales_by_nationality(df)
        if not sales_df.empty:
            c1, c2 = st.columns(2)
            with c1:
                fig = px.bar(
                    sales_df.sort_values("売上合計(円)", ascending=True),
                    x="売上合計(円)", y="国籍", orientation="h",
                    title="国籍別 売上合計（円）",
                    color="売上合計(円)", color_continuous_scale="Reds", text="売上合計(円)"
                )
                fig.update_traces(texttemplate="¥%{text:,}", textposition="outside")
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                fig2 = px.pie(sales_df, values="売上合計(円)", names="国籍", title="国籍別 売上シェア")
                st.plotly_chart(fig2, use_container_width=True)
            st.markdown("#### 詳細テーブル")
            s = sales_df.copy()
            s["売上合計(円)"] = s["売上合計(円)"].apply(lambda x: f"¥{x:,}")
            st.dataframe(s, use_container_width=True)

# タブ4：人気メニュー
with tab4:
    st.subheader("🏆 国籍別 人気メニューランキング")
    if df.empty:
        st.info("データがありません。")
    else:
        sel_nat = st.selectbox("🌍 国籍を選択", options=df["国籍"].unique().tolist())
        menu_df = calc_popular_menu_by_nationality(df, sel_nat)
        if menu_df.empty:
            st.info(f"{sel_nat} のデータがありません")
        else:
            c1, c2 = st.columns(2)
            with c1:
                fig = px.bar(
                    menu_df.head(10), x="メニュー名", y="注文数",
                    title=f"{sel_nat} の人気メニュー TOP10",
                    color="注文数", color_continuous_scale="Oranges"
                )
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                rank_df = menu_df.head(10).reset_index(drop=True)
                rank_df.index = rank_df.index + 1
                rank_df.index.name = "順位"
                rank_df["売上合計"] = rank_df["売上合計"].apply(lambda x: f"¥{x:,}")
                st.dataframe(rank_df, use_container_width=True)

# タブ5：平均客単価
with tab5:
    st.subheader("💰 国籍別 平均客単価（1人あたり）")
    if df.empty:
        st.info("データがありません。")
    else:
        avg_df = calc_avg_spend_per_person(df)
        if not avg_df.empty:
            overall_avg = avg_df["客単価(円/人)"].mean()
            st.metric("全国籍の平均客単価", f"¥{int(overall_avg):,} / 人")
            st.markdown("---")
            c1, c2 = st.columns(2)
            with c1:
                fig = px.bar(
                    avg_df, x="国籍", y="客単価(円/人)",
                    title="国籍別 平均客単価（円/人）",
                    color="客単価(円/人)", color_continuous_scale="Blues", text="客単価(円/人)"
                )
                fig.update_traces(texttemplate="¥%{text:,}", textposition="outside")
                fig.update_layout(showlegend=False)
                fig.add_hline(y=overall_avg, line_dash="dash", line_color="red",
                              annotation_text=f"全体平均: ¥{int(overall_avg):,}")
                st.plotly_chart(fig, use_container_width=True)
            with c2:
                d = avg_df.copy()
                d["客単価(円/人)"] = d["客単価(円/人)"].apply(lambda x: f"¥{x:,}")
                st.dataframe(d, use_container_width=True)

# タブ6：一括読み取り
with tab6:
    dx.render_batch_tab(
        st, pd, add_order, ALL_MENUS, NATIONALITY_CODES,
        VISIT_ROUTES, ocr_receipt_with_gemini,
        uuid, date, time,
    )

# タブ7：読み取り精度
with tab7:
    st.subheader("🎯 OCR読み取り精度")
    summary = dx.accuracy_summary()
    if summary is None:
        st.info("まだ記録がありません。伝票を読み取って登録すると集計されます。")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("品目単位精度", f"{summary['品目単位精度']:.1%}")
        c2.metric("伝票単位精度", f"{summary['伝票単位精度']:.1%}")
        c3.metric("読み取り品目数", f"{summary['品目数']:,}")
        c4.metric("伝票枚数", f"{summary['伝票枚数']:,}")
        with open(dx.ACCURACY_LOG, "rb") as _f:
            st.download_button("精度ログCSVをダウンロード", _f, file_name="ocr_accuracy_log.csv", mime="text/csv")
