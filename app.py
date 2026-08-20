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
import time
from datetime import date, datetime
import dx_addon as dx

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

CSV_FILE   = "yakiniku_orders.csv"
EXCEL_FILE = "yakiniku_orders.xlsx"

NATIONALITIES = [
    "日本", "アメリカ", "中国", "韓国", "台湾",
    "香港", "タイ", "シンガポール", "フランス", "ドイツ",
    "イギリス", "オーストラリア", "カナダ", "その他"
]

VISIT_ROUTES = [
    "Google Maps", "食べログ", "TripAdvisor", "Instagram",
    "口コミ（知人紹介）", "ホテルフロント", "旅行会社", "その他"
]

VISIT_HOURS = [
    "11:00〜12:00", "12:00〜13:00", "13:00〜14:00", "14:00〜15:00",
    "15:00〜16:00", "16:00〜17:00", "17:00〜18:00", "18:00〜19:00",
    "19:00〜20:00", "20:00〜21:00", "21:00〜22:00", "22:00〜23:00",
]

# 生ビール大の価格: 来店日が2026/6/1以前→¥800、6/2以降→¥950（visit_date確定後に動的計算）

# ============================================================
# 道頓堀 焼肉 新恵 メニュープリセット（PDFコード順）
# ============================================================
MENU_PRESETS = {
    "🥩 熟成牛肉": [
        ("特選ロース",    3720),   # A1
        ("上ロース",      2520),   # A2
        ("特選カルビ",    2380),   # A3
        ("上カルビ",      1580),   # A4
        ("特選ハラミ",    2380),   # A5
        ("上ハラミ",      1580),   # A6
        ("特選フィレ",    3380),   # A7
    ],
    "🫀 ホルモン": [
        ("タン",          3000),   # E1
        ("うす切りタン",  1200),   # E2
        ("ツラミ",         880),   # E3
        ("ミノ",           680),   # E4
        ("センマイ",       680),   # E5
        ("アカセンマイ",   680),   # E6
        ("シマチョウ",     680),   # E7
        ("ショウチョウ",   680),   # E8
        ("レバー",         680),   # E9
    ],
    "🌟 セット": [
        ("神戸ビーフセット",           21000),  # S1
        ("神戸ビーフセット スペシャル", 41000),  # S2
    ],
    "🥬 野菜・ナムル": [
        ("チシャ菜",     380),   # G1
        ("キャベツ",     380),   # G2
        ("サラダ",       600),   # G3
        ("ナムル盛合せ", 600),   # G4
        ("モヤシナムル", 380),   # G5
        ("焼野菜",       600),   # G6
    ],
    "🥣 キムチ": [
        ("キムチ盛合せ", 700),   # F1
        ("白菜キムチ",   500),   # F2
        ("大根キムチ",   500),   # F3
        ("胡瓜キムチ",   500),   # F4
    ],
    "🍚 ライス・スープ": [
        ("ライス",       380),   # H1
        ("ビビンバ",     800),   # H2
        ("石焼ビビンバ", 950),   # H3
        ("クッパ",       700),   # H4
        ("ワカメスープ", 500),   # H5
        ("玉子スープ",   500),   # H6
        ("冷麺",         800),   # H7
    ],
    "🍺 ビール": [
        ("生ビール 380ml",  600),              # J1
        ("生ビール 大",     950), # J2 ※来店日で価格を動的切替
    ],
    "🍋 サワー・チューハイ": [
        ("レモンサワー",          550),  # K1
        ("ライムサワー",          550),  # K2
        ("ピーチサワー",          550),  # K3
        ("グレープフルーツサワー", 550),  # K4
        ("カルピスサワー",         550),  # K5
    ],
    "🥃 ハイボール": [
        ("ブラックニッカクリア", 600),   # M1
        ("ジャックダニエル",     700),   # M2
        ("山崎ハイボール",      1400),   # M3
        ("白州ハイボール",      1400),   # M4
        ("知多ハイボール",      1000),   # M5
    ],
    "🍷 ワイン": [
        ("グラス(赤)",              800),   # P1
        ("カベルネ(赤)ハーフボトル", 2200), # P2
    ],
    "🌸 梅酒": [
        ("梅酒(ロック/ソーダ/水割)", 600),  # T1
    ],
    "🥃 焼酎": [
        ("二階堂(麦)",          600),   # L1
        ("佐藤(麦)",            900),   # L2
        ("黒霧島(芋)",          600),   # L3
        ("ジンロマッコリ 375ml", 1000), # L4
    ],
    "🍶 日本酒": [
        ("白鶴 一合 180ml",     600),   # N1
        ("白鶴 二合 360ml",     800),   # N2
        ("冷酒 300ml",          900),   # N3
        ("長兵衛 大吟醸 120ml", 1300),  # N4
    ],
    "🥤 ソフトドリンク": [
        ("ウーロン茶",      400),  # R1
        ("オレンジジュース", 400), # R2
        ("カルピス",        400),  # R3
        ("コーラ",          400),  # R4
        ("ジンジャーエール", 400), # R5
        ("緑茶",            400),  # R6
    ],
}

# 全メニューフラットdict（OCRマッチング・手入力用）
ALL_MENUS = {name: price for items in MENU_PRESETS.values() for name, price in items}

# ============================================================
# データ管理関数
# ============================================================

def load_data():
    if os.path.exists(CSV_FILE):
        df = pd.read_csv(CSV_FILE, encoding="utf-8-sig")
        if "顧客グループID" not in df.columns:
            df["顧客グループID"] = [str(uuid.uuid4())[:8] for _ in range(len(df))]
        if "テーブル番号" not in df.columns:
            df["テーブル番号"] = ""
        if "来店時間帯" not in df.columns:
            df["来店時間帯"] = ""
        return df
    return pd.DataFrame(columns=[
        "日付", "来店時間帯", "テーブル番号", "国籍", "人数", "メニュー名", "数量", "金額",
        "来店経路", "顧客グループID", "登録日時"
    ])


def save_data(df):
    df.to_csv(CSV_FILE, index=False, encoding="utf-8-sig")
    try:
        with open(EXCEL_FILE, "wb") as f:
            f.write(generate_excel(df))
    except Exception as e:
        st.warning(f"⚠️ Excelファイルへの同期保存に失敗しました: {e}")


def add_order(visit_date, visit_hour, table_num, nationality, num_people, menu_items, visit_route, group_id):
    df = load_data()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = [
        {
            "日付":           str(visit_date),
            "来店時間帯":     visit_hour,
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

def calc_sales_by_nationality(df):
    if df.empty:
        return pd.DataFrame()
    return df.groupby("国籍")["金額"].sum().reset_index().rename(columns={"金額": "売上合計(円)"})


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


# ============================================================
# Excelエクスポート関数
# ============================================================

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
    model = genai.GenerativeModel("gemini-2.5-flash")
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

if "ocr_raw" not in st.session_state:
    st.session_state.ocr_raw = []

if "ocr_header" not in st.session_state:
    st.session_state.ocr_header = {}

if "ocr_image" not in st.session_state:
    st.session_state.ocr_image = ""

if "ocr_elapsed" not in st.session_state:
    st.session_state.ocr_elapsed = None

# ============================================================
# ページ設定
# ============================================================

st.set_page_config(
    page_title="焼肉 新恵 DXツール",
    page_icon="🥩",
    layout="wide"
)

st.title("🥩 道頓堀 焼肉 新恵｜DX注文管理ツール")
st.caption("大学ゼミ研究用 | 外国人観光客向け注文データ分析システム")

# ============================================================
# サイドバー：来店情報
# ============================================================

st.sidebar.header("📝 注文入力")
st.sidebar.markdown("---")

st.sidebar.subheader("来店情報")
visit_date  = st.sidebar.date_input("📅 来店日", value=date.today())
# 来店日が2026/6/1以前なら¥800、6/2以降なら¥950
beer_large_price = 800 if visit_date <= date(2026, 6, 1) else 950
visit_hour  = st.sidebar.selectbox("🕐 来店時間帯", options=VISIT_HOURS)
table_num   = st.sidebar.text_input("🪑 テーブル番号", placeholder="例：3、A-2")
nationality = st.sidebar.selectbox("🌍 国籍", options=NATIONALITIES)
num_people  = st.sidebar.number_input("👥 人数（名）", min_value=1, max_value=20, value=2, step=1)
visit_route = st.sidebar.selectbox("📍 来店経路", options=VISIT_ROUTES)

st.sidebar.markdown("---")

# ============================================================
# サイドバー：AI/OCR 伝票読み取り
# ============================================================

st.sidebar.subheader("🤖 AI/OCR 伝票読み取り")
st.sidebar.caption("伝票を撮影してアップロード → AIが自動でカートに追加")

if not GEMINI_AVAILABLE:
    st.sidebar.warning(
        "OCR機能を使うには以下をインストールしてください：\n"
        "```\npip install google-generativeai Pillow\n```"
    )
else:
    gemini_api_key = st.sidebar.text_input(
        "Gemini API Key",
        type="password",
        placeholder="AIza...",
        help="[Google AI Studio](https://aistudio.google.com) で無料取得できます",
    )

    uploaded_img = st.sidebar.file_uploader(
        "伝票画像（JPG / PNG）",
        type=["jpg", "jpeg", "png"],
    )

    if uploaded_img and gemini_api_key:
        st.sidebar.image(uploaded_img, caption="アップロード画像", use_container_width=True)

        if st.sidebar.button("🔍 AIで伝票を読み取る", use_container_width=True, type="primary"):
            with st.spinner("AI読み取り中..."):
                try:
                    _t0 = time.time()
                    image_bytes = uploaded_img.read()
                    ocr_result  = ocr_receipt_with_gemini(image_bytes, gemini_api_key)
                    matched, unmatched = match_menu_to_presets(ocr_result)
                    st.session_state.ocr_results   = matched
                    st.session_state.ocr_unmatched = unmatched
                    st.session_state.ocr_raw       = ocr_result.get("items", []) if isinstance(ocr_result, dict) else []
                    st.session_state.ocr_header    = ocr_result if isinstance(ocr_result, dict) else {}
                    st.session_state.ocr_image     = uploaded_img.name
                    st.session_state.ocr_elapsed   = time.time() - _t0
                    # 警告表示
                    for w in dx.validate_ocr(ocr_result):
                        st.sidebar.warning(w)
                    st.sidebar.success(f"✅ {len(matched)} 品目を読み取りました")
                    if unmatched:
                        st.sidebar.warning(f"⚠️ {len(unmatched)} 品目はメニューに見つかりませんでした")
                except json.JSONDecodeError:
                    st.sidebar.error("❌ AI応答のパースに失敗しました。もう一度お試しください。")
                except Exception as e:
                    st.sidebar.error(f"❌ エラー: {e}")

    if st.session_state.ocr_results is not None:
        st.sidebar.markdown("---")
        st.sidebar.markdown("#### 📋 読み取り結果（確認）")

        confirmed_items = []
        for idx, item in enumerate(st.session_state.ocr_results):
            col_name, col_qty = st.sidebar.columns([3, 1])
            col_name.write(f"**{item['menu_name']}**  ¥{item['unit_price']:,}")
            new_qty = col_qty.number_input(
                "数量", min_value=0, max_value=99,
                value=item["quantity"],
                key=f"ocr_qty_{idx}",
                label_visibility="collapsed",
            )
            if new_qty > 0:
                confirmed_items.append({**item, "quantity": new_qty})

        if st.session_state.ocr_unmatched:
            st.sidebar.markdown("**⚠️ 手動で対応メニューを選択してください**")
            all_menu_labels = ["（スキップ）"] + [f"{n}　¥{p:,}" for n, p in ALL_MENUS.items()]
            all_menu_list = list(ALL_MENUS.items())
            for i, unk in enumerate(st.session_state.ocr_unmatched):
                st.sidebar.caption(f"読み取り: 「{unk['menu_name']}」× {unk['quantity']}")
                sel = st.sidebar.selectbox("対応メニュー", options=all_menu_labels, key=f"ocr_fallback_{i}")
                if sel != "（スキップ）":
                    sel_name, sel_price = all_menu_list[all_menu_labels.index(sel) - 1]
                    confirmed_items.append({"menu_name": sel_name, "quantity": unk["quantity"], "unit_price": sel_price})

        ocr_total = sum(i["unit_price"] * i["quantity"] for i in confirmed_items)
        st.sidebar.markdown(f"**合計: ¥{ocr_total:,}**")

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
            st.sidebar.success("✅ カートに追加しました！")
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
st.sidebar.caption("ボタンを押すと即カートに追加。同じメニューを押すと数量+1")

# 生ビール大の価格告知（来店日連動）
if beer_large_price == 800:
    st.sidebar.info("🍺 生ビール 大：¥800（来店日が2026/6/1以前）")
else:
    st.sidebar.info("🍺 生ビール 大：¥950（来店日が2026/6/2以降）")

for category, items in MENU_PRESETS.items():
    with st.sidebar.expander(category, expanded=False):
        for name, price in items:
            # 生ビール大は来店日で価格を動的切替
            actual_price = beer_large_price if name == "生ビール 大" else price
            if st.button(
                f"{name}　¥{actual_price:,}",
                key=f"btn_{name}",
                use_container_width=True
            ):
                for item in st.session_state.cart:
                    if item["menu_name"] == name:
                        item["quantity"] += 1
                        break
                else:
                    st.session_state.cart.append({"menu_name": name, "quantity": 1, "unit_price": actual_price})
                st.rerun()

# 手入力でメニュー追加
with st.sidebar.expander("✍️ 手入力でメニュー追加", expanded=False):
    manual_name = st.text_input("メニュー名", key="manual_menu_name")
    mc1, mc2 = st.columns(2)
    manual_qty   = mc1.number_input("数量", min_value=1, max_value=99, value=1, key="manual_menu_qty")
    manual_price = mc2.number_input("単価（円）", min_value=0, max_value=1000000, value=0, step=10, key="manual_menu_price")
    if st.button("🛒 カートに追加", key="manual_add_to_cart", use_container_width=True):
        if manual_name.strip():
            for item in st.session_state.cart:
                if item["menu_name"] == manual_name.strip():
                    item["quantity"] += manual_qty
                    break
            else:
                st.session_state.cart.append({"menu_name": manual_name.strip(), "quantity": manual_qty, "unit_price": manual_price})
            st.success(f"✅ {manual_name} を追加しました")
            st.rerun()
        else:
            st.warning("メニュー名を入力してください")

st.sidebar.markdown("---")

# ============================================================
# サイドバー：カート
# ============================================================

st.sidebar.subheader("🧾 カート")

if not st.session_state.cart:
    st.sidebar.caption("カートは空です")
else:
    total = 0
    for idx, item in enumerate(st.session_state.cart):
        subtotal = item["quantity"] * item["unit_price"]
        total += subtotal

        col_name, col_minus, col_num, col_plus, col_del = st.sidebar.columns([4, 1, 1, 1, 1])
        col_name.write(f"**{item['menu_name']}**")

        if col_minus.button("－", key=f"minus_{idx}"):
            if st.session_state.cart[idx]["quantity"] > 1:
                st.session_state.cart[idx]["quantity"] -= 1
            st.rerun()

        col_num.write(f"**{item['quantity']}**")

        if col_plus.button("＋", key=f"plus_{idx}"):
            st.session_state.cart[idx]["quantity"] += 1
            st.rerun()

        if col_del.button("🗑", key=f"del_{idx}"):
            st.session_state.cart.pop(idx)
            st.rerun()

        st.sidebar.caption(f"　¥{item['unit_price']:,} × {item['quantity']} = ¥{subtotal:,}")

    st.sidebar.markdown(f"### 合計: ¥{total:,}")
    st.sidebar.markdown("---")

    if st.sidebar.button("💾 このグループを保存", type="primary", use_container_width=True):
        group_id   = str(uuid.uuid4())[:8]
        menu_items = [(i["menu_name"], i["quantity"], i["quantity"] * i["unit_price"]) for i in st.session_state.cart]
        add_order(visit_date, visit_hour, table_num, nationality, num_people, menu_items, visit_route, group_id)

        # 3-6: 精度ログを記録（AIで読み取った場合のみ）
        if st.session_state.get("ocr_raw"):
            dx.log_reading(
                group_id   = group_id,
                image_name = st.session_state.get("ocr_image", ""),
                ai_items   = st.session_state.get("ocr_raw", []),
                final_items= [{"menu_name": i["menu_name"], "quantity": i["quantity"]}
                               for i in st.session_state.cart],
                elapsed_sec= st.session_state.get("ocr_elapsed"),
            )
            # 来店時間ログ（7月以降の伝票に記載あり）
            _vt = (st.session_state.get("ocr_header") or {}).get("visit_time")
            if _vt:
                dx.log_visit_time(group_id, visit_date, _vt, table_num, nationality, num_people)
            # ログ用セッションをリセット
            st.session_state.ocr_raw     = []
            st.session_state.ocr_header  = {}
            st.session_state.ocr_image   = ""
            st.session_state.ocr_elapsed = None

        st.session_state.cart = []
        st.sidebar.success(f"✅ 保存しました！（ID: {group_id}）")
        st.rerun()

    if st.sidebar.button("🗑 カートをリセット", use_container_width=True):
        st.session_state.cart = []
        st.rerun()

# ============================================================
# メインエリア
# ============================================================

df = load_data()

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📋 データ一覧",
    "👥 顧客グループ管理",
    "📊 国籍別売上",
    "🏆 人気メニュー",
    "💰 平均客単価"
])

# タブ1：データ一覧
with tab1:
    st.subheader("📋 保存済みデータ一覧")
    if df.empty:
        st.info("👈 左のメニューからデータを入力して保存してください")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("総レコード数", f"{len(df)} 件")
        c2.metric("来店国籍数",   f"{df['国籍'].nunique()} 国")
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
            tbl = f"🪑{row['テーブル番号']}　" if str(row.get("テーブル番号", "")).strip() else ""
            with st.expander(
                f"📅 {row['日付']}　{tbl}🌍 {row['国籍']}　👥 {row['人数']}名　"
                f"💴 ¥{int(row['合計金額']):,}　[ID: {gid}]"
            ):
                gdf = df[df["顧客グループID"] == gid].copy()

                if st.session_state.editing_group == gid:
                    st.markdown("#### ✏️ 編集中")
                    ec1, ec2, ec3, ec4, ec5, ec6 = st.columns(6)
                    edit_date  = ec1.date_input("来店日", value=pd.to_datetime(row["日付"]).date(), key=f"edate_{gid}")
                    cur_hour   = row.get("来店時間帯", VISIT_HOURS[7])
                    edit_hour  = ec2.selectbox("来店時間帯", VISIT_HOURS,
                                    index=VISIT_HOURS.index(cur_hour) if cur_hour in VISIT_HOURS else 7,
                                    key=f"ehour_{gid}")
                    edit_tbl   = ec3.text_input("テーブル番号", value=str(row.get("テーブル番号", "") or ""), key=f"etbl_{gid}")
                    edit_nat   = ec4.selectbox("国籍", NATIONALITIES,
                                    index=NATIONALITIES.index(row["国籍"]) if row["国籍"] in NATIONALITIES else 0,
                                    key=f"enat_{gid}")
                    edit_ppl   = ec5.number_input("人数", min_value=1, max_value=20,
                                    value=int(row["人数"]), key=f"eppl_{gid}")
                    edit_route = ec6.selectbox("来店経路", VISIT_ROUTES,
                                    index=VISIT_ROUTES.index(row["来店経路"]) if row["来店経路"] in VISIT_ROUTES else 0,
                                    key=f"eroute_{gid}")

                    st.markdown("**注文明細**（数量を変更できます）")
                    edited_rows = []
                    orig_indices = gdf.index.tolist()
                    for i, orig_idx in enumerate(orig_indices):
                        r = gdf.loc[orig_idx]
                        mc1, mc2, mc3 = st.columns([4, 2, 1])
                        mc1.write(f"**{r['メニュー名']}**　¥{int(r['金額']) // int(r['数量']):,}/個")
                        new_q = mc2.number_input("数量", min_value=0, max_value=99,
                                    value=int(r["数量"]), key=f"eq_{gid}_{i}",
                                    label_visibility="collapsed")
                        delete_row = mc3.checkbox("削除", key=f"edel_{gid}_{i}")
                        if not delete_row and new_q > 0:
                            unit_price = int(r["金額"]) // int(r["数量"])
                            edited_rows.append({"メニュー名": r["メニュー名"], "数量": new_q, "金額": unit_price * new_q})

                    st.markdown("**メニューを追加する（プリセット）**")
                    all_menus   = [(name, price) for items in MENU_PRESETS.values() for name, price in items]
                    menu_labels = [f"{n}　¥{p:,}" for n, p in all_menus]
                    add_sel = st.selectbox("追加メニュー", ["（追加しない）"] + menu_labels, key=f"eadd_{gid}")
                    add_qty = st.number_input("追加数量", min_value=1, max_value=99, value=1, key=f"eaddq_{gid}")

                    st.markdown("**手入力でメニューを追加する**")
                    hc1, hc2, hc3 = st.columns(3)
                    manual_add_name  = hc1.text_input("メニュー名", key=f"emanual_name_{gid}")
                    manual_add_price = hc2.number_input("単価（円）", min_value=0, max_value=1000000, value=0, step=10, key=f"emanual_price_{gid}")
                    manual_add_qty   = hc3.number_input("数量", min_value=1, max_value=99, value=1, key=f"emanual_qty_{gid}")

                    sc1, sc2 = st.columns(2)
                    if sc1.button("💾 変更を保存", key=f"esave_{gid}", type="primary"):
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
                            "来店時間帯":     edit_hour,
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

                    if sc2.button("❌ キャンセル", key=f"ecancel_{gid}"):
                        st.session_state.editing_group = None
                        st.rerun()

                else:
                    show_df = gdf[["メニュー名", "数量", "金額"]].copy()
                    show_df["金額"] = show_df["金額"].apply(lambda x: f"¥{x:,}")
                    st.dataframe(show_df.reset_index(drop=True), use_container_width=True)
                    tbl_disp = f"テーブル: {row['テーブル番号']}　｜　" if str(row.get("テーブル番号", "")).strip() else ""
                    st.caption(f"{tbl_disp}来店経路: {row['来店経路']}　｜　登録日時: {row['登録日時']}")

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

# フッター
st.markdown("---")
st.caption("🎓 大学ゼミ研究用プロトタイプ ｜ データは yakiniku_orders.csv に保存されます")
