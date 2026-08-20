# -*- coding: utf-8 -*-
"""
道頓堀 焼肉 新恵 DXツール 追加モジュール

app.py と同じフォルダに置いてください。

機能：
  1. メニューコード表（中間報告No.04 成果物A に対応）
  2. コードを使ったOCRプロンプト（既存の名前ベースを置き換え）
  3. 精度ログ（AI出力 vs 修正後 を自動記録）
  4. 国籍6区分マッピング（成果物B に対応）
"""

import csv
import json
import re
from datetime import datetime
from pathlib import Path

# ============================================================
# 1. メニューコード表
#    ※ I（アイ）・O（オー）・Q（キュー）は使用しない
#       1・0・数字との誤読を防ぐため
# ============================================================

MENU_CODES = {
    # 熟成牛肉
    "A1": ("特選ロース", 3720), "A2": ("上ロース", 2520),
    "A3": ("特選カルビ", 2380), "A4": ("上カルビ", 1580),
    "A5": ("特選ハラミ", 2380), "A6": ("上ハラミ", 1580),
    "A7": ("特選フィレ", 3380),
    # ホルモン
    "E1": ("タン", 3000), "E2": ("うす切りタン", 1200),
    "E3": ("ツラミ", 880), "E4": ("ミノ", 680),
    "E5": ("センマイ", 680), "E6": ("アカセンマイ", 680),
    "E7": ("シマチョウ", 680), "E8": ("ショウチョウ", 680),
    "E9": ("レバー", 680),
    # セット
    "S1": ("神戸ビーフセット", 21000),
    "S2": ("神戸ビーフセット スペシャル", 41000),
    # 野菜・ナムル
    "G1": ("チシャ菜", 380), "G2": ("キャベツ", 380),
    "G3": ("サラダ", 600), "G4": ("ナムル盛合せ", 600),
    "G5": ("モヤシナムル", 380), "G6": ("焼野菜", 600),
    # キムチ
    "F1": ("キムチ盛合せ", 700), "F2": ("白菜キムチ", 500),
    "F3": ("大根キムチ", 500), "F4": ("胡瓜キムチ", 500),
    # ライス・スープ
    "H1": ("ライス", 380), "H2": ("ビビンバ", 800),
    "H3": ("石焼ビビンバ", 950), "H4": ("クッパ", 700),
    "H5": ("ワカメスープ", 500), "H6": ("玉子スープ", 500),
    "H7": ("冷麺", 800),
    # ビール
    "J1": ("生ビール 380ml", 600),
    "J2": ("生ビール 大", 950),  # ※来店日で800/950が切り替わる
    # サワー・チューハイ
    "K1": ("レモンサワー", 550), "K2": ("ライムサワー", 550),
    "K3": ("ピーチサワー", 550), "K4": ("グレープフルーツサワー", 550),
    "K5": ("カルピスサワー", 550),
    # ハイボール
    "M1": ("ブラックニッカクリア", 600), "M2": ("ジャックダニエル", 700),
    "M3": ("山崎ハイボール", 1400), "M4": ("白州ハイボール", 1400),
    "M5": ("知多ハイボール", 1000),
    # ワイン
    "P1": ("グラス(赤)", 800), "P2": ("カベルネ(赤)ハーフボトル", 2200),
    # 梅酒
    "T1": ("梅酒(ロック/ソーダ/水割)", 600),
    # 焼酎・マッコリ
    "L1": ("二階堂(麦)", 600), "L2": ("佐藤(麦)", 900),
    "L3": ("黒霧島(芋)", 600), "L4": ("ジンロマッコリ 375ml", 1000),
    # 日本酒
    "N1": ("白鶴 一合 180ml", 600), "N2": ("白鶴 二合 360ml", 800),
    "N3": ("冷酒 300ml", 900), "N4": ("長兵衛 大吟醸 120ml", 1300),
    # ソフトドリンク
    "R1": ("ウーロン茶", 400), "R2": ("オレンジジュース", 400),
    "R3": ("カルピス", 400), "R4": ("コーラ", 400),
    "R5": ("ジンジャーエール", 400), "R6": ("緑茶", 400),
}

# 品名 → コード（既存の ALL_MENUS と繋ぐため）
NAME_TO_CODE = {name: code for code, (name, _) in MENU_CODES.items()}


def code_of(name):
    return NAME_TO_CODE.get(name, "")


def name_of(code):
    item = MENU_CODES.get(code)
    return item[0] if item else ""


def price_of(code):
    item = MENU_CODES.get(code)
    return item[1] if item else None


def same_price_groups():
    """同額の品目グループ。金額での自動検証が効かない組み合わせ。"""
    g = {}
    for code, (_, price) in MENU_CODES.items():
        g.setdefault(price, []).append(code)
    return {p: cs for p, cs in g.items() if len(cs) > 1}


def theoretical_old_accuracy():
    """
    旧方式（金額から品目を逆引き）の理論上の正答率。
    同額グループ内でランダムに1つ選んだ場合の期待値。
    12月の発表で「Before値」として使える。
    """
    g = {}
    for code, (_, price) in MENU_CODES.items():
        g.setdefault(price, []).append(code)
    expected = sum(1 / len(cs) for cs in g.values() for _ in cs)
    return expected / len(MENU_CODES)


# ============================================================
# 2. 国籍6区分（中間報告No.04 成果物B）
#    細分化すると1区分あたりの母数が小さくなり、
#    月次比較の統計的な意味が失われるため統合する
# ============================================================

NATIONALITY_GROUPS = {
    "J": ("日本", ["日本"]),
    "C": ("中国系", ["中国", "台湾", "香港", "マカオ"]),
    "K": ("韓国", ["韓国"]),
    "A": ("他アジア", ["タイ", "シンガポール", "ベトナム", "フィリピン",
                       "インドネシア", "マレーシア", "インド"]),
    "W": ("欧米", ["アメリカ", "カナダ", "イギリス", "フランス", "ドイツ",
                   "オーストラリア", "ニュージーランド", "スイス", "イタリア",
                   "スペイン", "オランダ"]),
    "O": ("その他・不明", ["その他", "不明", ""]),
}

_COUNTRY_TO_GROUP = {
    country: code
    for code, (_, countries) in NATIONALITY_GROUPS.items()
    for country in countries
}


def group_code(country_name):
    """個別の国名を6区分コードに変換する。未知の国は O。"""
    return _COUNTRY_TO_GROUP.get(country_name, "O")


def group_label(country_name):
    return NATIONALITY_GROUPS[group_code(country_name)][0]


# ============================================================
# 3. コードを使ったOCRプロンプト
#    既存の ocr_receipt_with_gemini を置き換える
# ============================================================

def build_code_table_text():
    lines = []
    for code, (name, price) in MENU_CODES.items():
        lines.append(f"{code}: {name} ¥{price:,}")
    return "\n".join(lines)


OCR_PROMPT_TEMPLATE = """あなたは焼肉店「新恵」の手書き伝票を読み取る専用AIです。

伝票には、ホールスタッフが以下を手書きしています。
  ・来店日と来店時間
  ・テーブル番号
  ・人数
  ・国籍コード（1文字）
  ・料理名の横のメニューコード

# メニューコード表（これ以外のコードは存在しません）
{code_table}

# 国籍コード表（これ以外のコードは存在しません）
{nat_table}

# 読み取りの絶対ルール
1. メニューコードは必ず「アルファベット1文字 + 数字1文字」の形式です。
2. **I（アイ）・O（オー）・Q（キュー）はメニューコードに使用されていません。**
   これらに見える文字は、数字の 1 や 0、または別の文字の誤認です。
   - 「I」に見えたら → 数字の「1」の可能性を最優先で検討
   - 「O」に見えたら → 数字の「0」の可能性を最優先で検討
   ※ ただし国籍コードの「O（その他・不明）」だけは例外で、実在します。
     国籍欄の O は数字の 0 ではありません。
3. 上の表に存在しないコードは出力しないでください。
   読めない場合は "?" にしてください。**推測で存在しないコードを作らないこと。**
4. 金額が読み取れる場合は、コード表の価格と照合してください。
   食い違う場合は、金額のほうを信じてコードを再検討してください。
5. 数量が明記されていない場合は 1 とします。
6. コードが書かれておらず料理名だけの場合は、料理名から該当コードを判断してください。
7. 来店時間は 24時間表記の "HH:MM" で出力してください（例: 19:30）。
   「7時半」のような表記でも、営業時間から判断して 19:30 としてください。
8. 読み取れない項目は null にしてください。空文字ではなく null です。

# 出力形式（このJSONのみ。説明文や```は不要）
{{
  "visit_date": "2026-07-15",
  "visit_time": "19:30",
  "table_num": "3",
  "num_people": 2,
  "nationality_code": "W",
  "items": [
    {{"code": "A1", "quantity": 1, "amount": 3720, "confidence": "high", "note": ""}},
    {{"code": "H1", "quantity": 2, "amount": 760, "confidence": "high", "note": ""}}
  ],
  "total": 4480,
  "header_confidence": "high"
}}

- confidence / header_confidence: "high" | "medium" | "low"
  自信がない箇所は必ず "low" にしてください。人間が確認します。
  **自信がないのに high と書くことが、このシステムで最も避けたい失敗です。**
- note: 迷った理由を短く記述（例: "A1かA4か不明瞭"）
- amount: 読み取れない場合は null
- visit_date: 年が書かれていない場合は年を省いて "MM-DD" でも構いません
"""


def build_nationality_table_text():
    return "\n".join(
        f"{code}: {label}" for code, (label, _) in NATIONALITY_GROUPS.items()
    )


def build_ocr_prompt():
    return OCR_PROMPT_TEMPLATE.format(
        code_table=build_code_table_text(),
        nat_table=build_nationality_table_text(),
    )


def parse_ocr_response(raw_text):
    """
    ```json フェンスを除去してパースする。
    新形式（辞書）でも旧形式（配列）でも受け取れるようにし、
    常に辞書で返す。
    """
    raw = raw_text.strip()
    m = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", raw)
    if m:
        raw = m.group(1)
    else:
        m = re.search(r"(\{[\s\S]*\}|\[[\s\S]*\])", raw)
        if m:
            raw = m.group(1)
    data = json.loads(raw)

    if isinstance(data, list):          # 旧形式（明細だけ）
        return {"items": data}
    data.setdefault("items", [])
    return data


def validate_ocr(result):
    """
    読み取り結果を検証し、警告文のリストを返す。
    result は辞書（新形式）でも配列（明細のみ）でも受け取れる。
    """
    warnings = []
    if isinstance(result, dict):
        items = result.get("items", [])
        # ヘッダー項目の検証
        if not result.get("visit_time"):
            warnings.append("来店時間が読み取れていません")
        if not result.get("num_people"):
            warnings.append("人数が読み取れていません")
        nat = result.get("nationality_code")
        if not nat:
            warnings.append("国籍コードが読み取れていません")
        elif nat not in NATIONALITY_GROUPS:
            warnings.append(f"国籍コード「{nat}」は表に存在しません")
        if result.get("header_confidence") == "low":
            warnings.append("来店情報（日時・人数・国籍）のAI確信度が低い")
    else:
        items = result

    for i, item in enumerate(items, start=1):
        code = item.get("code", "")
        qty = item.get("quantity") or 1
        amount = item.get("amount")

        if code in ("?", "", None):
            warnings.append(f"{i}行目: コードが読み取れていません")
            continue
        if code not in MENU_CODES:
            warnings.append(f"{i}行目: 「{code}」はコード表に存在しません")
            continue

        expected = price_of(code) * qty
        if amount is not None and amount != expected:
            warnings.append(
                f"{i}行目: {code}（{name_of(code)}）× {qty} は ¥{expected:,} のはずですが "
                f"¥{amount:,} と読まれています"
            )
        if item.get("confidence") == "low":
            note = item.get("note", "")
            warnings.append(
                f"{i}行目: {code} はAIの確信度が低い{f'（{note}）' if note else ''}"
            )

    # 合計金額の照合
    if isinstance(result, dict) and result.get("total"):
        calc = sum(
            (price_of(it.get("code", "")) or 0) * (it.get("quantity") or 1)
            for it in items if it.get("code") in MENU_CODES
        )
        if calc and calc != result["total"]:
            warnings.append(
                f"合計が合いません：明細の積み上げ ¥{calc:,} / 伝票の合計 ¥{result['total']:,}"
            )

    return warnings


# ============================================================
# 4. 精度ログ
#    【最重要】AIの出力と、人が修正した結果を「両方」残す。
#    その差分が読み取り精度そのものであり、
#    12月のBefore/After比較の根拠になる。
#    1〜6月の旧システムは正解データを残していなかったため
#    精度が測れなかった。同じ失敗を繰り返さないこと。
# ============================================================

ACCURACY_LOG = Path("ocr_accuracy_log.csv")

LOG_COLUMNS = [
    "記録日時", "顧客グループID", "画像名",
    "AI_コード", "AI_品名", "AI_数量", "AI確信度",
    "修正後_コード", "修正後_品名", "修正後_数量",
    "修正あり", "修正種別", "所要秒数",
]


def log_reading(group_id, image_name, ai_items, final_items, elapsed_sec=None):
    """
    ai_items   : AIが返した [{"code","quantity","confidence",...}, ...]
    final_items: 人が確定した [{"menu_name","quantity"}, ...]（既存カート形式）

    行数が違う場合（AIが読み落とした／余計に読んだ）も記録する。
    """
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    rows = []

    ai_norm = [
        {
            "code": it.get("code", ""),
            "name": name_of(it.get("code", "")),
            "qty": it.get("quantity") or 1,
            "conf": it.get("confidence", ""),
        }
        for it in (ai_items or [])
    ]
    fin_norm = [
        {
            "code": code_of(it.get("menu_name", "")),
            "name": it.get("menu_name", ""),
            "qty": it.get("quantity") or 1,
        }
        for it in (final_items or [])
    ]

    n = max(len(ai_norm), len(fin_norm))
    for i in range(n):
        a = ai_norm[i] if i < len(ai_norm) else {"code": "", "name": "", "qty": "", "conf": ""}
        f = fin_norm[i] if i < len(fin_norm) else {"code": "", "name": "", "qty": ""}

        if not a["code"] and f["code"]:
            kind = "AIが読み落とし"
            changed = True
        elif a["code"] and not f["code"]:
            kind = "AIが余計に読んだ"
            changed = True
        elif a["code"] != f["code"]:
            kind = "品目の誤り"
            changed = True
        elif a["qty"] != f["qty"]:
            kind = "数量の誤り"
            changed = True
        else:
            kind = ""
            changed = False

        rows.append({
            "記録日時": ts,
            "顧客グループID": group_id,
            "画像名": image_name,
            "AI_コード": a["code"], "AI_品名": a["name"],
            "AI_数量": a["qty"], "AI確信度": a["conf"],
            "修正後_コード": f["code"], "修正後_品名": f["name"],
            "修正後_数量": f["qty"],
            "修正あり": changed,
            "修正種別": kind,
            "所要秒数": round(elapsed_sec, 1) if elapsed_sec else "",
        })

    if not rows:
        return

    write_header = not ACCURACY_LOG.exists()
    with ACCURACY_LOG.open("a", newline="", encoding="utf-8-sig") as fp:
        w = csv.DictWriter(fp, fieldnames=LOG_COLUMNS)
        if write_header:
            w.writeheader()
        w.writerows(rows)


# ------------------------------------------------------------
# 来店時間ログ
#   既存の yakiniku_orders.csv の列構成を壊さないため別ファイルに持つ。
#   1〜6月の伝票には時間が書かれていないため、
#   時間帯分析は7月以降のデータでのみ可能。
#   教員コメント3️⃣「曜日や時間帯への言及が不足」への回答データになる。
# ------------------------------------------------------------

VISIT_TIME_LOG = Path("visit_times.csv")

VISIT_TIME_COLUMNS = [
    "顧客グループID", "来店日", "来店時間", "時間帯",
    "テーブル番号", "国籍", "国籍区分", "人数", "曜日",
]

WEEKDAY_JP = ["月", "火", "水", "木", "金", "土", "日"]


def time_slot(hhmm):
    """来店時間を時間帯に丸める。"""
    try:
        hour = int(str(hhmm).split(":")[0])
    except (ValueError, IndexError):
        return ""
    if hour < 15:
        return "〜14時台"
    if hour < 17:
        return "15-16時台"
    if hour < 19:
        return "17-18時台"
    if hour < 21:
        return "19-20時台"
    return "21時以降"


def log_visit_time(group_id, visit_date, visit_time, table_num,
                   nationality, num_people):
    row = {
        "顧客グループID": group_id,
        "来店日": str(visit_date),
        "来店時間": visit_time,
        "時間帯": time_slot(visit_time),
        "テーブル番号": table_num,
        "国籍": nationality,
        "国籍区分": group_label(nationality),
        "人数": num_people,
        "曜日": WEEKDAY_JP[visit_date.weekday()] if hasattr(visit_date, "weekday") else "",
    }
    write_header = not VISIT_TIME_LOG.exists()
    with VISIT_TIME_LOG.open("a", newline="", encoding="utf-8-sig") as fp:
        w = csv.DictWriter(fp, fieldnames=VISIT_TIME_COLUMNS)
        if write_header:
            w.writeheader()
        w.writerow(row)


def load_visit_times():
    import pandas as pd
    if not VISIT_TIME_LOG.exists():
        return pd.DataFrame(columns=VISIT_TIME_COLUMNS)
    return pd.read_csv(VISIT_TIME_LOG, encoding="utf-8-sig")


def load_accuracy_log():
    import pandas as pd
    if not ACCURACY_LOG.exists():
        return pd.DataFrame(columns=LOG_COLUMNS)
    return pd.read_csv(ACCURACY_LOG, encoding="utf-8-sig")


def accuracy_summary():
    """
    品目単位精度 = 修正されなかった品目数 ÷ 全品目数
    ※ 伝票単位は1品ミスで0%になり改善が見えにくいため副指標
    """
    df = load_accuracy_log()
    if df.empty:
        return None

    total = len(df)
    correct = int((~df["修正あり"].astype(bool)).sum())

    by_receipt = df.groupby("画像名")["修正あり"].apply(lambda s: not s.astype(bool).any())

    return {
        "品目数": total,
        "正答数": correct,
        "品目単位精度": correct / total if total else 0,
        "伝票枚数": int(df["画像名"].nunique()),
        "伝票単位精度": float(by_receipt.mean()) if len(by_receipt) else 0,
        "旧方式の理論値": theoretical_old_accuracy(),
    }


def confusion_table(top_n=15):
    """よく間違えるコードの組み合わせ。プロンプト改善の材料になる。"""
    df = load_accuracy_log()
    if df.empty:
        return df
    miss = df[df["修正あり"].astype(bool)]
    if miss.empty:
        return miss
    return (
        miss.groupby(["AI_コード", "修正後_コード", "修正種別"])
        .size().reset_index(name="件数")
        .sort_values("件数", ascending=False)
        .head(top_n)
    )


# ============================================================
# 5. 一括読み取りモード（PC用・過去分の入力向け）
#    手入力を極力使わず、AI読み取り → その場で修正 → 次へ
#    という流れで大量の伝票を処理する。
# ============================================================

def render_batch_tab(st, pd, add_order, all_menus, nationalities,
                     visit_routes, ocr_func, uuid_mod, date_mod, time_mod):
    """
    app.py 側から以下のように呼び出す：

        dx.render_batch_tab(st, pd, add_order, ALL_MENUS, NATIONALITIES,
                            VISIT_ROUTES, ocr_receipt_with_gemini,
                            uuid, date, time)
    """
    st.subheader("📦 一括読み取り（AI優先）")
    st.caption(
        "複数の伝票画像をまとめてアップロードし、1枚ずつAIに読ませて登録します。"
        "手入力は誤りの修正だけに使います。"
    )

    api_key = st.text_input(
        "Gemini API Key", type="password", key="batch_api_key",
        placeholder="AIza...",
    )

    files = st.file_uploader(
        "伝票画像（複数選択できます）",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="batch_files",
    )

    if not files or not api_key:
        st.info("APIキーを入力し、伝票画像を選んでください。")
        return

    # --- キューの初期化 ---
    names = [f.name for f in files]
    if st.session_state.get("batch_names") != names:
        st.session_state.batch_names = names
        st.session_state.batch_idx = 0
        st.session_state.batch_ocr = {}
        st.session_state.batch_done = []

    idx = st.session_state.batch_idx
    total = len(files)

    if idx >= total:
        st.success(f"✅ {total}枚すべて登録しました。")
        if st.button("最初からやり直す"):
            st.session_state.batch_idx = 0
            st.session_state.batch_ocr = {}
            st.rerun()
        return

    st.progress((idx) / total, text=f"{idx} / {total} 枚")

    current = files[idx]
    col_img, col_edit = st.columns([1, 2])

    with col_img:
        st.image(current, caption=current.name, use_container_width=True)

    # --- OCR（未実行なら自動で走らせる） ---
    if current.name not in st.session_state.batch_ocr:
        with st.spinner("AI読み取り中..."):
            t0 = time_mod.time()
            try:
                raw = ocr_func(current.getvalue(), api_key)
            except Exception as e:
                st.error(f"読み取り失敗: {e}")
                if st.button("この1枚を飛ばす", key=f"skip_{idx}"):
                    st.session_state.batch_idx += 1
                    st.rerun()
                return
            st.session_state.batch_ocr[current.name] = {
                "raw": raw,
                "elapsed": time_mod.time() - t0,
            }

    cache = st.session_state.batch_ocr[current.name]
    result = cache["raw"]
    raw_items = result.get("items", []) if isinstance(result, dict) else result
    header = result if isinstance(result, dict) else {}
    elapsed = cache["elapsed"]

    with col_edit:
        warnings = validate_ocr(result)
        if warnings:
            for w in warnings:
                st.warning(w)
        else:
            st.success(f"警告なし（{elapsed:.1f}秒）")

        # --- 明細の編集（コードはプルダウン。存在しないコードを入力できない） ---
        rows = []
        for it in raw_items:
            code = it.get("code", "")
            rows.append({
                "コード": code if code in MENU_CODES else "",
                "品名": name_of(code),
                "数量": int(it.get("quantity", 1) or 1),
                "確信度": it.get("confidence", ""),
            })
        if not rows:
            rows = [{"コード": "", "品名": "", "数量": 1, "確信度": ""}]

        edited = st.data_editor(
            pd.DataFrame(rows),
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key=f"batch_editor_{idx}",
            column_config={
                "コード": st.column_config.SelectboxColumn(
                    "コード", options=[""] + sorted(MENU_CODES.keys()), width="small"),
                "品名": st.column_config.TextColumn("品名", disabled=True),
                "数量": st.column_config.NumberColumn("数量", min_value=1, step=1, width="small"),
                "確信度": st.column_config.TextColumn("AI確信度", disabled=True, width="small"),
            },
        )

        valid = edited[edited["コード"].isin(MENU_CODES.keys())].copy()
        if len(valid):
            valid["小計"] = valid.apply(
                lambda r: price_of(r["コード"]) * int(r["数量"]), axis=1)
            st.metric("合計", f"¥{int(valid['小計'].sum()):,}")

    st.markdown("---")

    # --- 来店情報（AIが読んだ値を初期値にする。読めなければ前の伝票を引き継ぐ） ---
    st.markdown("##### 来店情報（AIが読み取った値です。違っていれば直してください）")

    # 日付
    ai_date = header.get("visit_date")
    date_default = st.session_state.get("batch_last_date", date_mod.today())
    if ai_date:
        try:
            parts = str(ai_date).split("-")
            if len(parts) == 3:
                date_default = date_mod(int(parts[0]), int(parts[1]), int(parts[2]))
            elif len(parts) == 2:
                date_default = date_mod(date_default.year, int(parts[0]), int(parts[1]))
        except (ValueError, TypeError):
            pass

    c1, c2, c3, c4, c5 = st.columns(5)
    visit_date = c1.date_input("来店日", value=date_default, key=f"bd_{idx}")
    visit_time = c2.text_input(
        "来店時間", value=header.get("visit_time") or "", key=f"btm_{idx}",
        placeholder="19:30")
    table_num = c3.text_input(
        "テーブル番号", value=str(header.get("table_num") or ""), key=f"bt_{idx}")

    # 国籍：AIが読んだ6区分コードを、既存の14区分リストに対応づける
    ai_nat_code = header.get("nationality_code")
    nat_index = 0
    if ai_nat_code in NATIONALITY_GROUPS:
        candidates = NATIONALITY_GROUPS[ai_nat_code][1]
        for cand in candidates:
            if cand in nationalities:
                nat_index = nationalities.index(cand)
                break
    nationality = c4.selectbox(
        "国籍", nationalities, index=nat_index, key=f"bn_{idx}",
        help=f"AIの読み取り: {ai_nat_code or '不明'}")

    people_default = header.get("num_people") or st.session_state.get("batch_last_people", 2)
    num_people = c5.number_input(
        "人数", min_value=1, max_value=20, value=int(people_default), key=f"bp_{idx}")

    visit_route = st.selectbox(
        "来店経路（伝票には書かれていないため手動）", visit_routes,
        index=visit_routes.index(st.session_state.get("batch_last_route", visit_routes[0]))
        if st.session_state.get("batch_last_route") in visit_routes else 0,
        key=f"br_{idx}")

    b1, b2 = st.columns([3, 1])

    if b1.button("✅ 登録して次へ", type="primary", use_container_width=True,
                 key=f"reg_{idx}"):
        if not len(valid):
            st.error("有効な明細がありません。コードを1つ以上選んでください。")
            return

        group_id = str(uuid_mod.uuid4())[:8]
        menu_items = [
            (name_of(r["コード"]),
             int(r["数量"]),
             all_menus.get(name_of(r["コード"]), price_of(r["コード"])))
            for _, r in valid.iterrows()
        ]
        add_order(visit_date, table_num, nationality, num_people,
                  menu_items, visit_route, group_id)

        # 来店時間を別ファイルに記録（既存CSVの列を壊さないため）
        if visit_time:
            log_visit_time(group_id, visit_date, visit_time, table_num,
                           nationality, num_people)

        # 精度ログ（これが12月の中核データ）
        log_reading(
            group_id=group_id,
            image_name=current.name,
            ai_items=raw_items,
            final_items=[{"menu_name": n, "quantity": q} for n, q, _ in menu_items],
            elapsed_sec=elapsed,
        )

        # 次の伝票へ設定を引き継ぐ
        st.session_state.batch_last_date = visit_date
        st.session_state.batch_last_people = num_people
        st.session_state.batch_last_route = visit_route
        st.session_state.batch_idx += 1
        st.rerun()

    if b2.button("⏭ 飛ばす", use_container_width=True, key=f"skp_{idx}"):
        st.session_state.batch_idx += 1
        st.rerun()
