#!/usr/bin/env python3
"""
寮設備点検レポート生成ツール（Streamlit 版）

環境変数 または .streamlit/secrets.toml で以下を設定してください:
    SPREADSHEET_ID              : Google スプレッドシートの ID
    GOOGLE_SERVICE_ACCOUNT_JSON : サービスアカウントJSONファイルのパス（省略時: service_account.json）
    GEMINI_API_KEY              : Google Gemini API キー
"""

import os
import io
import json
import datetime
import platform
from pathlib import Path

import streamlit as st

# ============================================================
# サードパーティライブラリのインポート
# ============================================================
try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    st.error("エラー: gspread または google-auth がインストールされていません。`pip install gspread google-auth`")
    st.stop()

try:
    from google import genai
except ImportError:
    st.error("エラー: google-genai がインストールされていません。`pip install google-genai`")
    st.stop()

try:
    import pandas as pd
except ImportError:
    st.error("エラー: pandas がインストールされていません。`pip install pandas`")
    st.stop()

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
except ImportError:
    st.error("エラー: matplotlib がインストールされていません。`pip install matplotlib`")
    st.stop()

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Image,
        Table, TableStyle, HRFlowable
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
except ImportError:
    st.error("エラー: reportlab がインストールされていません。`pip install reportlab`")
    st.stop()


# ============================================================
# 設定読み込み（環境変数 → st.secrets の順）
# ============================================================

def _get_secret(key: str, default: str = "") -> str:
    val = os.environ.get(key, "")
    if not val:
        try:
            val = st.secrets.get(key, default)
        except Exception:
            val = default
    return val or default


SPREADSHEET_ID = _get_secret("SPREADSHEET_ID")
GEMINI_API_KEY = _get_secret("GEMINI_API_KEY")

TAB_NAME = "修正版全回答"
GEMINI_MODEL = "gemini-2.5-flash"

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

COLOR_HEADER = colors.HexColor("#1a237e")
COLOR_SECTION = colors.HexColor("#3949ab")
COLOR_SECTION_BG = colors.HexColor("#e8eaf6")
_COLOR_PALETTE = ["#3949ab", "#e53935", "#43a047", "#fb8c00", "#8e24aa", "#00acc1", "#f06292"]


# ============================================================
# 日本語フォント設定
# ============================================================

@st.cache_resource
def setup_fonts() -> tuple[str, bool]:
    """ReportLab と matplotlib 用の日本語フォントを設定する。"""
    font_path = _find_japanese_font()

    if font_path is None:
        try:
            import japanize_matplotlib  # noqa: F401
        except ImportError:
            pass
        return "Helvetica", False

    rl_font_name = "JapaneseFont"
    try:
        pdfmetrics.registerFont(TTFont(rl_font_name, font_path))
    except Exception:
        rl_font_name = "Helvetica"

    try:
        fm.fontManager.addfont(font_path)
        fp = fm.FontProperties(fname=font_path)
        plt.rcParams["font.family"] = fp.get_name()
        plt.rcParams["axes.unicode_minus"] = False
    except Exception:
        pass

    return rl_font_name, True


def _find_japanese_font() -> str | None:
    # リポジトリ同梱フォントを最優先（Streamlit Cloud 対応）
    local_first = [
        Path("fonts/NotoSansJP-Regular.ttf"),
        Path("fonts/ipaexg.ttf"),
        Path("ipaexg.ttf"),
    ]
    for p in local_first:
        if p.exists():
            return str(p)

    # システムフォント（ローカル実行時のフォールバック）
    system = platform.system()
    if system == "Windows":
        font_dir = Path("C:/Windows/Fonts")
        candidates = ["meiryo.ttc", "msgothic.ttc", "YuGothM.ttc", "YuGothR.ttc", "msmincho.ttc"]
        paths = [font_dir / f for f in candidates]
    elif system == "Darwin":
        candidates = [
            "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
            "/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
        ]
        paths = [Path(p) for p in candidates]
    else:
        candidates = [
            "/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf",
            "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
            "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        ]
        paths = [Path(p) for p in candidates]

    for p in paths:
        if p.exists():
            return str(p)
    return None


# ============================================================
# Google Sheets 接続・データ読み込み
# ============================================================

def connect_to_sheets():
    if not SPREADSHEET_ID:
        raise ValueError("環境変数 SPREADSHEET_ID が設定されていません。")

    # Streamlit Secrets から認証情報を取得（優先）
    try:
        sa_dict = dict(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
        client = gspread.service_account_from_dict(sa_dict)
        return client.open_by_key(SPREADSHEET_ID)
    except KeyError:
        pass  # Secrets にない場合はファイルにフォールバック

    # ローカル開発用: service_account.json ファイルから読み込み
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "service_account.json")
    json_path = Path(sa_json)
    if not json_path.exists():
        raise FileNotFoundError(
            f"サービスアカウントファイルが見つかりません: {json_path}\n"
            "Streamlit Secrets に GOOGLE_SERVICE_ACCOUNT を設定するか、\n"
            "環境変数 GOOGLE_SERVICE_ACCOUNT_JSON でファイルパスを指定してください。"
        )

    creds = Credentials.from_service_account_file(str(json_path), scopes=GOOGLE_SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(SPREADSHEET_ID)


def load_data(spreadsheet) -> pd.DataFrame:
    try:
        ws = spreadsheet.worksheet(TAB_NAME)
        records = ws.get_all_records()
    except gspread.exceptions.WorksheetNotFound:
        raise ValueError(f"タブ「{TAB_NAME}」が見つかりません。")

    if not records:
        raise ValueError(f"タブ「{TAB_NAME}」にデータがありません。")

    return pd.DataFrame(records)


def find_room_column(df: pd.DataFrame) -> str | None:
    keywords = ["部屋番号", "部屋", "号室", "room", "ルーム"]
    for col in df.columns:
        if any(kw in str(col).lower() for kw in keywords):
            return col
    return None


def filter_by_room(df: pd.DataFrame, room_number: str, room_col: str | None) -> pd.DataFrame:
    if room_col is None:
        return df

    mask = df[room_col].astype(str).str.strip() == str(room_number).strip()
    filtered = df[mask]

    if filtered.empty:
        available = sorted(df[room_col].astype(str).unique())
        raise ValueError(
            f"部屋番号「{room_number}」のデータが見つかりません。\n"
            f"利用可能な部屋番号: {available}"
        )

    return filtered


# ============================================================
# Gemini API による分析
# ============================================================

def generate_analysis(df: pd.DataFrame, room_number: str | None, is_all: bool) -> dict:
    if not GEMINI_API_KEY:
        raise ValueError("環境変数 GEMINI_API_KEY が設定されていません。")

    client = genai.Client(api_key=GEMINI_API_KEY)

    data_str = df.to_json(orient="records", indent=2, force_ascii=False)
    if len(data_str) > 60000:
        data_str = data_str[:60000] + "\n...(データが長いため省略)"

    if is_all:
        prompt = f"""以下は寮の設備点検データです（全部屋）。
このデータを分析し、以下の JSON 形式だけで回答してください（余計な文章・コードブロック不要）。

データ:
{data_str}

{{
  "good_points": "寮全体の良い点・維持できている点（箇条書き、1行1項目、各行頭は「・」）",
  "improvement_points": "寮全体の悪い点・改善すべき点（同上）",
  "trend_analysis": "全体の傾向分析（繰り返し問題、設備集中など。複数段落可）",
  "future_issues": "今後予想される問題と推奨対応策（箇条書き）",
  "summary": "全体総括（3〜5文）"
}}"""
    else:
        prompt = f"""以下は寮の{room_number}号室の設備点検データです。
このデータを分析し、以下の JSON 形式だけで回答してください（余計な文章・コードブロック不要）。

データ:
{data_str}

{{
  "improvement_points": "{room_number}号室の改善すべき点・問題点（箇条書き、1行1項目、各行頭は「・」）",
  "good_points": "{room_number}号室の良い点・維持できている点（同上）",
  "trend_analysis": "繰り返し発生している問題や傾向の分析（複数段落可）",
  "future_issues": "今後予想される問題と推奨対応策（箇条書き）",
  "summary": "この部屋の総括（3〜5文）"
}}"""

    response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
    raw = response.text.strip()

    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"):
            raw = raw[4:]
        raw = raw.strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {
            "good_points": "",
            "improvement_points": raw,
            "trend_analysis": "",
            "future_issues": "",
            "summary": "",
        }


# ============================================================
# ReportLab ユーティリティ
# ============================================================

def build_styles(font: str) -> dict:
    return {
        "title": ParagraphStyle("title", fontName=font, fontSize=20, spaceAfter=4,
                                textColor=COLOR_HEADER, alignment=TA_CENTER),
        "subtitle": ParagraphStyle("subtitle", fontName=font, fontSize=12, spaceAfter=2,
                                   textColor=colors.HexColor("#283593"), alignment=TA_CENTER),
        "section": ParagraphStyle("section", fontName=font, fontSize=12, spaceBefore=10,
                                  spaceAfter=4, textColor=colors.white,
                                  backColor=COLOR_SECTION, leftIndent=6),
        "body": ParagraphStyle("body", fontName=font, fontSize=10, spaceAfter=3,
                               leading=16, alignment=TA_JUSTIFY),
        "bullet": ParagraphStyle("bullet", fontName=font, fontSize=10, spaceAfter=2,
                                 leading=15, leftIndent=12),
        "footer": ParagraphStyle("footer", fontName=font, fontSize=8,
                                 textColor=colors.grey, alignment=TA_CENTER),
    }


def bullets_to_paragraphs(text: str, style) -> list:
    result = []
    for line in text.strip().splitlines():
        line = line.strip()
        if not line:
            continue
        if not line.startswith("・"):
            line = "・" + line
        result.append(Paragraph(line, style))
    return result


def text_to_paragraphs(text: str, style) -> list:
    return [Paragraph(line.strip(), style) for line in text.strip().splitlines() if line.strip()]


# ============================================================
# PDF① 点検レポート（BytesIO を返す）
# ============================================================

def create_report_pdf(
    analysis: dict,
    df: pd.DataFrame,
    room_number: str | None,
    is_all: bool,
    font: str,
) -> bytes:
    today_str = datetime.date.today().strftime("%Y年%m月%d日")
    title = "寮設備点検レポート（全体）" if is_all else f"寮設備点検レポート（{room_number}号室）"

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=20*mm, leftMargin=20*mm, topMargin=25*mm, bottomMargin=20*mm,
    )
    s = build_styles(font)
    story = []

    story += [
        Paragraph(title, s["title"]),
        Paragraph(f"作成日：{today_str}", s["subtitle"]),
        Spacer(1, 6*mm),
        HRFlowable(width="100%", thickness=2, color=COLOR_SECTION),
        Spacer(1, 5*mm),
    ]

    story.append(Paragraph("■ データ概要", s["section"]))
    story.append(Spacer(1, 3*mm))

    room_col = find_room_column(df)
    overview = [["総点検件数", f"{len(df)} 件"]]
    if is_all and room_col:
        overview.append(["対象部屋数", f"{df[room_col].nunique()} 部屋"])
    elif not is_all:
        overview.append(["対象部屋", f"{room_number}号室"])

    tbl = Table(overview, colWidths=[55*mm, 100*mm])
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), font),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BACKGROUND", (0, 0), (0, -1), COLOR_SECTION_BG),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("PADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 5*mm))

    sections = [
        ("■ 良い点・維持できている点", "good_points", True),
        ("■ 改善すべき点・問題点", "improvement_points", True),
        ("■ 傾向分析", "trend_analysis", False),
        ("■ 今後予想される問題・推奨対応", "future_issues", True),
        ("■ 総括", "summary", False),
    ]

    for heading, key, is_bullet in sections:
        story.append(Paragraph(heading, s["section"]))
        story.append(Spacer(1, 3*mm))
        value = analysis.get(key, "")
        if isinstance(value, list):
            value = "\n".join(str(v) for v in value)
        content = value.strip()
        if not content:
            story.append(Paragraph("（データなし）", s["body"]))
        elif is_bullet:
            paras = bullets_to_paragraphs(content, s["bullet"])
            story.extend(paras if paras else [Paragraph("（データなし）", s["body"])])
        else:
            story.extend(text_to_paragraphs(content, s["body"]))
        story.append(Spacer(1, 4*mm))

    story += [
        Spacer(1, 6*mm),
        HRFlowable(width="100%", thickness=1, color=colors.grey),
        Spacer(1, 2*mm),
        Paragraph(
            f"本レポートは Google Gemini ({GEMINI_MODEL}) により自動生成されました",
            s["footer"],
        ),
    ]

    doc.build(story)
    return buf.getvalue()


# ============================================================
# PDF② 傾向グラフレポート（BytesIO を返す）
# ============================================================

def _fig_to_buf(fig) -> io.BytesIO:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def generate_graph_room_count(df: pd.DataFrame, room_col: str | None) -> io.BytesIO:
    fig, ax = plt.subplots(figsize=(10, 5))
    if room_col and room_col in df.columns:
        counts = df[room_col].astype(str).value_counts().sort_index()
        bars = ax.bar(
            range(len(counts)), counts.values,
            color=[_COLOR_PALETTE[i % len(_COLOR_PALETTE)] for i in range(len(counts))],
            alpha=0.85,
        )
        ax.set_xticks(range(len(counts)))
        ax.set_xticklabels(counts.index, rotation=45, ha="right", fontsize=9)
        for bar, val in zip(bars, counts.values):
            ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.05,
                    str(val), ha="center", va="bottom", fontsize=8)
        ax.set_title("部屋別 点検記録件数", fontsize=13, fontweight="bold")
        ax.set_xlabel("部屋番号")
        ax.set_ylabel("件数")
    else:
        ax.text(0.5, 0.5, "部屋番号データなし", ha="center", va="center",
                transform=ax.transAxes, fontsize=12)
        ax.set_title("部屋別 点検記録件数", fontsize=13)
    plt.tight_layout()
    return _fig_to_buf(fig)


def generate_graph_facility_count(df: pd.DataFrame) -> io.BytesIO:
    facility_keywords = {
        "エアコン・空調": ["エアコン", "クーラー", "空調", "暖房", "冷房"],
        "照明・電気": ["照明", "電気", "電球", "ライト", "蛍光灯", "LED"],
        "水回り": ["水", "排水", "蛇口", "トイレ", "シャワー", "浴室", "洗面"],
        "壁・床・天井": ["壁", "床", "クロス", "天井", "剥がれ", "ひび"],
        "窓・ドア": ["窓", "ドア", "鍵", "扉", "網戸", "サッシ"],
        "インターネット": ["Wi-Fi", "ネット", "インターネット", "LAN", "通信"],
        "家具・収納": ["家具", "収納", "棚", "ベッド", "机", "椅子"],
    }

    text_cols = df.select_dtypes(include=["object"]).columns
    counts: dict[str, int] = {}

    for facility, keywords in facility_keywords.items():
        cnt = 0
        for _, row in df.iterrows():
            hit = any(
                kw in str(row[col])
                for col in text_cols
                for kw in keywords
                if pd.notna(row[col])
            )
            if hit:
                cnt += 1
        if cnt > 0:
            counts[facility] = cnt

    fig, ax = plt.subplots(figsize=(10, 5))
    if counts:
        items = sorted(counts.items(), key=lambda x: x[1], reverse=True)
        labels = [i[0] for i in items]
        values = [i[1] for i in items]
        bars = ax.barh(
            range(len(labels)), values,
            color=[_COLOR_PALETTE[i % len(_COLOR_PALETTE)] for i in range(len(labels))],
            alpha=0.85,
        )
        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(labels)
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + 0.05, bar.get_y() + bar.get_height() / 2,
                    str(val), ha="left", va="center", fontsize=8)
        ax.set_title("設備別 関連記録件数", fontsize=13, fontweight="bold")
        ax.set_xlabel("件数")
    else:
        ax.text(0.5, 0.5, "設備別データを集計できませんでした",
                ha="center", va="center", transform=ax.transAxes, fontsize=12)
        ax.set_title("設備別 関連記録件数", fontsize=13)
    plt.tight_layout()
    return _fig_to_buf(fig)


def generate_graph_time_series(df: pd.DataFrame) -> io.BytesIO | None:
    date_keywords = ["タイムスタンプ", "日時", "点検日", "date", "time", "日付"]
    date_col = next(
        (c for c in df.columns if any(kw in str(c) for kw in date_keywords)), None
    )
    if date_col is None:
        return None

    try:
        df2 = df.copy()
        df2["_parsed"] = pd.to_datetime(df2[date_col], errors="coerce")
        df2 = df2.dropna(subset=["_parsed"])
        if df2.empty:
            return None
        df2["_ym"] = df2["_parsed"].dt.to_period("M")
        monthly = df2.groupby("_ym").size()

        fig, ax = plt.subplots(figsize=(10, 4))
        ax.plot(
            range(len(monthly)), monthly.values,
            marker="o", color=_COLOR_PALETTE[0], linewidth=2, markersize=6,
        )
        ax.fill_between(range(len(monthly)), monthly.values, alpha=0.15, color=_COLOR_PALETTE[0])
        ax.set_xticks(range(len(monthly)))
        ax.set_xticklabels([str(p) for p in monthly.index], rotation=45, ha="right", fontsize=9)
        ax.set_title("月別 点検記録件数の推移", fontsize=13, fontweight="bold")
        ax.set_xlabel("年月")
        ax.set_ylabel("件数")
        plt.tight_layout()
        return _fig_to_buf(fig)
    except Exception:
        return None


def generate_graph_rating(df: pd.DataFrame) -> io.BytesIO | None:
    rating_kw = ["評価", "点数", "スコア", "満足", "rating", "score"]
    numeric_cols = df.select_dtypes(include=["number"]).columns
    room_col = find_room_column(df)
    rating_cols = [
        c for c in numeric_cols
        if any(kw in str(c) for kw in rating_kw) and c != room_col
    ]
    if not rating_cols:
        return None

    cols_to_plot = rating_cols[:2]
    fig, axes = plt.subplots(1, len(cols_to_plot), figsize=(6 * len(cols_to_plot), 4))
    if len(cols_to_plot) == 1:
        axes = [axes]

    for ax, col in zip(axes, cols_to_plot):
        vals = df[col].dropna()
        if vals.empty:
            continue
        vc = vals.value_counts().sort_index()
        ax.bar(vc.index.astype(str), vc.values, color=_COLOR_PALETTE[2], alpha=0.85)
        ax.set_title(f"{col} の分布", fontsize=11, fontweight="bold")
        ax.set_xlabel("評価値")
        ax.set_ylabel("件数")

    plt.tight_layout()
    return _fig_to_buf(fig)


def create_graph_pdf(
    df: pd.DataFrame,
    room_number: str | None,
    is_all: bool,
    font: str,
) -> bytes:
    today_str = datetime.date.today().strftime("%Y年%m月%d日")
    title = (
        "寮設備点検 傾向グラフレポート（全体）"
        if is_all
        else f"寮設備点検 傾向グラフレポート（{room_number}号室）"
    )

    room_col = find_room_column(df)
    graph_entries: list[tuple[str, io.BytesIO | None]] = [
        ("グラフ1：部屋別 点検記録件数", generate_graph_room_count(df, room_col)),
        ("グラフ2：設備別 関連記録件数", generate_graph_facility_count(df)),
        ("グラフ3：月別 点検件数の推移", generate_graph_time_series(df)),
        ("グラフ4：評価点数の分布", generate_graph_rating(df)),
    ]

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        rightMargin=15*mm, leftMargin=15*mm, topMargin=20*mm, bottomMargin=15*mm,
    )
    s = build_styles(font)
    story = []

    story += [
        Paragraph(title, s["title"]),
        Paragraph(f"作成日：{today_str}", s["subtitle"]),
        Spacer(1, 5*mm),
        HRFlowable(width="100%", thickness=2, color=COLOR_SECTION),
        Spacer(1, 5*mm),
    ]

    for heading, img_buf in graph_entries:
        if img_buf is None:
            continue
        story.append(Paragraph(f"■ {heading}", s["section"]))
        story.append(Spacer(1, 3*mm))
        story.append(Image(img_buf, width=170*mm, height=80*mm))
        story.append(Spacer(1, 6*mm))

    story += [
        HRFlowable(width="100%", thickness=1, color=colors.grey),
        Spacer(1, 2*mm),
        Paragraph(
            f"本レポートは Google Gemini ({GEMINI_MODEL}) により自動生成されました",
            s["footer"],
        ),
    ]

    doc.build(story)
    return buf.getvalue()


# ============================================================
# Streamlit UI
# ============================================================

st.set_page_config(page_title="寮設備点検レポート生成ツール", page_icon="🏠", layout="centered")

st.title("🏠 寮設備点検レポート生成ツール")
st.markdown("部屋番号または「全体」を入力してボタンを押すと、PDF レポートを2つ生成します。")

st.divider()

room_input = st.text_input(
    "部屋番号（例: 101）または「全体」",
    placeholder="101 または 全体",
)

generate_btn = st.button("📄 PDF を生成", type="primary", disabled=not room_input.strip())

if generate_btn and room_input.strip():
    user_input = room_input.strip()
    is_all = user_input in ("全体", "全", "すべて", "all", "ALL")
    room_number = None if is_all else user_input
    today = datetime.date.today()

    font_name, _ = setup_fonts()

    try:
        with st.spinner("Google Sheets に接続中..."):
            spreadsheet = connect_to_sheets()
            df_all = load_data(spreadsheet)

        room_col = find_room_column(df_all)

        if is_all:
            target_df = df_all
            st.info(f"全体モード: {len(target_df)} 件のデータを対象にします。")
        else:
            target_df = filter_by_room(df_all, room_number, room_col)
            st.info(f"{room_number}号室: {len(target_df)} 件のデータが見つかりました。")

        with st.spinner("Gemini API で分析中（しばらくお待ちください）..."):
            analysis = generate_analysis(target_df, room_number, is_all)

        with st.spinner("PDF を生成中..."):
            pdf1_bytes = create_report_pdf(analysis, target_df, room_number, is_all, font_name)
            pdf2_bytes = create_graph_pdf(df_all, room_number, is_all, font_name)

        st.success("PDF の生成が完了しました！下のボタンからダウンロードしてください。")

        suffix = "all" if is_all else f"room{room_number}"
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="⬇️ PDF① 点検レポート",
                data=pdf1_bytes,
                file_name=f"inspection_report_{suffix}_{today}.pdf",
                mime="application/pdf",
            )
        with col2:
            st.download_button(
                label="⬇️ PDF② グラフレポート",
                data=pdf2_bytes,
                file_name=f"graph_report_{suffix}_{today}.pdf",
                mime="application/pdf",
            )

    except (ValueError, FileNotFoundError) as e:
        st.error(str(e))
    except Exception as e:
        st.error(f"予期しないエラーが発生しました: {e}")
