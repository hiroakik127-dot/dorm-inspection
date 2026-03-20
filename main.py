#!/usr/bin/env python3
"""
寮設備点検レポート生成ツール
Google スプレッドシートの点検データを読み込み、PDFレポートを2つ生成します。

使い方:
    python main.py

環境変数:
    SPREADSHEET_ID              : Google スプレッドシートの ID
    GOOGLE_SERVICE_ACCOUNT_JSON : サービスアカウントJSONファイルのパス（省略時: service_account.json）
    GEMINI_API_KEY              : Google Gemini API キー
"""

import os
import sys
import json
import datetime
import platform
from pathlib import Path

# ============================================================
# サードパーティライブラリのインポート
# ============================================================
try:
    import gspread
    from google.oauth2.service_account import Credentials
except ImportError:
    print("エラー: gspread または google-auth がインストールされていません")
    print("  pip install gspread google-auth")
    sys.exit(1)

try:
    from google import genai
except ImportError:
    print("エラー: google-genai がインストールされていません")
    print("  pip install google-genai")
    sys.exit(1)

try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("エラー: pandas または numpy がインストールされていません")
    print("  pip install pandas numpy")
    sys.exit(1)

try:
    import matplotlib
    matplotlib.use("Agg")  # GUI不要のバックエンド
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm
except ImportError:
    print("エラー: matplotlib がインストールされていません")
    print("  pip install matplotlib")
    sys.exit(1)

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
    print("エラー: reportlab がインストールされていません")
    print("  pip install reportlab")
    sys.exit(1)


# ============================================================
# 設定
# ============================================================

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "service_account.json")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

TAB_NAME = "修正版全回答"
GEMINI_MODEL = "gemini-2.5-flash"

GOOGLE_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

OUTPUT_DIR = "reports"


# ============================================================
# 日本語フォント設定
# ============================================================

def find_japanese_font() -> str | None:
    """日本語 TTF/TTC フォントを探す。リポジトリ同梱フォントを最優先で確認する。"""
    # リポジトリ同梱フォントを最優先（Streamlit Cloud / CI 対応）
    local_first = [
        Path("fonts/NotoSansJP-Regular.ttf"),
        Path("fonts/NotoSansJP-VariableFont_wght.ttf"),
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
    else:  # Linux
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


def setup_fonts() -> tuple[str, bool]:
    """
    ReportLab と matplotlib 用の日本語フォントを設定する。
    Returns:
        (reportlab_font_name, success)
    """
    font_path = find_japanese_font()

    if font_path is None:
        print("警告: 日本語フォントが見つかりませんでした。")
        print("  → fontsフォルダに NotoSansJP-Regular.ttf を配置するか、")
        print("    pip install japanize-matplotlib を実行してください。")
        try:
            import japanize_matplotlib  # noqa: F401
            print("  → japanize-matplotlib を使用します。")
        except ImportError:
            print("  → フォールバック: グラフの日本語が文字化けする可能性があります。")
        return "Helvetica", False

    print(f"日本語フォントを使用: {font_path}")

    # ReportLab 用フォント登録
    rl_font_name = "JapaneseFont"
    try:
        pdfmetrics.registerFont(TTFont(rl_font_name, font_path))
    except Exception as e:
        print(f"ReportLab フォント登録エラー: {e}")
        rl_font_name = "Helvetica"

    # matplotlib 用フォント設定
    try:
        fm.fontManager.addfont(font_path)
        fp = fm.FontProperties(fname=font_path)
        plt.rcParams["font.family"] = fp.get_name()
        plt.rcParams["axes.unicode_minus"] = False
    except Exception as e:
        print(f"matplotlib フォント設定エラー: {e}")

    return rl_font_name, True


# ============================================================
# Google Sheets 接続・データ読み込み
# ============================================================

def connect_to_sheets():
    """Google Sheets API に接続して Spreadsheet オブジェクトを返す。"""
    if not SPREADSHEET_ID:
        print("エラー: 環境変数 SPREADSHEET_ID が設定されていません。")
        sys.exit(1)

    json_path = Path(SERVICE_ACCOUNT_JSON)
    if not json_path.exists():
        print(f"エラー: サービスアカウントファイルが見つかりません: {json_path}")
        print("  環境変数 GOOGLE_SERVICE_ACCOUNT_JSON でパスを指定してください。")
        sys.exit(1)

    try:
        creds = Credentials.from_service_account_file(str(json_path), scopes=GOOGLE_SCOPES)
        client = gspread.authorize(creds)
        return client.open_by_key(SPREADSHEET_ID)
    except Exception as e:
        print(f"Google Sheets 接続エラー: {e}")
        sys.exit(1)


def load_data(spreadsheet) -> pd.DataFrame:
    """「修正版全回答」タブからデータを読み込む。"""
    try:
        ws = spreadsheet.worksheet(TAB_NAME)
        records = ws.get_all_records()
    except gspread.exceptions.WorksheetNotFound:
        print(f"エラー: タブ「{TAB_NAME}」が見つかりません。")
        sys.exit(1)
    except Exception as e:
        print(f"データ読み込みエラー: {e}")
        sys.exit(1)

    if not records:
        print(f"エラー: タブ「{TAB_NAME}」にデータがありません。")
        sys.exit(1)

    df = pd.DataFrame(records)
    print(f"データ読み込み完了: {len(df)} 行 / {len(df.columns)} 列")
    print(f"列一覧: {list(df.columns)}")
    return df


def find_room_column(df: pd.DataFrame) -> str | None:
    """部屋番号が入っている列名を自動検出する。"""
    keywords = ["部屋番号", "部屋", "号室", "room", "ルーム"]
    for col in df.columns:
        if any(kw in str(col).lower() for kw in keywords):
            return col
    print(f"警告: 部屋番号列を自動検出できませんでした。列: {list(df.columns)}")
    return None


def filter_by_room(df: pd.DataFrame, room_number: str, room_col: str | None) -> pd.DataFrame:
    """特定部屋のデータに絞り込む。"""
    if room_col is None:
        print("警告: 部屋番号列が不明なため全件を対象にします。")
        return df

    mask = df[room_col].astype(str).str.strip() == str(room_number).strip()
    filtered = df[mask]

    if filtered.empty:
        available = sorted(df[room_col].astype(str).unique())
        print(f"エラー: 部屋番号「{room_number}」のデータが見つかりません。")
        print(f"  利用可能な部屋番号: {available}")
        sys.exit(1)

    return filtered


# ============================================================
# Gemini API による分析
# ============================================================

def generate_analysis(df: pd.DataFrame, room_number: str | None, is_all: bool) -> dict:
    """Gemini API でデータを分析し、辞書形式の考察を返す。"""
    if not GEMINI_API_KEY:
        print("エラー: 環境変数 GEMINI_API_KEY が設定されていません。")
        sys.exit(1)

    client = genai.Client(api_key=GEMINI_API_KEY)

    # 利用可能なモデル一覧を表示
    print("利用可能な Gemini モデル一覧:")
    try:
        for m in client.models.list():
            print(f"  - {m.name}")
    except Exception as e:
        print(f"  モデル一覧の取得に失敗しました: {e}")
    print(f"使用モデル: {GEMINI_MODEL}\n")

    # データを JSON 文字列に変換（大きすぎる場合は切り詰め）
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

    print("Gemini API で分析中（しばらくお待ちください）...")
    try:
        response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
        raw = response.text.strip()

        # コードブロック除去
        if raw.startswith("```"):
            raw = raw.split("```")[1]
            if raw.startswith("json"):
                raw = raw[4:]
            raw = raw.strip()

        return json.loads(raw)

    except json.JSONDecodeError:
        print("警告: Gemini の応答を JSON としてパースできませんでした。テキストとして扱います。")
        return {
            "good_points": "",
            "improvement_points": raw,
            "trend_analysis": "",
            "future_issues": "",
            "summary": "",
        }
    except Exception as e:
        print(f"Gemini API エラー: {e}")
        sys.exit(1)


# ============================================================
# ReportLab ユーティリティ
# ============================================================

COLOR_HEADER = colors.HexColor("#1a237e")
COLOR_SECTION = colors.HexColor("#3949ab")
COLOR_SECTION_BG = colors.HexColor("#e8eaf6")


def build_styles(font: str) -> dict:
    styles = {
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
    return styles


def bullets_to_paragraphs(text: str, style) -> list:
    """改行区切りのテキストを Paragraph リストに変換する。"""
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
# PDF① 点検レポート
# ============================================================

def create_report_pdf(
    analysis: dict,
    df: pd.DataFrame,
    room_number: str | None,
    is_all: bool,
    font: str,
) -> str:
    today_str = datetime.date.today().strftime("%Y年%m月%d日")
    if is_all:
        title = "寮設備点検レポート（全体）"
        filename = Path(OUTPUT_DIR) / f"inspection_report_all_{datetime.date.today()}.pdf"
    else:
        title = f"寮設備点検レポート（{room_number}号室）"
        filename = Path(OUTPUT_DIR) / f"inspection_report_room{room_number}_{datetime.date.today()}.pdf"

    doc = SimpleDocTemplate(
        str(filename), pagesize=A4,
        rightMargin=20*mm, leftMargin=20*mm, topMargin=25*mm, bottomMargin=20*mm,
    )
    s = build_styles(font)
    story = []

    # ── ヘッダー ──
    story += [
        Paragraph(title, s["title"]),
        Paragraph(f"作成日：{today_str}", s["subtitle"]),
        Spacer(1, 6*mm),
        HRFlowable(width="100%", thickness=2, color=COLOR_SECTION),
        Spacer(1, 5*mm),
    ]

    # ── データ概要 ──
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

    # ── 各セクション ──
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
            content = "（データなし）"
            story.append(Paragraph(content, s["body"]))
        elif is_bullet:
            paras = bullets_to_paragraphs(content, s["bullet"])
            story.extend(paras if paras else [Paragraph("（データなし）", s["body"])])
        else:
            story.extend(text_to_paragraphs(content, s["body"]))
        story.append(Spacer(1, 4*mm))

    # ── フッター ──
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
    print(f"PDF① 生成完了: {filename}")
    return str(filename)


# ============================================================
# PDF② 傾向グラフレポート
# ============================================================

_COLOR_PALETTE = ["#3949ab", "#e53935", "#43a047", "#fb8c00", "#8e24aa", "#00acc1", "#f06292"]


def _safe_save(fig, path: str) -> bool:
    try:
        fig.savefig(path, dpi=150, bbox_inches="tight")
        return True
    except Exception as e:
        print(f"グラフ保存エラー ({path}): {e}")
        return False
    finally:
        plt.close(fig)


def generate_graph_room_count(df: pd.DataFrame, room_col: str | None) -> str:
    """グラフ1：部屋別 点検記録件数の棒グラフ"""
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
    path = "_graph1_tmp.png"
    _safe_save(fig, path)
    return path


def generate_graph_facility_count(df: pd.DataFrame) -> str:
    """グラフ2：設備別 関連記録件数の横棒グラフ（テキスト解析）"""
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
    path = "_graph2_tmp.png"
    _safe_save(fig, path)
    return path


def generate_graph_time_series(df: pd.DataFrame, room_col: str | None) -> str | None:
    """グラフ3：月別 点検件数の折れ線グラフ（日付列がある場合）"""
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

        path = "_graph3_tmp.png"
        _safe_save(fig, path)
        return path
    except Exception as e:
        print(f"時系列グラフ生成エラー: {e}")
        return None


def generate_graph_rating(df: pd.DataFrame) -> str | None:
    """グラフ4：評価点数の分布（評価列がある場合）"""
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
    path = "_graph4_tmp.png"
    _safe_save(fig, path)
    return path


def create_graph_pdf(
    df: pd.DataFrame,
    room_number: str | None,
    is_all: bool,
    font: str,
) -> str:
    today_str = datetime.date.today().strftime("%Y年%m月%d日")
    if is_all:
        title = "寮設備点検 傾向グラフレポート（全体）"
        filename = Path(OUTPUT_DIR) / f"graph_report_all_{datetime.date.today()}.pdf"
    else:
        title = f"寮設備点検 傾向グラフレポート（{room_number}号室）"
        filename = Path(OUTPUT_DIR) / f"graph_report_room{room_number}_{datetime.date.today()}.pdf"

    print("グラフを生成中...")
    room_col = find_room_column(df)

    graph_entries: list[tuple[str, str]] = []

    p1 = generate_graph_room_count(df, room_col)
    graph_entries.append(("グラフ1：部屋別 点検記録件数", p1))

    p2 = generate_graph_facility_count(df)
    graph_entries.append(("グラフ2：設備別 関連記録件数", p2))

    p3 = generate_graph_time_series(df, room_col)
    if p3:
        graph_entries.append(("グラフ3：月別 点検件数の推移", p3))

    p4 = generate_graph_rating(df)
    if p4:
        graph_entries.append(("グラフ4：評価点数の分布", p4))

    # PDF 組み立て
    doc = SimpleDocTemplate(
        str(filename), pagesize=A4,
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

    for heading, img_path in graph_entries:
        story.append(Paragraph(f"■ {heading}", s["section"]))
        story.append(Spacer(1, 3*mm))
        if img_path and Path(img_path).exists():
            story.append(Image(img_path, width=170*mm, height=80*mm))
        else:
            story.append(Paragraph("グラフの生成に失敗しました。", s["body"]))
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

    # 一時ファイル削除
    for _, p in graph_entries:
        if p and Path(p).exists():
            try:
                Path(p).unlink()
            except Exception:
                pass

    print(f"PDF② 生成完了: {filename}")
    return str(filename)


# ============================================================
# メイン
# ============================================================

def main():
    print("=" * 55)
    print("  寮設備点検レポート生成ツール")
    print("=" * 55)

    # フォント設定
    font_name, _ = setup_fonts()

    # 出力ディレクトリ作成
    Path(OUTPUT_DIR).mkdir(exist_ok=True)

    # ユーザー入力
    print("\n部屋番号を入力してください（例: 101）")
    print("全部屋まとめてレポートを作成する場合は「全体」と入力")
    user_input = input("入力: ").strip()
    if not user_input:
        print("エラー: 入力が空です。")
        sys.exit(1)

    is_all = user_input in ("全体", "全", "すべて", "all", "ALL")
    room_number = None if is_all else user_input

    # データ読み込み
    print("\nGoogle Sheets に接続中...")
    spreadsheet = connect_to_sheets()
    df_all = load_data(spreadsheet)

    room_col = find_room_column(df_all)

    # 対象データ絞り込み
    if is_all:
        target_df = df_all
        print(f"\n全体モード: {len(target_df)} 件のデータを対象にします。")
    else:
        target_df = filter_by_room(df_all, room_number, room_col)
        print(f"\n{room_number}号室: {len(target_df)} 件のデータが見つかりました。")

    # Gemini 分析
    print()
    analysis = generate_analysis(target_df, room_number, is_all)

    # PDF 生成
    print("\nPDF を生成中...")
    pdf1 = create_report_pdf(analysis, target_df, room_number, is_all, font_name)
    pdf2 = create_graph_pdf(df_all, room_number, is_all, font_name)

    print("\n" + "=" * 55)
    print("  完了！")
    print(f"  PDF①（点検レポート）  : {pdf1}")
    print(f"  PDF②（グラフレポート）: {pdf2}")
    print("=" * 55)


if __name__ == "__main__":
    main()
