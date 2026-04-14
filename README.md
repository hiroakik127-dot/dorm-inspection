# 寮設備点検レポート生成ツール

Google スプレッドシートの点検データを読み込み、PDF レポートを 2 つ自動生成します。

---

## 生成される PDF

| ファイル | 内容 |
|---|---|
| `reports/inspection_report_*.pdf` | 点検レポート（良い点・問題点・傾向分析・推奨対応） |
| `reports/graph_report_*.pdf` | 傾向グラフレポート（部屋別件数・設備別件数・時系列など） |

---

## セットアップ手順

### 1. Python のインストール

Python 3.11 以上が必要です。
https://www.python.org/downloads/

### 2. ライブラリのインストール

```bash
cd Documnets
cd dorm_inspection
pip install -r requirements.txt
```

### 3. Google サービスアカウントの作成と設定

#### 3-1. Google Cloud プロジェクトを作成

1. https://console.cloud.google.com/ を開く
2. 「プロジェクトを作成」→ 任意の名前でプロジェクトを作成

#### 3-2. API を有効化

1. 左メニュー「API とサービス」→「ライブラリ」
2. 「Google Sheets API」を検索して有効化
3. 「Google Drive API」も同様に有効化

#### 3-3. サービスアカウントを作成

1. 「API とサービス」→「認証情報」→「認証情報を作成」→「サービスアカウント」
2. サービスアカウント名を入力して作成
3. 作成したサービスアカウントをクリック →「キー」タブ
4. 「鍵を追加」→「新しい鍵を作成」→「JSON」→「作成」
5. ダウンロードされた JSON ファイルを `dorm_inspection` フォルダ内に
   `service_account.json` という名前で保存

#### 3-4. スプレッドシートを共有

1. Google スプレッドシートを開く
2. 右上「共有」をクリック
3. サービスアカウントのメールアドレス（JSON ファイル内の `client_email` の値）を
   「閲覧者」として追加

### 4. Gemini API キーの取得

1. https://aistudio.google.com にログイン
2. 「API Keys」→「Create Key」
3. 生成されたキーをコピーしておく

### 5. 環境変数の設定

1. `.streamlit`フォルダを`dorm-inspection`フォルダ内に作る
2. `.streamlit`フォルダ内に`secrets.toml`ファイルをメモ帳などで作る
3. 以下を書き込む
```
SPREADSHEET_ID = "あなたのスプレッドシートID"
GEMINI_API_KEY = "あなたのGeminiAPIキー"
GOOGLE_SERVICE_ACCOUNT_JSON = "service_account.json"
```


**スプレッドシート ID の確認方法：**
スプレッドシートの URL `https://docs.google.com/spreadsheets/d/【ここがID】/edit` の部分

---

## 実行方法

```bash
python main.py
or
streamlit run app.py
```

起動後、プロンプトが表示されます：

```
部屋番号を入力してください（例: 101）
全部屋まとめてレポートを作成する場合は「全体」と入力
入力:
```

- **部屋番号を入力**（例: `101`）→ その部屋の点検レポートを生成
- **「全体」と入力** → 寮全体の点検レポートを生成

---

## スプレッドシートの形式

- 参照タブ名：**「修正版全回答」**
- 1行目がヘッダー（列名）になっている形式
- 部屋番号を含む列名に「部屋」「号室」「room」などのキーワードが入っていると自動検出されます
- 日付列（「タイムスタンプ」「点検日」「日時」など）があると時系列グラフが生成されます

---

## ファイル構成

```
dorm_inspection/
├── main.py                  # コマンドライン版スクリプト
├── app.py                   # Streamlit Web アプリ
├── requirements.txt         # 必要ライブラリ
├── README.md                # このファイル
├── service_account.json     # サービスアカウントキー（自分で配置・gitignore済み）
├── fonts/                   # 日本語フォントを配置（要ダウンロード）
│   └── NotoSansJP-Regular.ttf  # ← 手順を参照して配置してください
└── reports/                 # 生成された PDF の保存先（自動作成・gitignore済み）
    ├── inspection_report_*.pdf
    └── graph_report_*.pdf
```

---

## 日本語フォントのセットアップ（必須）

Streamlit Cloud および Linux 環境では、システムに日本語フォントがないため、
リポジトリの `fonts/` フォルダに `NotoSansJP-Regular.ttf` を配置する必要があります。

### ダウンロード手順

#### 方法1：Google Fonts からブラウザでダウンロード

1. 以下の URL にアクセスして ZIP をダウンロード
   `https://fonts.google.com/noto/specimen/Noto+Sans+JP`
2. 「Download family」ボタンをクリック
3. ZIP を解凍し、`NotoSansJP-Regular.ttf` を `fonts/` フォルダに配置

#### 方法2：コマンドラインでダウンロード

```bash
# Linux / Mac
curl -L "https://github.com/googlefonts/noto-cjk/raw/main/Sans/OTF/Japanese/NotoSansCJKjp-Regular.otf" \
  -o fonts/NotoSansJP-Regular.ttf

# または pip で直接取得（Python 環境）
python -c "
import urllib.request
url = 'https://github.com/googlefonts/noto-cjk/releases/download/Sans2.004/03_NotoSansJP.zip'
# ZIPを解凍して fonts/ に配置してください
print('上記 URL から ZIP をダウンロードして解凍してください')
"
```

#### 方法3：Windows ローカル実行の場合

Windows では `C:\Windows\Fonts\` のシステムフォント（メイリオ等）が自動的に使用されるため、
`fonts/` フォルダへの配置は不要です。

### Streamlit Cloud へのデプロイ時

`fonts/NotoSansJP-Regular.ttf` を Git リポジトリにコミットして push してください。
アプリ起動時に自動的に検出されます。

```bash
git add fonts/NotoSansJP-Regular.ttf
git commit -m "日本語フォントを追加"
git push origin main
```

---

## トラブルシューティング

| エラー | 対処 |
|---|---|
| `SPREADSHEET_ID が設定されていません` | 環境変数を設定してください |
| `service_account.json が見つかりません` | JSON ファイルを正しい場所に配置してください |
| `タブ「修正版全回答」が見つかりません` | スプレッドシートのタブ名を確認してください |
| `部屋番号のデータが見つかりません` | 入力した部屋番号がスプレッドシートに存在するか確認してください |
| `ANTHROPIC_API_KEY が設定されていません` | API キーを環境変数に設定してください |
| 日本語が文字化け | `fonts/NotoSansJP-Regular.ttf` を配置してください（上記「日本語フォントのセットアップ」参照） |
