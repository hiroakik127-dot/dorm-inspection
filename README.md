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

### 4. Anthropic API キーの取得

1. https://console.anthropic.com/ にログイン
2. 「API Keys」→「Create Key」
3. 生成されたキーをコピーしておく

### 5. 環境変数の設定

#### Windows（PowerShell）

```powershell
$env:SPREADSHEET_ID = "あなたのスプレッドシートID"
$env:ANTHROPIC_API_KEY = "sk-ant-..."
$env:GOOGLE_SERVICE_ACCOUNT_JSON = "service_account.json"  # 省略可
```

#### Windows（コマンドプロンプト）

```cmd
set SPREADSHEET_ID=あなたのスプレッドシートID
set ANTHROPIC_API_KEY=sk-ant-...
```

#### Mac / Linux

```bash
export SPREADSHEET_ID="あなたのスプレッドシートID"
export ANTHROPIC_API_KEY="sk-ant-..."
```

**スプレッドシート ID の確認方法：**
スプレッドシートの URL `https://docs.google.com/spreadsheets/d/【ここがID】/edit` の部分

---

## 実行方法

```bash
python main.py
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
├── main.py                  # メインスクリプト
├── requirements.txt         # 必要ライブラリ
├── README.md                # このファイル
├── service_account.json     # サービスアカウントキー（自分で配置）
├── fonts/                   # 任意：日本語フォントを配置する場合
│   └── NotoSansJP-Regular.ttf
└── reports/                 # 生成された PDF の保存先（自動作成）
    ├── inspection_report_*.pdf
    └── graph_report_*.pdf
```

---

## 日本語フォントについて

**Windows** の場合は自動的にシステムフォント（メイリオ・MS ゴシックなど）を使用します。

**Mac / Linux** でフォントが見つからない場合は以下のいずれかを行ってください：

1. `fonts/` フォルダを作成し、`NotoSansJP-Regular.ttf` を配置
   （Google Fonts からダウンロード可能）
2. または `japanize-matplotlib` を使用（グラフのみ対応）

---

## トラブルシューティング

| エラー | 対処 |
|---|---|
| `SPREADSHEET_ID が設定されていません` | 環境変数を設定してください |
| `service_account.json が見つかりません` | JSON ファイルを正しい場所に配置してください |
| `タブ「修正版全回答」が見つかりません` | スプレッドシートのタブ名を確認してください |
| `部屋番号のデータが見つかりません` | 入力した部屋番号がスプレッドシートに存在するか確認してください |
| `ANTHROPIC_API_KEY が設定されていません` | API キーを環境変数に設定してください |
| 日本語が文字化け | `fonts/` フォルダに日本語フォントを配置してください |
