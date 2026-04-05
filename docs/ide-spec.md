# vba-devkit IDE 仕様書

## コンセプト

vba-devkit を「ツール集」から「VBA 特化の IDE」へ再構成する。
VBA 移行プロジェクトに必要な操作（コード閲覧・解析・差分・アンロック・環境テスト）を一画面で完結させる。

## アーキテクチャ

```
devkit.bat → devkit.ps1
  ├── src/ide/*.cs を連結
  ├── Add-Type -ReferencedAssemblies WebView2 DLL
  └── [DevKit.App]::Run()

┌─────────────────────────────────────────┐
│  C# Host (WinForms + WebView2)          │
│  ├── WebView2 ウィンドウ (全画面)        │
│  ├── ファイル I/O                       │
│  ├── OLE2 / VBA バイナリ解析            │
│  ├── Analyze エンジン                   │
│  └── PowerShell 連携 (EnvTest 等)       │
│          ↕ postMessage                  │
│  JS Frontend (WebView2 内)              │
│  ├── Monaco Editor                      │
│  ├── ファイルツリー                      │
│  ├── Problems パネル                    │
│  ├── Diff viewer (Monaco inline diff)   │
│  └── Output ログ                        │
└─────────────────────────────────────────┘
```

WinForms ウィンドウは WebView2 コントロール 1 つだけ。レイアウトは全て HTML/CSS/JS。
C# ホストはネイティブ処理（ファイル操作、バイナリ解析）に専念する。

## 技術スタック

| レイヤー | 技術 |
|----------|------|
| ランチャー | BAT → PowerShell 5.1 |
| ホスト | C# (Add-Type) + WinForms + WebView2 |
| フロント | HTML / CSS / JS |
| エディタ | Monaco Editor (ローカル同梱) |
| バイナリ解析 | 既存 VBAToolkit.psm1 の C# 部分を移植 |
| 設定 | config/ide.json |

## 依存 (同梱)

| パッケージ | 用途 | 取得方法 |
|-----------|------|---------|
| Monaco Editor | コードエディタ | npm pack → min/ を配置 |
| Microsoft.Web.WebView2 | C# ↔ WebView2 | NuGet → DLL を配置 |

WebView2 Runtime は Windows 11 にプリインストール済み。
依存取得は `setup.bat` で一度だけ実行。以後はオフライン動作。

## 画面構成

```
┌─────────────────────────────────────────────────────────┐
│  DevKit                                       [_][□][×] │
├───────┬─────────────────────────────────┬───────────────┤
│ FILES │  editor (Monaco)          [tab] │  PROBLEMS     │
│       │                                 │               │
│ ▼ 📁  │  Attribute VB_Name = "Mod1"     │  ■ EDR (2)    │
│  ├ A  │                                 │   Declare..   │
│  ├ B  │  Public Sub DoWork()            │   GetWindow.. │
│  └ C  │    Dim p As String              │               │
│       │    p = ThisWorkbook.Path ⚠      │  ■ ENV (3)    │
│ ──────│    Shell "cmd /c ..."   ■       │   ThisWork..  │
│ TOOLS │  End Sub                        │   Dir(path).. │
│       │                                 │   SaveAs..    │
│ Extract│                                │               │
│ Analyze│                                │  ■ COMPAT (1) │
│ Diff   │                                │   DAO 3.6     │
│ Unlock │                                │               │
│ EnvTest│                                │  ■ BIZ (0)    │
├───────┴─────────────────────────────────┴───────────────┤
│  OUTPUT                                                  │
│  [12:00:01] Opened: C:\data\file1.xlsm                  │
│  [12:00:02] Extract: 5 modules found                    │
│  [12:00:03] Analyze: 6 issues (2 EDR, 3 Env, 1 Compat) │
└─────────────────────────────────────────────────────────┘
```

### パネル説明

| パネル | 役割 |
|--------|------|
| FILES | xlsm ファイル一覧 → 展開で VBA モジュールツリー |
| Editor | Monaco Editor。VBA コードをシンタックスハイライト + 検知マーカー付きで表示 |
| PROBLEMS | Analyze 検知結果の一覧。クリックでエディタの該当行へジャンプ |
| TOOLS | 各ツールの実行ボタン |
| OUTPUT | ログ出力 |

### エディタ機能

- VBA シンタックスハイライト (カスタム言語定義)
- Analyze 検知箇所にインラインデコレーション (波線 + 色分け)
- ホバーで検知詳細 + 代替 API 提案をツールチップ表示
- ミニマップ (Monaco 内蔵)
- Diff 表示 (Monaco inline diff viewer)
- 読み取り専用 (VBA ソースの直接編集はスコープ外)

## C# ⇔ JS 通信プロトコル

WebView2 の `postMessage` / `WebMessageReceived` で JSON をやり取りする。

### JS → C# (リクエスト)

```json
{ "id": "req_001", "action": "openFolder", "params": { "path": "C:\\data" } }
{ "id": "req_002", "action": "extract", "params": { "file": "C:\\data\\test.xlsm" } }
{ "id": "req_003", "action": "analyze", "params": { "file": "C:\\data\\test.xlsm" } }
{ "id": "req_004", "action": "readFile", "params": { "path": "C:\\data\\mod.bas" } }
{ "id": "req_005", "action": "diff", "params": { "left": "...", "right": "..." } }
{ "id": "req_006", "action": "unlock", "params": { "file": "..." } }
{ "id": "req_007", "action": "listFiles", "params": { "path": "C:\\data", "pattern": "*.xlsm" } }
```

### C# → JS (レスポンス)

```json
{ "id": "req_001", "status": "ok", "data": { "files": [...] } }
{ "id": "req_002", "status": "ok", "data": { "modules": [...] } }
{ "id": "req_003", "status": "ok", "data": { "issues": [...], "summary": {...} } }
```

### C# → JS (プッシュ通知)

```json
{ "event": "log", "data": { "time": "12:00:01", "message": "..." } }
{ "event": "progress", "data": { "percent": 50, "message": "Analyzing..." } }
```

## アクション一覧

| action | 処理 | ホスト側実装 |
|--------|------|-------------|
| openFolder | フォルダ選択ダイアログ → xlsm 一覧取得 | FolderBrowserDialog + Directory.GetFiles |
| listFiles | 指定パスの xlsm 再帰列挙 | Directory.GetFiles recursive |
| extract | xlsm から VBA モジュール抽出 | OLE2 + VBA decompress (既存 C# 移植) |
| analyze | 抽出コードを解析 | Pattern matching engine (既存移植) |
| readFile | テキストファイル読み込み | File.ReadAllText |
| diff | 2 ファイルの VBA 差分取得 | extract × 2 → テキスト返却 (diff は Monaco 側) |
| unlock | VBA プロジェクトパスワード解除 | DPB binary patch (既存移植) |
| getSettings | analyze 設定読み込み | config/analyze.json 読み込み |
| saveSettings | analyze 設定保存 | config/analyze.json 書き込み |

## ファイル構成

```
vba-devkit/
  devkit.bat                 ← IDE ランチャー
  devkit.ps1                 ← PowerShell ローダー
  setup.bat                  ← 初回セットアップ (依存取得)
  src/
    ide/
      01_App.cs              ← エントリポイント、WebView2 初期化
      02_MessageRouter.cs    ← postMessage ルーティング
      03_FileService.cs      ← ファイル操作
      04_ExtractService.cs   ← OLE2 + VBA 抽出
      05_AnalyzeService.cs   ← 解析エンジン
      06_UnlockService.cs    ← パスワード解除
    ui/
      index.html             ← メイン画面
      css/
        style.css            ← VSCode 風ダークテーマ
      js/
        app.js               ← アプリケーション本体
        host-bridge.js       ← C# 通信レイヤー
        file-tree.js         ← ファイルツリー UI
        problems.js          ← Problems パネル
        vba-language.js      ← Monaco VBA 言語定義
      vendor/
        monaco/              ← Monaco Editor (ローカル同梱)
  lib/                       ← 既存ツール群 (従来の BAT 起動も維持)
  config/
    analyze.json             ← 既存
    ide.json                 ← IDE 設定 (ウィンドウサイズ等)
```

## 既存ツールとの関係

IDE は既存の BAT ツール群を置き換えない。併存する。

- 従来の BAT ドラッグ＆ドロップ運用はそのまま残る
- IDE はバイナリ解析エンジン等のコア処理を C# で直接持つ (psm1 経由ではない)
- EnvTest (Survey/Probe) は PowerShell 実行が必要なので、IDE からは子プロセス起動

## MVP スコープ

### Phase 1: シェル + コードビューア

- [ ] WebView2 ウィンドウ起動
- [ ] フォルダを開く → xlsm 一覧表示
- [ ] xlsm 選択 → Extract → モジュールツリー表示
- [ ] モジュール選択 → Monaco でコード表示 (VBA ハイライト)
- [ ] OUTPUT パネルにログ表示

### Phase 2: Analyze 統合

- [ ] Analyze 実行 → 検知結果取得
- [ ] Monaco 上にインラインデコレーション (波線 + 色分け)
- [ ] ホバーツールチップ (検知詳細 + 代替 API)
- [ ] PROBLEMS パネルに検知一覧
- [ ] クリックで該当行ジャンプ

### Phase 3: Diff + Unlock

- [ ] 2 ファイル選択 → Monaco inline diff viewer
- [ ] Unlock ボタン
- [ ] 結果通知

### Phase 4: EnvTest + Settings

- [ ] EnvTest 起動 → 結果表示
- [ ] Analyze 設定 GUI (analyze.json 編集)
- [ ] IDE 設定 (テーマ、ウィンドウサイズ等)

## VBA 言語定義 (Monaco カスタム言語)

最低限の定義:

- キーワード: Sub, Function, End, Dim, As, Set, If, Then, Else, For, Next, Do, Loop, While, Wend, Select, Case, With, Public, Private, Static, Const, Type, Enum, Property, Get, Let, ByVal, ByRef, Optional, ParamArray, GoTo, On Error, Resume, Exit, New, Nothing, Me, True, False
- 型: String, Long, Integer, Double, Single, Boolean, Date, Variant, Object, Byte, Currency, LongPtr, LongLong
- コメント: ' (シングルクォート) / Rem
- 文字列: "..."
- 行継続: _ (アンダースコア)
- プリプロセッサ: #If, #Else, #End If, #Const
- 特殊: Declare PtrSafe Function/Sub, Attribute VB_Name

## 設計判断

### なぜ WebView2 + Monaco か

- VSCode のエディタエンジンそのもの。VBA IDE として必要な機能が全て API で揃う
- 既存の Analyze HTML ビューア資産 (ダークテーマ、ミニマップ、ツールチップ) の知見が活きる
- WPF でエディタを自前実装するより圧倒的に軽量

### なぜ WinForms + WebView2 全画面か

- WPF の GridSplitter 等は不要 (レイアウトは全て CSS)
- WinForms のほうが Add-Type でのセットアップが軽い
- C# ホストの責務はネイティブ API のみ、UI は一切持たない

### なぜ読み取り専用か

- devkit の目的は「移行対象の分析と選別」であり、VBA コードの編集ではない
- 編集は VBE (Excel VBA Editor) の責務
- 将来的に編集機能を足す余地は残す (Monaco は編集対応)
