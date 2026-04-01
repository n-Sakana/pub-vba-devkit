# vba-devkit プロジェクト概要

## 目的

レガシー環境の VBA マクロ資産（500+ ファイル）を新環境に移行するための支援ツール群。

新環境の制約:
- **EDR（Endpoint Detection & Response）**: `Declare PtrSafe Function` を含むファイルを開くとファイルが破損する
- **64bit Office**: 32bit 前提のコード（Long 型ハンドル、PtrSafe 未対応、DAO 3.6）が動かない
- **OneDrive / SharePoint**: ファイルパスの前提が変わる（ローカルパス → URL）
- **セキュリティポリシー強化**: Shell, PowerShell, WScript 等の外部プロセス実行がブロックされる

## 4層アプローチ

```
┌─────────────────────────────────────────────────┐
│ 1. Analyze（静的解析）                            │
│    Excel を開かずにバイナリレベルで VBA コードを解析  │
│    → 何が使われているか、何が危ないかを洗い出す      │
├─────────────────────────────────────────────────┤
│ 2. Probe（環境実測）                              │
│    新環境で VBA コードを注入・実行してテスト         │
│    → 何が通って何がブロックされるかを実測する        │
├─────────────────────────────────────────────────┤
│ 3. Survey（環境棚卸し）                            │
│    端末スペック・ランタイム・Office・Acrobat を棚卸し │
│    → 何が入っていて何が前提にできるかを固定化する    │
├─────────────────────────────────────────────────┤
│ 4. 突き合わせ → 修正対象の確定                      │
│    Analyze の検知結果 × Probe の実測結果             │
│    → 本当に修正が必要なファイル・箇所を特定           │
└─────────────────────────────────────────────────┘
```

### 1. Analyze（静的解析）

Excel を開かずに、OLE2 + MS-OVBA 圧縮形式をバイナリレベルで解析する PowerShell ツール群。

| ツール | 役割 |
|--------|------|
| **Extract** | VBA ソースコードをテキスト抽出（per-file subfolder + combined.txt） |
| **Analyze** | 4カテゴリ検知 + サニタイズ + 移行ガイド + CSV |
| **Diff** | 2ファイル間の VBA コード差分比較 |
| **Unlock** | VBA プロジェクトのパスワード保護解除 |

特徴:
- Excel を一切開かない（Unlock の .xls 変換、Probe の VBA 注入を除く）
- パスワード保護されたプロジェクトも解析可能
- BAT にドラッグ＆ドロップで動作
- コア処理は C#（Add-Type）で高速化

### 2. Probe（環境実測）

PowerShell + Excel COM で一時 xlsm を作成し、VBA コードを注入→保存→再オープン→マクロ実行。EDR やポリシーによるブロックを実測する。

| モード | 説明 |
|--------|------|
| **B** (Basic) | EDR + 互換性 + 境界テスト |
| **E** (Extended) | Basic + Shell/PowerShell/WMI/DDE/SendKeys |
| **G** (Generate) | `probe_storage.xlsm` を生成（SharePoint/OneDrive 比較用） |

結果ステータス: OK / FAIL / BLOCKED / SKIP

### 3. Survey（環境棚卸し）

`Survey.bat` は実測プローブではなく、端末の前提条件を固定するための棚卸しツール。

収集対象:
- PC スペック（OS / CPU / メモリ / GPU / ドライブ）
- PowerShell / .NET / Python / Node.js / Java / WSH
- Office ホスト（Excel / Word / Outlook / Access）
- Adobe Acrobat / Acrobat Pro の実行ファイルと COM automation 登録

出力:
- `survey.txt`
- `survey.json`

### 4. 突き合わせ

```
probe: CreateObject("InternetExplorer.Application") → FAIL
analyze.csv: FileA.xlsm に IE Automation 検知あり
→ FileA.xlsm は IE 関連コードの修正が必要
```

## 検知カテゴリ（4分類）

| カテゴリ | 内容 | ハイライト色 |
|----------|------|-------------|
| **EDR** | Win32 API, Shell, COM, WMI, DLL 等 | 青 |
| **環境依存** (Risk/Review/Info) | パス解決, SaveAs, Dir(), 外部リンク, AutoSave | 緑 |
| **互換性** | PtrSafe, LongPtr, DAO, DefType, DDE 等 | 紫 |
| **業務依存** | Outlook, Word, Access, 印刷, PDF, 外部 EXE | 橙 |

ハイライト優先順: サニタイズ済(黄) > EDR(青) > API(青) > 環境依存(緑) > 互換性(紫) > 業務依存(橙)

### 環境依存の3段階

環境依存パターンは severity で3段階に分類する。危ないのは Path そのものではなく、Path をどう使っているか。

| Severity | 意味 | 例 |
|----------|------|-----|
| **Risk** | クラウド環境で壊れる | `Dir(ThisWorkbook.Path)`, 固定ドライブレター, Path & "\..." |
| **Review** | 文脈次第で危険 | `CurDir`, `BeforeSave` イベント |
| **Info** | 単独では安全、組み合わせで注意 | `ThisWorkbook.Path`, `ActiveWorkbook.Path` |

### サニタイズ

デフォルトで sanitize=true は **Win32 API (Declare) のみ**。EDR がファイルを強制破損させる唯一の確定パターン。
Shell 等は実行時ブロック（ファイル破損なし）のため、デフォルトでは sanitize=false。

## ワークフロー

```
1. Analyze で全 xlsm を一括スキャン
   → analyze.csv（全ファイルの検知結果一覧）
   → 各ファイルの HTML レポート + テキストレポート

2. Probe を新環境で実行
   → probe_result.txt（各パターンの OK/FAIL/BLOCKED）

3. 2つの結果を突き合わせ
   → FAIL/BLOCKED パターンを使っているファイルが修正対象

4. 修正対象ファイルに対して:
   - Analyze のサニタイズ機能で Declare 文を自動コメントアウト
   - HTML レポートのツールチップで代替手段を確認
   - Extract でコードを抽出して手動修正

5. 修正後に Diff で変更確認
```

## プロジェクト構成

```
vba-devkit/
├── Extract.bat / Analyze.bat / Diff.bat / Unlock.bat / Probe.bat
├── config/
│   └── analyze.json         パターンごとの detect/sanitize 設定
├── lib/
│   ├── VBAToolkit.psm1      共通モジュール (OLE2, VBA 圧縮/展開, C# Add-Type,
│   │                        分析エンジン, API 代替 DB 60+ 件, HTML テンプレート)
│   ├── Extract.ps1
│   ├── Analyze.ps1
│   ├── Diff.ps1
│   ├── Unlock.ps1
│   └── Probe.ps1
├── test/                    テストフィクスチャ (.xlsm)
└── docs/                    仕様書・調査結果
```

## 出力

```
output/
├── 20260328_120000_extract/
│   ├── modules/<baseName>/   .bas / .cls / .frm (per-file subfolder)
│   └── <baseName>_combined.txt
├── 20260328_120500_analyze/
│   ├── analyze.csv           全ファイル一覧 (EDR/Compat/Env/Biz/判定列)
│   ├── <name>_analyze.txt    テキストレポート
│   ├── <name>_analyze.html   HTML ビューア (サイドバー + コード + ツールチップ)
│   └── <name>.xlsm           サニタイズ済みコピー (該当時のみ)
├── 20260328_121000_diff/
│   ├── diff.txt
│   └── diff.html
├── 20260328_121500_unlock/
│   └── <name>.xlsm
└── 20260328_122000_probe/
    └── probe_result_<datetime>.txt
```

実行ログは `vba-devkit.log` に追記される。

## 技術スタック

| レイヤー | 技術 |
|----------|------|
| エントリポイント | BAT（ドラッグ＆ドロップ） |
| オーケストレーション | PowerShell 5.1 |
| バイナリ処理 | C#（Add-Type でインラインコンパイル） |
| OLE2 解析 | 自前実装（セクタチェーン、FAT、ディレクトリ） |
| VBA 圧縮/展開 | MS-OVBA 2.4.1 準拠の自前実装 |
| ZIP 操作 | System.IO.Compression |
| HTML ビューア | インライン生成の静的 HTML（ダークテーマ、ミニマップ付き） |
| 設定 GUI | WinForms（Analyze の設定モード） |
| VBA 注入テスト | Excel COM + VBProject（Probe のみ） |
| パスワード解除 | Excel COM（.xls 変換 + DPB= バイナリパッチ、Unlock のみ） |
