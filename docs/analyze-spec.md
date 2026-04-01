# Analyze 仕様書

## 概要

Analyze は vba-devkit の中核ツール。Excel を開かずに OLE2 バイナリレベルで VBA コードを解析し、4カテゴリのリスクを検知する。

3モード:
1. **設定 GUI**（引数なし）: WinForms ダイアログでパターンごとに detect/sanitize を設定
2. **ファイル分析**: ファイルをドロップ → HTML ビューア + テキストレポート + サニタイズ済みコピー + CSV
3. **フォルダ分析**: フォルダをドロップ → 再帰走査で全 xlsm/xlam/xls を一括分析

## 検知カテゴリ

### 4本柱

| カテゴリ | 内容 | ハイライト色 | パターン数 |
|----------|------|-------------|-----------|
| **EDR** | Win32 API, Shell, COM, WMI, DLL, レジストリ等 | 青 `#1b2e4a` | 16 |
| **環境依存** | パス解決, SaveAs, Dir(), 外部リンク, AutoSave | 緑 `#1b3a2a` | 20+ |
| **互換性** | PtrSafe, DAO, DDE, IE, レガシーコントロール等 | 紫 `#3a1b4a` | 10 |
| **業務依存** | Outlook, Word, Access, 印刷, PDF, 外部 EXE | 橙 `#4a3a1b` | 6 |

ハイライト優先順: サニタイズ済(黄) > EDR(青) > API(青) > 環境依存(緑) > 互換性(紫) > 業務依存(橙)

環境依存を互換性より上に置く理由: 移行レビューでは SharePoint/OneDrive 由来のパス問題が最重要であり、PtrSafe 等の互換性問題より先に視認する必要がある。

## 環境依存パターンの設計思想

> 危ないのは `Path` そのものではなく、`Path をどう使っているか` である。

背景:
- `ThisWorkbook.Path` / `FullName` は OneDrive 同期ローカルでも SharePoint Open-in-App でも URL になる
- `Dir(ThisWorkbook.Path)` のような相対解決はクラウド条件で失敗する
- `SaveAs` / `Workbooks.Open` 自体は死んでいないが、入力パスがクラウド条件では URL 寄りになる
- `SaveAs to TEMP` は安定している
- `AutoSave` はクラウド条件で True になる

### 3段階 Severity

| Severity | 意味 | 例 |
|----------|------|-----|
| **Risk** | クラウド環境で壊れやすい前提を直接持つ | `Dir(ThisWorkbook.Path)`, 固定ドライブレター, `Path & "\..."` |
| **Review** | 単体では断定しにくいが周辺文脈次第で危険 | `CurDir`, `BeforeSave` イベント, `Workbooks.Open someVar` |
| **Info** | 単独では安全、確認促進の意味が主 | `ThisWorkbook.Path`, `ActiveWorkbook.Path` |

## 環境依存パターン一覧

### Risk パターン

| パターン名 | 検知対象 | 正規表現 |
|---|---|---|
| 固定ドライブレター | `C:\`, `D:\` 等 | `(?mi)^[^'\r\n]*"[A-Z]:\\"` |
| UNC パス | `\\server\share` | `(?mi)^[^'\r\n]*"\\\\[^"]+\\` |
| ユーザーフォルダ | `C:\Users\` | `(?mi)^[^'\r\n]*C:\\Users\\` |
| Desktop / Documents | 固定パスでのデスクトップ/ドキュメント参照 | `(?mi)^[^'\r\n]*\\(Desktop\|Documents\|ドキュメント\|デスクトップ)\\` |
| AppData | AppData パス | `(?mi)^[^'\r\n]*\\AppData\\` |
| Program Files | Program Files パス | `(?mi)^[^'\r\n]*\\Program Files` |
| 固定プリンタ名 | ActivePrinter への文字列設定 | `(?mi)^[^'\r\n]*\.ActivePrinter\s*=\s*"` |
| 固定 IP アドレス | ハードコードされた IP | `(?mi)^[^'\r\n]*"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"` |
| 固定接続先 | Server= / Host= | `(?mi)^[^'\r\n]*(Server\s*=\|Host\s*=)` |
| ~~localhost~~ | *(Review に降格 — 下記参照)* | |
| 接続文字列 | Provider= / DSN= / Data Source= | `(?mi)^[^'\r\n]*(Provider\s*=\|DSN\s*=\|Data\s+Source\s*=)` |
| 外部ブック参照 (リテラル) | Workbooks.Open に文字列リテラル | `(?mi)^[^'\r\n]*\bWorkbooks\.Open\s*\(\s*"` |
| Dir() パスチェック | Dir() によるパス存在確認 | `(?mi)^[^'\r\n]*\bDir\s*\(` |
| パス文字列連結 | Path & "\..." 型の連結 | 複合パターン |
| SaveAs 呼び出し | SaveAs | `(?mi)^[^'\r\n]*\.SaveAs\b` |
| 外部ブック参照 | 数式内の外部参照 | `\[.*\]` |
| LinkSources / UpdateLink | 外部リンク操作 | `(?mi)^[^'\r\n]*\.(LinkSources\|UpdateLink)` |

### Review パターン

| パターン名 | 検知対象 | 正規表現 |
|---|---|---|
| BeforeSave イベント | 保存前イベント依存 | `(?mi)\b(Workbook_BeforeSave\|BeforeSave)\b` |
| AfterSave イベント | 保存後イベント依存 | `(?mi)\b(Workbook_AfterSave\|AfterSave)\b` |
| Workbooks.Open (変数) | 変数引数での Open | `(?mi)^[^'\r\n]*\bWorkbooks\.Open\b` (リテラル以外) |
| CurDir | カレントディレクトリ | `(?mi)^[^'\r\n]*\bCurDir\b` |
| ChDir | ディレクトリ変更 | `(?mi)^[^'\r\n]*\bChDir\b` |
| localhost | localhost 参照（バックエンド通信の可能性） | `(?mi)^[^'\r\n]*\blocalhost\b` |

### Info パターン

| パターン名 | 検知対象 | 理由 |
|---|---|---|
| ThisWorkbook.Path | パス値参照 | SharePoint/OneDrive では URL を返す場合あり。単独では安全 |
| ActiveWorkbook.Path | パス値参照 | 同上 |
| ThisWorkbook.FullName | フルパス値参照 | 同上 |
| ActiveWorkbook.FullName | フルパス値参照 | 同上 |

Info は HTML ハイライトなし。テキストレポートに出力、CSV の `InfoCount` 列で件数を記録。

## 業務依存パターン一覧

| パターン名 | 検知対象 | 正規表現 |
|---|---|---|
| Outlook 連携 | Outlook COM | `(?mi)^[^'\r\n]*\bOutlook\.Application\b` |
| Word 連携 | Word COM | `(?mi)^[^'\r\n]*\bWord\.Application\b` |
| Access / DB 連携 | Access COM / DAO / ADO | `(?mi)^[^'\r\n]*\b(Access\.Application\|CurrentDb\|DoCmd)\b` |
| PDF 出力 | ExportAsFixedFormat | `(?mi)^[^'\r\n]*\.ExportAsFixedFormat\b` |
| 印刷 | PrintOut / PrintPreview | `(?mi)^[^'\r\n]*\.(PrintOut\|PrintPreview)\b` |
| 外部 EXE 起動 | Shell で exe 起動 | `(?mi)^[^'\r\n]*\bShell\s.*\.exe` |

## サニタイズ

デフォルトで sanitize=true は **Win32 API (Declare) のみ**。

理由: EDR がファイルを強制破損させるのは Declare 文が存在する場合のみ（ファイルを開く/保存する時点で破損）。Shell 等は実行時ブロック（ファイル自体は無傷）のため、デフォルトではサニタイズしない。

サニタイズ処理: 該当行を `' [REMOVED by sanitize] ApiName -- original had N chars` に置換。EDR はコメント内の `Declare ... Lib` もパターンマッチしてファイルを破壊するため、元のキーワードは一切残さない。API名と元の行長だけ記録し、何が除去されたか追跡可能にする。p-code もゼロ埋めする。

## CSV カラム

### 基本列

Timestamp, RelativePath, FileName, Bas, Cls, Frm, TotalModules, CodeLines, EdrIssues, CompatIssues, SanitizedLines, References, Error

### 追加列

| 列名 | 説明 |
|------|------|
| EnvIssues | 環境依存パターン検知数（Risk + Review） |
| EnvRisk | Risk severity の件数 |
| EnvReview | Review severity の件数 |
| EnvInfo | Info severity の件数 |
| BizIssues | 業務依存パターン検知数 |
| InfoCount | 全 Info 件数 |
| RiskLevel | 技術危険度 (High/Medium/Low) |
| MigrationClass | 対応方針（複数値、セミコロン区切り） |
| PrimaryConcern | 主要懸念（重み付き、1つ） |
| NeedsReviewBy | 確認担当（複数値、セミコロン区切り） |
| TopApiNames | 検出 API 宣言名（先頭3件） |
| TopComProgIds | 検出 COM ProgID（先頭3件） |
| SampleEvidence | 最も重い検出の代表行 |

### RiskLevel（技術危険度）

| 値 | 条件 |
|---|---|
| High | GUI操作系 API (FindWindow, SendMessage, SetForegroundWindow 等) あり、Shell/process あり、または PowerShell/WScript あり |
| Medium | Win32 API あり（GUI操作系以外）、DAO あり、または環境依存 Risk が 3件以上 |
| Low | 上記に該当しない |

### MigrationClass（対応方針）

| 値 | 条件 |
|---|---|
| そのまま可 | 全カテゴリのリスクがすべて 0 |
| 軽微修正 | 互換性リスクのみ |
| 要代替実装 | Win32 API（GUI操作系以外）、または DAO → ADO 移行 |
| 再構築必要 | GUI操作系 API 依存、または Shell/PowerShell 依存（新環境でブロック確認済み） |
| 要保存先見直し | 固定パス/UNC パス/共有フォルダ依存 |

### PrimaryConcern（主要懸念）

重み順で最初にマッチしたものを1つ採用:

1. GUI操作系 API → `GUI`
2. Shell / PowerShell → `Process`
3. 保存先移行関連 → `StorageMigration`
4. DB 連携 → `DB`
5. COM / 外部連携 → `COM`
6. ネットワーク → `Network`
7. メール → `Mail`
8. ファイル I/O → `File`
9. その他 → `Other`

### NeedsReviewBy（確認担当）

| 値 | 条件 |
|---|---|
| Security | EDR リスクあり |
| Infra | 環境依存（パス、プリンタ、接続文字列） |
| DB | DAO / ADO / 接続文字列あり |
| BusinessOwner | 業務フロー影響（Outlook, 外部システム） |
| ClientPC | 印刷, PDF, ActiveX |
| Developer | 互換性リスクのみ（PtrSafe, DefType 等） |

## ツールチップの方針

ツールチップの内容は Probe の実測結果に基づく。推測や一般論ではなく、新環境で実際に確認された動作に合わせる。

重要: `ThisWorkbook.Path` は SharePoint/OneDrive 条件で URL を返す。代替手段として推奨しない。

## 設定ファイル

`config/analyze.json` にパターンごとの detect/sanitize 設定を保持。4セクション: edr, compat, env, biz。

設定 GUI（Analyze.bat を引数なしで実行）で WinForms ダイアログから編集可能。
