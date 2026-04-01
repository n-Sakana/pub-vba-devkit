# Probe 仕様書

## 概要

`Probe.bat` は PowerShell + Excel COM を使い、新環境での EDR リスク・互換性リスク・ストレージ境界を自動テストするツール。VBA コードを一時 xlsm に注入（inject）して実行し、OK/FAIL を記録する。

## アーキテクチャ: 二段プローブ構成

### メインプローブ (`Probe.bat B` / `Probe.bat E`)

PowerShell から Excel COM 経由で一時 xlsm を作成し、VBA コードを注入→保存→再オープン→マクロ実行。テスト完了後に一時ファイルを削除する。

テスト対象:
- **EDR 検知**: Win32 API Declare、COM CreateObject（FSO, ADODB, XMLHTTP, Shell 等）、ファイル I/O、レジストリ、クリップボード、SendKeys
- **互換性**: DAO, Legacy Controls, IE Automation, DDE
- **境界テスト**: Declare 有無（呼び出さない場合）、COM breadth、WScript.Shell 分離テスト
- **情報収集**: Office バージョン、参照ライブラリ、パス情報

重要: メインプローブの Storage テストは **一時ファイルのパス（ローカル temp）** を観測する。SharePoint/OneDrive の実パスではない。

### ストレージプローブ (`Probe.bat G`)

`probe_storage.xlsm` を生成する。このファイルを手動で SharePoint にアップロードし「アプリで開く」で実行することで、SharePoint/OneDrive 環境での実際のパス・保存・オープン動作をテストする。

テスト対象:
- ThisWorkbook.Path / FullName（マスク済み出力）
- CurDir
- パスタイプ判定（Local / URL / UNC / OneDrive）
- Dir() による相対パス解決
- AutoSave 状態
- Workbooks.Open（隣接ファイル作成→オープン）
- SaveAs（temp へ保存）
- COM オブジェクト生成（FSO, Dictionary, ADODB, XMLHTTP）

## 3つのモード

| モード | コマンド | 説明 |
|--------|----------|------|
| **B** (Basic) | `Probe.bat` → `B` | EDR + 互換性 + 境界テスト（安全プローブ）|
| **E** (Extended) | `Probe.bat` → `E` | Basic + Shell/PowerShell/WMI/DDE/SendKeys（強いプローブ）|
| **G** (Generate) | `Probe.bat` → `G` | `probe_storage.xlsm` を output/ に生成 |

## 実装方式

### Test-VbaCode 関数

メインプローブの中核。以下のフローで各パターンをテスト:

```
1. Excel COM で空の xlsm を作成
2. VBProject にモジュールを追加し、VBA コードを注入
3. xlsm として保存（EDR がここでブロックする場合がある）
4. 閉じて再オープン（EDR がファイル破損させる場合がある）
5. マクロを実行し、戻り値で OK/FAIL を判定
6. 一時ファイルを削除
```

### パスマスキング

プライバシー保護のため、全てのパス出力は `MaskPath` 関数でマスクされる:

- `C:\Users\xxx\Documents\folder` → `Local:C:\***\***\***\***(depth=4)`
- `https://company.sharepoint.com/sites/team` → `URL:https://***/***/***/***(depth=4)`
- `\\server\share\folder` → `UNC:\\***\***\***(depth=3)`

ドライブレター/プロトコル、深さ、タイプは保持し、ディレクトリ名は隠蔽する。

## テスト項目

### Basic（安全プローブ）

| カテゴリ | パターン | テスト内容 |
|----------|----------|------------|
| EDR | Win32 API (Declare) | `Declare PtrSafe Function GetTickCount` |
| EDR | COM / CreateObject | FSO, Dictionary, ADODB, XMLHTTP, WinHttp, Shell, DOMDocument |
| EDR | File I/O | Open/Write/Kill |
| EDR | FileSystemObject | FSO methods |
| EDR | Registry | GetSetting |
| EDR | Environment | Environ$ |
| EDR | Clipboard | MSForms.DataObject |
| EDR | VBA Baseline | Pure VBA（外部依存なし） |
| Compat | Deprecated: DAO | DAO.DBEngine.36 |
| Compat | Deprecated: Legacy Controls | MSComDlg, MSCAL |
| Compat | Deprecated: IE Automation | InternetExplorer.Application |
| Info | ThisWorkbook.Path | パス値（マスク済み） |
| Info | ThisWorkbook.FullName | フルパス値（マスク済み） |
| Info | Workbooks.Open | 隣接ファイル生成→オープン |
| Info | SaveAs | temp へ保存（パスマスク済み） |
| Storage | CurDir | カレントディレクトリ（マスク済み） |
| Storage | Relative Path | Dir(ThisWorkbook.Path) |
| Storage | AutoSave | AutoSaveOn 状態 |
| Storage | Path Type | Local / URL / UNC / OneDrive 判定 |
| Reference | VBProject.References | 参照一覧 + Missing 確認 |

### Extended（強いプローブ）

| カテゴリ | パターン | テスト内容 |
|----------|----------|------------|
| EDR | Shell / process | WScript.Shell.Run cmd |
| EDR | PowerShell / WScript | powershell via WScript.Shell |
| EDR | Process / WMI | GetObject winmgmts |
| EDR | SendKeys | SendKeys "" |
| Compat | Deprecated: DDE | DDEInitiate |

## 出力

### メインプローブ結果

ファイル: `output/probe_result_<日時>.txt` (BOM UTF-8)

```
Level  Category  Pattern  Target  Result  Phase  ErrMsg  Detail
```

### ストレージプローブ結果

ファイル: `probe_storage_<scenario>_<日時>.txt`（xlsm と同じフォルダまたは temp）

```
TestName  Result  Detail
```

## toolkit との対応

メインプローブの結果と Analyze の検知結果を突き合わせることで:

- toolkit が「危ない」と言ったものが新環境で**本当に動かないのか**を確認
- FAIL パターンを使っているファイルを Analyze CSV から抽出 → 修正対象の確定

ストレージプローブは SharePoint/OneDrive 固有の問題（URL パス、AutoSave、相対パス解決の違い）を検証する。
