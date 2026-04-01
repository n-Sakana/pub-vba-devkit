# 新環境の制約まとめ

更新日: 2026-03-30
対象: analyze リライトと移行方針の前提整理

## 目的

この文書は、2026-03-30 時点の最新 Probe / Survey 結果と、2026-03-29 の analyze.csv 集計をもとに、新環境で前提にしてよいこと・捨てるべきことを一枚にまとめたものです。

見るべき生データは主に以下です。

- `probe_Local_20260330_020742.txt`
- `probe_OneDrive_Synced_20260330_020158.txt`
- `survey_Local_20260330_020742.txt`
- `survey_OneDrive_Synced_20260330_020158.txt`
- `analyze.csv`（2026-03-29 22:31-22:32 集計）

## 結論

新環境の改修方針は、端的に言えばこれです。

- VBA は薄い窓口として残す
- Win32 API、外部プロセス起動、GUI 自動操作は VBA から外へ出す
- パスは固定文字列や `ThisWorkbook.Path` 前提で決めず、環境から解決する
- helper は「必要時起動」ではなく「先に起動して待ち受ける」前提で考える
- analyze は「危険 API」だけでなく「保存先・環境依存」を主戦場として再設計する

## 1. 確定した制約

### 1-1. `Declare` を含む VBA ブックは EDR で壊れる

2026-03-30 の Local / OneDrive_Synced の両方で、`Declare PtrSafe Function` は `BLOCKED` でした。
しかも単なる実行失敗ではなく、保存・再オープン時に「ファイル形式またはファイル拡張子が正しくありません」となっています。

意味は明快です。

- Win32 API を VBA から直接呼ぶ案は捨てる
- `Declare only` でもだめ
- sanitize の最優先対象は引き続き `Declare`

これは実行時のブロックではなく、ファイルレベルの破壊です。最優先で避けるべき棚です。

### 1-2. VBA からの外部プロセス起動は前提にできない

2026-03-30 の Extended Probe では、以下の差が出ています。

- `CreateObject("WScript.Shell")` は OK
- `WScript.Shell cmd` は FAIL
- `powershell via VBA` は FAIL

失敗内容はどちらも `Run:書き込みできません。` です。

つまり、

- COM オブジェクト生成だけなら通る場合がある
- しかし `Run` で cmd / PowerShell を起動する前提は危険
- 「必要時に VBA から helper を起動する」設計は避ける

helper を使うなら、別の手段で先に起動しておく前提が必要です。

### 1-3. 64bit / 廃止系の互換性欠落は現実に発生している

2026-03-30 の Probe で確認できた失敗は以下です。

- `DAO.DBEngine.36` → FAIL（クラス未登録）
- `MSComDlg.CommonDialog` → FAIL
- `MSCAL.Calendar` → FAIL
- `DDEInitiate` → FAIL

意味はこうです。

- DAO 3.6 延命は前提にしない
- 古い ActiveX コントロール置換は必要
- DDE 前提も外す

一方で `InternetExplorer.Application` は OK でしたが、これは「今この端末で見えている」だけです。新設計の前提にはしないほうがよいです。

## 2. 通っているもの

2026-03-30 の Probe / Survey では、以下は比較的安定していました。

- `Scripting.FileSystemObject`
- `Scripting.Dictionary`
- `ADODB.Connection`
- `ADODB.Recordset`
- `MSXML2.XMLHTTP.6.0`
- `MSXML2.DOMDocument.6.0`
- `WinHttp.WinHttpRequest.5.1`
- `Shell.Application`
- `GetObject("winmgmts")`
- `Open/Write/Kill` などの VBA 標準 File I/O
- `GetSetting`
- `Environ$`
- `MSForms.DataObject`
- VBA baseline

Survey 側でも以下が確認できています。

- Windows PowerShell 5.1 あり
- .NET Framework 4.8.1 あり
- `Add-Type` による C# コンパイル可
- Windows Forms / WPF / UIAutomation 可
- Office は Excel / Word / Outlook / Access / PowerPoint が導入済み
- Acrobat COM あり

要するに、新環境は「VBA 単独の危険棚は閉じるが、PowerShell + C# helper と標準 COM はかなり使える」環境です。

## 3. パスと保存先の制約

### 3-1. 最新結果だけ見ると、OneDrive 同期フォルダをローカルパスで開いた場合は `ThisWorkbook.Path` は Local

2026-03-30 の `probe_Local` と `probe_OneDrive_Synced` では、両方とも以下でした。

- `ThisWorkbook.Path` → Local
- `ThisWorkbook.FullName` → Local
- `Path Type` → Local
- `Dir(ThisWorkbook.Path)` → OK
- `Workbooks.Open`（隣接ファイル）→ OK
- `SaveAs to temp` → OK

この観測だけを見ると、少なくとも「OneDrive 同期済みローカルフォルダを普通に開いたケース」では、URL 化は起きていません。

### 3-2. ただし、それで `ThisWorkbook.Path` 前提へ戻してはいけない

2026-03-28 の SharePoint / Open-in-App 調査では、`ThisWorkbook.Path` と `FullName` が URL になるケースがすでに出ています。

つまり現時点の理解はこうです。

- ローカル同期済みファイルをローカルパスで開けば Local になるケースがある
- SharePoint / Open-in-App 系では URL になるケースがある
- 利用者の開き方まで制御できないなら、`ThisWorkbook.Path` を業務ルートの根拠にしないほうが安全

このため、改修方針としては引き続き以下を推奨します。

- 共有業務フォルダは `Environ$("OneDriveCommercial")` / `Environ$("OneDrive")` から同期ルートを解決する
- コードには共通の相対パスだけを残す
- `ThisWorkbook.Path` は「今開いている場所の情報」としては見ても、「保存先の基準」としては見ない

### 3-3. 今回の Probe では OneDrive 環境変数そのものは Host 側で確認できている

2026-03-30 の Probe では、Basic 側の OneDrive 系テストが一連で FAIL していますが、Extended Host 側では以下が OK です。

- `OneDriveCommercial` → Exists=True
- `OneDrive` → Exists=True
- `Preferred sync root` → OK
- `Get-ChildItem` → OK
- `.NET Directory.GetFiles` → OK

したがって、今回の FAIL は「OneDrive 環境変数が無い」より、「Basic 側のマクロ実行フローに不調がある」可能性が高いです。

分析用の前提としては、

- Host 側では同期ルート取得は現実的
- VBA 側の直接テストは再確認が必要

と置くのが妥当です。

## 4. IPC / helper まわりの制約

### 4-1. Host 側の Win32 / .NET は使える

2026-03-30 の Host テストでは以下が通っています。

- `PowerShell / C# Add-Type` → OK
- `GetTickCount64` → OK
- `GetCurrentThreadId` → OK
- `MessageBeep` → OK
- Window handle lookup / metadata / `SetForegroundWindow` / `ShowWindow` / `SendMessage` / `PostMessage` → OK
- UIAutomation → OK

つまり、Win32 API を使う必要があるなら、VBA ではなく helper 側へ寄せるのが筋です。

### 4-2. 通信方式はまだ「完全確定」ではない

2026-03-30 の Host 側結果はこうです。

- NamedPipe server 作成 → OK
- NamedPipe roundtrip → TIMEOUT
- TCP listener → OK
- TCP loopback roundtrip → OK
- HttpListener 起動 → OK
- HttpListener loopback → TIMEOUT

この結果から言えるのは、

- TCP は少なくとも Host 単体 roundtrip では通った
- Named Pipe と HttpListener は「作成はできるが往復成功は未確認」
- 本命候補を決め切るには再試験が必要

ただし VBA 側との相性まで考えると、なおも HTTP は有力です。
理由は、VBA 側で `XMLHTTP` / `WinHttpRequest` が通っているからです。

なので現時点の整理はこうです。

- 実測 green: TCP
- 実装都合の本命候補: localhost HTTP
- ただし HttpListener の往復 timeout 原因は切り分け要
- Named Pipe は VBA 主回線としては優先度低め

## 5. analyze.csv から見える主戦場

2026-03-29 の `analyze.csv` 集計では、件数として最も多いのは `StorageReview` と `RefactorPathHandling` です。
一方で、濃い危険棚として以下が見えています。

- `WScript.Shell`
- `keybd_event`
- `Sleep`
- `SHCreateDirectoryEx`
- DAO 系

意味はこうです。

- 件数ベースの本命はストレージ移行
- 技術的に危険なのは Win32 / process / GUI 自動操作
- analyze の書き換えでは「Path の使い方」と「VBA から外へ出すべき処理」の二軸を前面に出すべき

特に `MigrationClass` で目立つのは以下です。

- `StorageReview`
- `RefactorPathHandling`
- `NeedsReplacement`
- `Rebuild`

このため、新 analyze の出力は「危険 API の有無」だけでなく、少なくとも次を一目で分ける必要があります。

- ストレージ共通化で救えるもの
- helper へ逃がすべきもの
- GUI / process で再構築が必要なもの
- 互換性差分だけで済むもの

## 6. 改修方針

### 6-1. VBA 側に残すもの

- 標準的なファイル I/O
- FSO / Dictionary / ADODB / XMLHTTP など既知で通る COM
- シート操作、帳票操作、軽い業務ロジック
- 環境変数ベースのパス解決

### 6-2. VBA から外へ出すもの

- `Declare` が必要な Win32 API
- `WScript.Shell.Run` / PowerShell 起動前提の処理
- GUI 自動操作
- 不安定な待機やウィンドウ制御
- 64bit 非対応の古いコントロール依存

### 6-3. helper 側の前提

- PowerShell 5.1 + .NET Framework 4.8.1 前提で組める
- C# は `Add-Type` でも専用ビルドでもよい
- helper は先行起動型を基本とする
- localhost 通信は HTTP を第一候補、TCP を保険として扱う

### 6-4. ストレージの前提

- 固定ローカルパスは共有しない
- `ThisWorkbook.Path` を共有業務パスの根拠にしない
- 共有業務フォルダは同期ルート + 相対パスで解決する
- Save/Open の扱いは「今どの開き方で起きた挙動か」を区別する

## 7. analyze リライトで落とし込むべき判定軸

新 analyze では、少なくとも以下を主軸に据えるべきです。

### A. ファイルレベルで即除外すべきもの

- `Declare`
- GUI 系 Win32 API
- `WScript.Shell.Run` / PowerShell 起動依存

### B. ストレージ見直しでまとめて救えるもの

- 固定パス
- `Dir(...)` ベースの存在確認
- `SaveAs`
- `CurDir`
- `ThisWorkbook.Path` とその連結
- OneDrive / SharePoint 前提がにじむコード

### C. 互換性置換で処理できるもの

- DAO 3.6
- CommonDialog / Calendar
- DDE

### D. そのまま残しやすいもの

- ADODB
- XMLHTTP / WinHttp
- FSO / Dictionary
- ふつうの VBA 処理

## 8. 未解決事項

2026-03-30 時点で、まだ確定しきっていない点もあります。

- HttpListener loopback timeout の原因
- NamedPipe roundtrip timeout の原因
- Basic 側 OneDrive environment テスト失敗の原因
- SharePoint / Open-in-App を含めたパス挙動の再確認
- AutoSave=True 環境での BeforeSave / AfterSave 実際の発火挙動

このため、「helper は HTTP で確定」とまではまだ書かないほうが正確です。
現時点では「HTTP 第一候補、TCP 実証済み、Named Pipe は保留」が妥当です。

## 9. 実務上の一文要約

新環境は、VBA を中心に何でもやる環境ではありません。

- VBA は残せるが、危険処理は持たせない
- Win32 と外部実行は helper 側へ逃がす
- ストレージ移行は同期ルート解決を共通部品にする
- analyze は API 検知ツールではなく、「どこまで VBA に残せるか」を選別する道具として書き直す

これが 2026-03-30 時点の最新調査結果から引ける前提です。
