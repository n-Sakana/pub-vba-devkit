# Session Summary — 2026-03-30

## What was done

### EnvTest / Probe

1. **Probe セクション整理** (`84ad884`)
   - 旧: Core Automation / Boundary / Legacy が混在
   - 新: VBA Baseline → EDR/Security → Workbook Context → Legacy → Storage/Path
   - 出力順序 = 実行順序のルールは維持

2. **90秒/項目タイムアウト** (`84ad884`, `f2a515c`, `f4f0f1d`)
   - `Test-VbaCode`: 各フェーズ間で cooperative タイムアウト + Excel PID 追跡で強制終了
   - `Test-HostAction`: 別ランスペース + `BeginInvoke` + `WaitOne(90s)` でプリエンプティブタイムアウト
   - `$ps.Stop()` → `$ps.BeginStop()` に変更（ネイティブブロッキング呼び出しで Stop が帰らない問題）

3. **Ctrl+C スキップ** (`f4f0f1d`)
   - `[Console]::TreatControlCAsInput = $true` でキー入力として捕捉
   - Ctrl+C = 現在の項目をスキップ (SKIP)、Ctrl+Esc = 全体中断
   - `Test-HostAction`: 500ms ポーリングループでキー入力チェック
   - `Test-VbaCode`: フェーズ間で `OperationCanceledException`

4. **OneDrive 空白パス対応** (`35c1159`)
   - `Invoke-DevkitScript` の `Start-Process -ArgumentList` でパス値をクォート
   - OneDrive 同期パス (`OneDrive - 会社名`) のスペースで引数が切れていた

5. **出力統合** (`e1d60e0`, `abdb5c6`, `64a0ad2`)
   - survey.txt + probe.txt + envtest.txt → 単一の `report.txt` に統合
   - ファイル名に シナリオ + タイムスタンプ: `report_Local_20260330_143000.txt`
   - survey.json 削除

6. **新テスト項目** (`f670dd5`, `98dfb56`)
   - 最小 `Environ$` テスト (OneDriveCommercial / OneDrive / OneDriveConsumer)
     - `$storageVbaHelpers` を使わず6行のVBAのみ → EDR がコードサイズで弾いているか内容で弾いているかの切り分け
   - クロスプロセス HTTP ラウンドトリップテスト (VBA → PS バックエンド)
     - `Start-Job` で別プロセスに HttpListener 起動、VBA が XMLHTTP で POST

### Analyze

7. **判断基準の更新** (`ed723b2`, `f2837ae`, `455cb67`)
   - RiskLevel: `Shell / process` を High に追加（VBA Shell.Run はブロック確認済み）
   - guiApiNames: SetForegroundWindow, ShowWindow 等を追加
   - localhost: Risk → Review に降格（バックエンド通信で使用）
   - detect デフォルト: probe で安全確認済みの項目を false に（COM, File I/O, FSO, Registry 等）
   - DLL loading: sanitize を false に（Declare と重複）
   - MoveToBackend → Rebuild に戻した（方針は未決定、事実のみ記載）

8. **サニタイズ修正** (`acaf053`, `4f65a1e`)
   - コメントアウト方式 (`' [EDR] Declare ...`) → stub 方式 (`' [REMOVED by sanitize] ApiName -- original had N chars`)
     - EDR はコメント内の Declare/Lib もパターンマッチする
   - API 呼び出し行のサニタイズを廃止（Declare 文のみ対象に）
     - 旧: hits=23 なのに sanitized=1200 → 新: hits=sanitized
   - p-code は完全保持、圧縮ソースのみ差し替え
     - p-code ゼロ埋め → ファイル破壊（Test C で確認）
     - p-code 保持 → ファイル正常（Test B で確認）
   - HTML レポートには元コード表示（`OriginalLines` で保持）
   - フォルダ走査: `-Include` → `Where-Object Extension` に変更（.bat/.ps1 を処理していた）

### Demos

9. **3つのデモ** (`e75c3a2` / `1a26020`)
   - `01_path_problem.bat`: ThisWorkbook.Path / Dir() / FSO の失敗デモ
   - `02_environ_solution.bat`: Environ$ でパス解決するデモ
   - `03_http_backend.bat`: PS+C# HTTP サーバー + VBA クライアント
     - Demo_Win32Api: 画面解像度、カーソル位置、前面ウィンドウ、稼働時間（虹色バー付き）
     - Demo_Wiggle: SetWindowPos で Excel ウィンドウを揺らす

## Probe 実環境結果 (report4)

### 確認済みの事実

| パターン | VBA 内 | ホスト側 (PS) |
|---------|--------|-------------|
| 基本 VBA / COM (FSO, Dict, ADODB, XMLHTTP) | OK | - |
| File I/O, Registry, Environ$, Clipboard | OK | - |
| Win32 API Declare | BLOCKED (ファイル破壊) | OK (PInvoke) |
| WScript.Shell.Run / PowerShell 起動 | BLOCKED | OK |
| DDE, AppActivate (VBA版) | FAIL | - |
| DAO, CommonDialog, Calendar | 未登録 | - |
| OneDrive パス操作 ($storageVbaHelpers 使用) | FAIL (マクロ実行不可) | OK |
| XMLHTTP / WinHttp | OK | - |
| HttpListener (サーバー起動) | - | OK |
| NamedPipe / HttpListener ラウンドトリップ | - | TIMEOUT (90s) |
| PInvoke (user32/kernel32) | - | OK |
| UIAutomation | - | OK |

### 未検証

- 最小 Environ$ テスト (OneDriveCommercial) → EDR トリガーの切り分け
- クロスプロセス HTTP ラウンドトリップ → casedesk アーキテクチャの実現可能性
- サニタイズ済み xlsm の EDR 環境での動作

## 未解決の課題

### 1. サニタイズ済みファイルの VBE 表示

**状況**: ソースは書き換わっている（バイナリ抽出で確認済み）が、VBE が p-code キャッシュから古い Declare 文を表示する。

**試したこと**:
- p-code ゼロ埋め → ファイル破壊 (Test C)
- p-code ヘッダーのみ保持 + 本体ゼロ → ファイル破壊
- `_VBA_PROJECT` ストリームのバージョン変更 → Excel クラッシュ
- p-code バージョンバイト変更 → 効果なし

**事実**:
- p-code 内に `Declare`, `kernel32`, `advapi32`, `PtrSafe` は**存在しない**（API名のみ）
- EDR は p-code ではなくソーステキストでパターンマッチしている
- p-code を一切触らなければファイルは正常 (Test B)

**方針案**:
- 現状: ソースは確実にサニタイズ済み。EDR 環境ではソースが評価対象なので実用上問題ない可能性
- 要検証: EDR 環境でサニタイズ済みファイルが実際に開けるか

### 2. NamedPipe / HttpListener ラウンドトリップのタイムアウト

**状況**: 同一プロセス内のラウンドトリップテストが 90 秒でタイムアウト。別プロセス間（クロスプロセス HTTP）は未検証。

**原因候補**: EDR がループバック通信をブロック、または同一プロセス内のデッドロック。

### 3. OneDrive パス操作の VBA 側 FAIL

**状況**: `$storageVbaHelpers`（80行超の VBA コードブロック）を使うテストが全滅。最小 `Environ$` テスト（6行）を追加済みだが未実行。

**切り分け**: 最小テストが OK なら、大きなコードブロックがEDR を誘発。FAIL なら `OneDrive` 文字列自体がブロック対象。
