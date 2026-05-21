# CTU Toolkit

VBA マクロを Excel を起動せずに調査・解析・サニタイズするための公開用ツール集。`.xlsm` / `.xlam` / `.xls` 内の VBA コードを直接バイナリパース層で取り出し、抽出・解析・EDR トリガー行のマスク・パスワード解除を行う。

## 含まれるツール

| BAT | 用途 | Excel 起動 |
|---|---|---|
| `Analyze.bat` | VBA コードのリスク解析 (EDR / 互換性 / 環境依存 / 業務影響) + 移行ガイド + CSV | 不要 |
| `Extract.bat` | `.xlsm` 等から VBA ソース取り出し (`.bas` / `.cls` / `.frm` + 結合テキスト) | 不要 |
| `Sanitize.bat` | EDR トリガーになりやすい VBA 行をコメント化した別ファイルを出力 | 不要 |
| `Unlock.bat` | VBA プロジェクトのパスワード保護を解除 (非破壊、別ファイル出力) | `.xlsm` / `.xlam` の場合のみ必須 |

## 必要環境

- Windows 10 / 11
- PowerShell 5.1 以降 (Windows 標準同梱)

## 使い方

### 共通

- 各 `.bat` は引数なしで起動するか、対象ファイル / フォルダを D&D で投入する
- 結果は `output\<タイムスタンプ>_<コマンド名>\` 以下に書き出される (実行のたびに新規ディレクトリが作成されるため過去結果は保持される)

### Analyze.bat

VBA コードに含まれる以下 4 カテゴリのリスク要素を検出する。

- **EDR**: Win32 API / Shell / COM / WMI / DLL 呼び出し
- **互換性**: PtrSafe / DAO / DDE 等の旧仕様
- **環境依存**: パス解決 / SaveAs / Dir / AutoSave (3 段階 severity)
- **業務影響**: Outlook / Word / PDF / 外部 EXE 連携

出力: HTML ビューア (色分け表示)、元のソース + sanitized copy、CSV、移行ガイド。

起動モードは「引数なし (設定 GUI)」「ファイル D&D (単体解析)」「フォルダ D&D (再帰解析)」の 3 つ。

### Extract.bat

引数または D&D で渡された `.xlsm` 等から `.bas` / `.cls` / `.frm` を取り出し、`combined.txt` に全モジュール結合版も出力する。

### Sanitize.bat

`.xlsm` / `.xlam` / `.xls` のファイルまたはフォルダを渡すと、EDR トリガーになりやすい VBA 行をコメント化したコピーを `output\<タイムスタンプ>_sanitize\` に出力する。元ファイルは変更しない。

対象は以下のような、環境調査で BLOCKED 扱いになりやすい行に限定する。

- Win32 API `Declare` 行
- 宣言済み Win32 API の呼び出し行
- `Shell` / `WScript.Shell` / `cmd /c` などのプロセス起動
- `powershell` / `wscript` / `cscript` / `mshta` などのスクリプトホスト起動

出力にはサニタイズ済みコピー、サニタイズ前ソースを表示して対象行をハイライトする HTML レポート、`sanitize.csv` の一覧が含まれる。

### Unlock.bat

VBA プロジェクトのパスワード保護を解除する。元ファイルは保持し、解除版を別ファイルとして出力する非破壊操作。

**注意**: パスワード保護回避の性質を持つため、社内ポリシーや法的根拠を確認の上、明示的に必要なケースに限定して実行すること。

## 仕組み

`.xlsm` (ZIP パッケージ) → `vbaProject.bin` 抽出 → OLE2 複合ドキュメントとしてセクタチェーン解析 → 各モジュールストリームを MS-OVBA 仕様の独自圧縮から伸長、までを PowerShell + C# (Add-Type) で処理する。Excel プロセスを介在させないため、大量ファイルの一括解析が可能。

Excel が必要なのは `Unlock.bat` で `.xlsm` / `.xlam` を扱う場合のみ (内部で `.xls` 経由変換が走るため)。

## 参照仕様

- MS-OVBA (Office VBA File Format Structure) 2.4.1
- MS-CFB (Compound File Binary Format)
