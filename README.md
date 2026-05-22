# CTU Toolkit

CTU Toolkit は、Excel VBA マクロを Excel を起動せずに調査・抽出・比較・サニタイズするための Windows 向けツール集です。`.xlsm` / `.xlam` / `.xls` に含まれる VBA プロジェクトをファイル構造から読み取り、リスク解析、ソース抽出、VBA 差分比較、危険行の無効化コピー作成、VBA プロジェクト保護の解除を行います。

## 含まれるツール

| BAT | 用途 | Excel 起動 |
|---|---|---|
| `Analyze.bat` | VBA コードのリスク解析、HTML レポート、CSV 出力 | 不要 |
| `Extract.bat` | VBA ソースを `.bas` / `.cls` / `.frm` と結合テキストに抽出 | 不要 |
| `Diff.bat` | 2つのブックに含まれる VBA モジュール差分を HTML / text で出力 | 不要 |
| `Sanitize.bat` | 検出対象の危険な VBA 文を固定コメントへ置換したコピーを作成 | 不要 |
| `Unlock.bat` | VBA プロジェクトのパスワード保護を解除した別ファイルを作成 | `.xlsm` / `.xlam` の場合のみ必要 |

環境テスト系の機能は含めていません。

## 必要環境

- Windows 10 / 11
- Windows PowerShell 5.1 以降
- `.xlsm` / `.xlam` の `Unlock.bat` 実行時のみ Microsoft Excel

## 使い方

各 `.bat` に対象ファイルをドラッグ & ドロップしてください。`Analyze` / `Extract` / `Sanitize` はフォルダ指定にも対応します。`Diff` は比較する2ファイルを指定します。コマンドプロンプトから直接呼び出すこともできます。

```cmd
Analyze.bat  C:\path\to\book.xlsm
Extract.bat  C:\path\to\book.xlsm
Diff.bat     C:\path\to\before.xlsm C:\path\to\after.xlsm
Sanitize.bat C:\path\to\book.xlsm
Unlock.bat   C:\path\to\book.xlsm
```

フォルダ対応コマンドでは、対応する Excel ファイルを再帰的に処理します。結果は `output\<timestamp>_<command>\` 以下に作成され、元ファイルは上書きされません。

## Analyze.bat

VBA コードを読み取り、次のカテゴリで検出結果を出します。

- **EDR / Security**: Win32 API `Declare`、Shell、WMI、外部プロセス起動など
- **Compatibility**: `PtrSafe`、DAO、DDE、古いコントロールなど
- **Hardcoded Path**: 固定パス、`ThisWorkbook.Path`、`Dir()`、外部リンクなど

出力には HTML ビューア、テキストレポート、`analyze.csv` が含まれます。引数なしで起動すると、検出設定を調整する簡易 GUI が開きます。

## Extract.bat

VBA プロジェクトから各モジュールのソースを抽出します。

出力例:

```text
output\20260521_120000_extract\
├── modules\<workbook-name>\
│   ├── Module1.bas
│   ├── Class1.cls
│   └── UserForm1.frm
└── <workbook-name>_combined.txt
```

## Diff.bat

2つの Excel ブックから VBA モジュールを読み取り、追加・削除・変更されたモジュールと行単位の差分を出力します。Excel は起動しません。

```cmd
Diff.bat C:\path\to\before.xlsm C:\path\to\after.xlsm
```

出力には以下が含まれます。

- `diff.html`: 左右比較できる HTML レポート
- `diff.txt`: 追加・削除・変更モジュールの概要

## Sanitize.bat

検出対象の危険な VBA 文を壊したコピーを作成します。元ファイルは変更しません。

対象例:

- Win32 API `Declare` 行
- 宣言済み API の呼び出し行
- 宣言済み API を呼ぶラッパー `Sub` / `Function` と、その呼び出し行
- `AddressOf` で危険行へ渡されるコールバック関数名
- `Shell` / `WScript.Shell` / `cmd /c` などのプロセス起動
- `powershell` / `pwsh` / `wscript` / `cscript` / `mshta` などのスクリプトホスト起動

置換後の VBA 行は、内容や種類を残さず、次の固定コメントだけになります。

```vb
' sanitized: *****
```

複数行継続文も、対象文全体を同じ固定コメントに置換します。サニタイズ後のブックはマクロが壊れる前提です。目的はマクロ動作の維持ではなく、ブック構造・ワークシート・セルデータを取り出せる状態へ近づけることです。

出力には以下が含まれます。

- `<name>_sanitized.<ext>`: 危険行を固定コメント化したコピー
- `<name>_sanitized.html`: サニタイズ前の元コードを表示し、置換対象行をハイライトする HTML レポート
- `sanitize.csv`: 処理結果の一覧

HTML レポートには元コードが表示されます。公開・共有する場合は、レポート内に機密情報が含まれないか確認してください。

## Unlock.bat

VBA プロジェクトのパスワード保護を解除したコピーを出力します。元ファイルは保持されます。

この機能は、正当な権限を持つファイルに対してのみ使用してください。

## ディレクトリ構成

```text
CTU Toolkit/
├── Analyze.bat
├── Extract.bat
├── Diff.bat
├── Sanitize.bat
├── Unlock.bat
├── config/
│   └── analyze.json
└── lib/
    ├── Analyze.ps1
    ├── Extract.ps1
    ├── Diff.ps1
    ├── Sanitize.ps1
    ├── Unlock.ps1
    └── VBAToolkit.psm1
```

## 注意

- 不審な Excel ファイルは、業務端末ではなく隔離された環境で扱ってください。
- `Sanitize.bat` はマクロを安全に実行できるようにするツールではありません。検出対象の文を破壊し、調査・救出用のコピーを作るためのツールです。
- HTML レポートはサニタイズ前のコードを含みます。外部共有時はレポートを同梱しない、または内容を確認してください。
