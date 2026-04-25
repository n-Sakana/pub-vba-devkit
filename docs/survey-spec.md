# Survey 仕様書

## 概要

`Survey.bat` は、端末に追加インストールを行わずに、開発・配布の前提となる環境情報を棚卸しするツール。

`Probe` と違い、VBA コード注入や EDR 実測は行わない。目的は「この PC で何が入っていて、何を前提にできるか」を固定情報として取得することにある。

## 取得対象

### 1. マシンスペック

- PC 名
- メーカー / モデル
- BIOS バージョン
- OS 名 / バージョン / ビルド / アーキテクチャ
- CPU 名 / コア数 / スレッド数
- 物理メモリ容量
- GPU 名
- 固定ドライブ容量 / 空き容量

### 2. ランタイム / 言語

- Windows PowerShell 5.1
- PowerShell 7
- .NET Framework 4.x
- `dotnet` ランタイム / SDK
- C# `Add-Type`
- Python
- `py` launcher
- Node.js
- Java
- Windows Script Host (`cscript.exe`, `wscript.exe`)

### 3. Office ホスト

- Excel
- Word
- Outlook
- Access

確認内容:
- 実行ファイルの所在
- ファイルバージョン
- COM automation の ProgID 登録有無

### 4. PDF / Acrobat

- Adobe Acrobat / Acrobat Pro のインストール情報
- `Acrobat.exe` の所在
- 実行ファイルバージョン
- COM automation 登録:
  - `AcroExch.App`
  - `AcroExch.AVDoc`
  - `AcroExch.PDDoc`

## 出力

出力先:

```text
output/<timestamp>_survey/
  survey.txt
  survey.json
```

- `survey.txt`: 人が読む用
- `survey.json`: 後続処理・比較用

個人情報保護:
- PC 名はマスクする
- パスは `Probe` と同様にマスクして出力する
- 取得処理の内部では実パスを参照しても、レポートにはマスク済み値だけを書き出す

## 非対象

このツールは次を行わない。

- EDR ブロックの実測
- VBA コード注入
- 外部アプリの新規インストール
- 端末設定の変更
- Acrobat API の機能テスト

Acrobat については「登録されている automation 入口」を棚卸しするのみで、PDF を実際に開いて API 呼び出しを試すところまでは行わない。
