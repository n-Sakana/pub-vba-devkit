# AGENTS.md - pub/vba-devkit

Codex 用の入口メモです。repo 固有の詳細は [README.md](README.md) と [CLAUDE.md](CLAUDE.md)、横断トポロジは [fin/hub/ARCHITECTURE.md](../../fin/hub/ARCHITECTURE.md) を参照してください。

## 役割

- Excel / VBA 資産向けの CLI 開発キット
- 抽出、差分、解析、環境テスト、秘匿化、保護解除をまとめて提供
- Windows + Excel 前提のローカルツール群

## runtime / 接続

- runtime: local Windows only
- shell: PowerShell 5.1 + `*.bat` wrapper
- automation: Excel COM

## まず見る場所

- `lib/VBAToolkit.psm1` - 共通関数
- `lib/Extract.ps1`, `Diff.ps1`, `Analyze.ps1`, `Sanitize.ps1`, `Unlock.ps1`
- `config/` - 設定テンプレ
- `test/` - 自動テスト

## コマンド

```cmd
Extract.bat <path.xlsm>
Diff.bat <old.xlsm> <new.xlsm>
Analyze.bat <path.xlsm>
EnvTest.bat
Sanitize.bat <path>
Unlock.bat <path.xlsm>
test\Run-All.bat
```

## ガードレール

- `env-test-results/` は秘匿情報を含みうる。git に入れない
- COM オブジェクトは必ず解放する
- `Unlock.bat` は強力なので、明示指示なしに実行しない
- Windows 専用前提を崩さない
