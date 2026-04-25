# CLAUDE.md - pub/vba-devkit

## プロジェクト概要

Excel / VBA 資産のコマンドライン開発キット。
`.xlsm` から VBA ソースを抽出、差分比較、環境情報テスト、秘匿化、保護解除の
一連操作を PowerShell から提供する。先生（dummy-org）の VBA 作業基盤。

## 技術スタック

- PowerShell 5.1（Windows 同梱版）+ `.ps1` / `.bat` ラッパ
- Excel 自動化: COM （Excel.Application）
- 配布: ディレクトリコピーで完結、git clone で使える

## ディレクトリ構成

```
vba-devkit/
├── *.bat                # エントリポイント（Analyze / Diff / EnvTest / Extract / Sanitize / Unlock）
├── lib/
│   ├── VBAToolkit.psm1  # 共通モジュール
│   ├── {Analyze|Diff|EnvTest|Extract|Sanitize|Unlock}.ps1
│   └── internal/
├── config/              # 各コマンドの設定テンプレ
├── demos/               # 動作確認用サンプル
├── samples/             # 典型的な入力例
├── test/                # 自動テスト
├── docs/                # 使い方詳細
└── env-test-results/    # EnvTest.bat の実行結果（**.gitignore 対象**）
```

## コマンド

```cmd
Extract.bat <path.xlsm>      # VBA ソース抽出 → .bas / .cls / .frm
Diff.bat <old.xlsm> <new>    # 2 ブックの VBA 差分
Analyze.bat <path.xlsm>      # VBA 構造・import・コールグラフ解析
EnvTest.bat                  # PC 固有情報の収集（テナント名・パス・レジストリ）
Sanitize.bat <path>          # 秘匿化（テナント名を *** に置換）
Unlock.bat <path.xlsm>       # VBA プロジェクト保護解除
```

## 設計原則

- **秘匿対応**: `EnvTest.bat` の出力と `env-test-results/` はテナント名等の環境依存文字列を含みうる
  - `.gitignore` で確実に除外
  - commit 前に `Sanitize.bat` を通すか、手動で `Mask-Path` ロジックを確認
  - 2026-04-18 に `OneDrive - {tenant}` 形式の漏洩を修正済（commit `e4c17ee`）
- **Windows 専用**: PowerShell 5.1 同梱 + Excel が入っている PC を前提
- **COM リーク回避**: Excel.Application は使い終わったら必ず Quit + `[GC]::Collect()` 相当の後処理
- **ビルド不要**: `.psm1` / `.ps1` を直接実行

## テスト

```cmd
test\Run-All.bat           # 全テスト
test\Test-Extract.ps1      # 個別
```

## 注意事項

- VBA 資産は `.xlsm` の中に完結、git 管理しやすい形 (.bas/.cls) に抽出してから管理する運用
- 秘匿化忘れ防止: git hooks（pre-commit）で `env-test-results/` のステージングをブロック推奨
- AI に操作させる場合、`Unlock.bat` は強力なので先生の明示指示を待つ

## 関連

- pub/casedesk — この devkit の主要ユーザー（ケース管理用 VBA アドイン）
- pub/watchbox — VBA ではなく PS+C# だが、同じ配布ノリ（launch.bat でダブルクリック起動）
