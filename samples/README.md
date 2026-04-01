# samples

この配下は Windows 側で実行する生成型サンプルです。

- `01-hardcoded-path-break`: 固定パス / `ThisWorkbook.Path` 前提が OneDrive 同期で壊れる
- `02-envvar-path-fix`: 環境変数で同期ルートを解決して回復する
- `03-ps-csharp-ipc`: helper 常駐 + VBA HTTP クライアントで Win32 を外へ逃がす

`.xlsm` はリポジトリに直置きしていません。環境依存が強いので、各 `bat` からその場で生成します。
