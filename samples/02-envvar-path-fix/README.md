# 02-envvar-path-fix

`OneDriveCommercial` / `OneDrive` を使って同期ルートを解決し、OneDrive / SharePoint ローカル同期上でも安定して対象フォルダへ届く例です。

生成物:

- `out\EnvVarPathFixDemo.xlsm`
- `out\demo-data\...`

実行:

1. `generate-sample.bat`
2. 生成された `.xlsm` を開く
3. `RunEnvVarPathDemo` を実行する

期待結果:

- 同期ルートが環境変数から解決される
- SharePoint 相当のローカルパス上に置いたサンプルファイルを列挙できる
- `ThisWorkbook.Path` に依存しない
