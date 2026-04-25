# 01-hardcoded-path-break

OneDrive / SharePoint のローカル同期前提で、古い VBA のパス決め打ちが壊れる例です。

このサンプルでは 2 つを見せます。

- 固定パス `C:\SharePoint\...` を前提にした参照が失敗する
- `ThisWorkbook.Path & "\\案件データ"` を前提にした走査が失敗する

生成物:

- `out\HardcodedPathBreakDemo.xlsm`
- `out\demo-data\...`

実行:

1. `generate-sample.bat`
2. 生成された `.xlsm` を開く
3. `RunHardcodedPathDemo` を実行する

期待結果:

- 固定パスは `Missing` になる
- `ThisWorkbook.Path` 隣接走査も `Missing` になる
- 実データは OneDrive 同期ルート配下に存在するが、古いロジックでは見つからない
