# 03-ps-csharp-ipc

VBA から PowerShell / C# を直接起動せず、外で先に helper を立ち上げて localhost HTTP で叩く例です。

構成:

- `start-helper.bat`: PowerShell 実行ポリシーを回避しつつ helper を起動
- `start-helper.ps1`: C# を `Add-Type` でコンパイルし、HTTP helper を待受
- `build-client-sample.bat`: VBA クライアント `.xlsm` を生成
- `build-client-sample.ps1`: Excel COM でクライアントブックを生成
- `vba/IpcHttpClientDemo.bas`: VBA クライアント本体

実行:

1. `start-helper.bat`
2. 別のコンソールで `build-client-sample.bat`
3. `out\IpcHttpClientDemo.xlsm` を開く
4. `RunVisualWin32Demo` を実行する

期待結果:

- helper が `http://127.0.0.1:8765/` で待受
- VBA は `MSXML2.XMLHTTP.6.0` で helper を叩く
- helper は Win32 API で前景ウィンドウ情報を取得
- Excel シート側で色付きパネルとメタ情報が更新される
