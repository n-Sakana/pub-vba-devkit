$ErrorActionPreference = 'Stop'

$source = @"
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;

public static class Win32HelperServer
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }

    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowTextW(IntPtr hWnd, StringBuilder text, int maxCount);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassNameW(IntPtr hWnd, StringBuilder text, int maxCount);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

    public static void Run(string prefix)
    {
        HttpListener listener = new HttpListener();
        listener.Prefixes.Add(prefix);
        listener.Start();
        Console.WriteLine("Helper listening at " + prefix);
        while (true)
        {
            HttpListenerContext context = listener.GetContext();
            try
            {
                string path = context.Request.Url.AbsolutePath ?? "/";
                if (path.Equals("/health", StringComparison.OrdinalIgnoreCase))
                {
                    Write(context, 200, "text/plain; charset=utf-8", "ok");
                    continue;
                }
                if (path.Equals("/api/foreground", StringComparison.OrdinalIgnoreCase))
                {
                    Write(context, 200, "text/plain; charset=utf-8", GetForegroundSummary());
                    continue;
                }
                Write(context, 404, "text/plain; charset=utf-8", "not-found");
            }
            catch (Exception ex)
            {
                Write(context, 500, "text/plain; charset=utf-8", ex.GetType().Name + ":" + ex.Message);
            }
        }
    }

    static string GetForegroundSummary()
    {
        IntPtr hwnd = GetForegroundWindow();
        StringBuilder title = new StringBuilder(512);
        StringBuilder klass = new StringBuilder(256);
        GetWindowTextW(hwnd, title, title.Capacity);
        GetClassNameW(hwnd, klass, klass.Capacity);
        RECT rect;
        GetWindowRect(hwnd, out rect);
        int width = rect.Right - rect.Left;
        int height = rect.Bottom - rect.Top;
        return Sanitize(title.ToString()) + "|" + Sanitize(klass.ToString()) + "|" + width + "|" + height + "|" + rect.Left + "|" + rect.Top;
    }

    static string Sanitize(string value)
    {
        return (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Replace("|", "/");
    }

    static void Write(HttpListenerContext context, int status, string contentType, string body)
    {
        byte[] bytes = Encoding.UTF8.GetBytes(body ?? string.Empty);
        context.Response.StatusCode = status;
        context.Response.ContentType = contentType;
        context.Response.ContentLength64 = bytes.LongLength;
        using (Stream output = context.Response.OutputStream)
        {
            output.Write(bytes, 0, bytes.Length);
        }
    }
}
"@

Add-Type -TypeDefinition $source -Language CSharp
[Win32HelperServer]::Run('http://127.0.0.1:8765/')
