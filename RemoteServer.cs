using System;
using System.IO;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTWebBrowserAddIn
{
    public class RemoteServer
    {
        private TcpListener _listener;
        private Thread _serverThread;
        private bool _isRunning = false;
        private int _port;
        private SynchronizationContext _syncContext;

        public event Action<string> StatusChanged;
        
        private string _serverUrl;
        public string ServerUrl 
        { 
            get { return _serverUrl; }
            private set { _serverUrl = value; } 
        }

        public static string CurrentProxyTarget { get; set; }
        public static string CurrentCastImage = "";
        public static int CastVersion = 0;
        public static string CastMode = "none"; // none, photo, stream, mirror
        public static string CastStreamUrl = "";
        public static int CastRotation = 0; // 0, 90, 180, 270

        private static int _lastSlideId = -1;
        private static HashSet<string> _playedVideoShapeNames = new HashSet<string>();
        private static string _activeVideoShapeName = null;
        private static Dictionary<string, bool> _videoPlayingStates = new Dictionary<string, bool>();

        public static void ResetActiveVideoFocus()
        {
            _activeVideoShapeName = null;
            _videoPlayingStates.Clear();
            try { Globals.ThisAddIn.LogToFile("ControlVideo: Reset active video focus."); } catch {}
        }

        public RemoteServer(int port, SynchronizationContext syncContext)
        {
            _port = port;
            _syncContext = syncContext;
        }

        public void Start()
        {
            if (_isRunning) return;

            try
            {
                // Auto detect active LAN IPv4
                string localIp = null;
                var nics = NetworkInterface.GetAllNetworkInterfaces();
                foreach (var nic in nics)
                {
                    if (nic.OperationalStatus == OperationalStatus.Up && 
                        nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                    {
                        var ipProps = nic.GetIPProperties();
                        foreach (var ip in ipProps.UnicastAddresses)
                        {
                            if (ip.Address.AddressFamily == AddressFamily.InterNetwork && 
                                !IPAddress.IsLoopback(ip.Address))
                            {
                                localIp = ip.Address.ToString();
                                break;
                            }
                        }
                        if (localIp != null) break;
                    }
                }
                if (string.IsNullOrEmpty(localIp))
                {
                    localIp = "127.0.0.1";
                }
                
                // Bind to standard TCP socket
                _listener = new TcpListener(IPAddress.Any, _port);
                _listener.Start();
                
                ServerUrl = string.Format("http://{0}:{1}/", localIp, _port);
                _isRunning = true;

                _serverThread = new Thread(ListenLoop);
                _serverThread.IsBackground = true;
                _serverThread.Start();

                if (StatusChanged != null)
                {
                    StatusChanged("运行中: " + ServerUrl);
                }
            }
            catch (Exception ex)
            {
                if (StatusChanged != null)
                {
                    StatusChanged("启动失败: " + ex.Message);
                }
                MessageBox.Show("启动局域网遥控服务器失败，端口 " + _port + " 可能已被占用。\n详情: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _isRunning = false;
            }
        }

        public void Stop()
        {
            if (!_isRunning) return;

            try
            {
                SetWindowsSystemProxy("", false);
                _isRunning = false;
                if (_listener != null)
                {
                    _listener.Stop();
                }
                if (_serverThread != null)
                {
                    _serverThread.Join(500);
                }
                if (StatusChanged != null)
                {
                    StatusChanged("已停止");
                }
            }
            catch { }
        }

        private void ListenLoop()
        {
            while (_isRunning)
            {
                try
                {
                    var client = _listener.AcceptTcpClient();
                    ThreadPool.QueueUserWorkItem(new WaitCallback(HandleClient), client);
                }
                catch
                {
                    if (!_isRunning) break;
                }
            }
        }

        private void HandleClient(object state)
        {
            using (var client = (TcpClient)state)
            {
                try
                {
                    client.ReceiveTimeout = 6000;
                    client.SendTimeout = 6000;
                    
                    using (var stream = client.GetStream())
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        // Read request line
                        string requestLine = reader.ReadLine();
                        if (string.IsNullOrEmpty(requestLine)) return;

                        string[] parts = requestLine.Split(' ');
                        if (parts.Length < 2) return;
                        
                        string method = parts[0].ToUpper();
                        string path = parts[1].ToLower();
                        int queryIdx = path.IndexOf('?');
                        if (queryIdx >= 0)
                        {
                            path = path.Substring(0, queryIdx);
                        }

                        // Parse Content-Length and consume headers
                        int contentLength = 0;
                        string headerLine;
                        while (!string.IsNullOrEmpty(headerLine = reader.ReadLine()))
                        {
                            if (headerLine.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
                            {
                                int.TryParse(headerLine.Substring(15).Trim(), out contentLength);
                            }
                        }

                        if (method == "OPTIONS")
                        {
                            string optHeaders = "HTTP/1.1 204 No Content\r\n" +
                                                "Access-Control-Allow-Origin: *\r\n" +
                                                "Access-Control-Allow-Methods: GET, POST, OPTIONS\r\n" +
                                                "Access-Control-Allow-Headers: *\r\n" +
                                                "Content-Length: 0\r\n" +
                                                "Connection: close\r\n\r\n";
                            byte[] optHeaderBytes = Encoding.UTF8.GetBytes(optHeaders);
                            stream.Write(optHeaderBytes, 0, optHeaderBytes.Length);
                            return;
                        }

                        // Read request body for POST
                        string requestBody = "";
                        if (contentLength > 0 && contentLength < 30000000)
                        {
                            char[] bodyBuffer = new char[contentLength];
                            int totalRead = 0;
                            while (totalRead < contentLength)
                            {
                                int read = reader.Read(bodyBuffer, totalRead, contentLength - totalRead);
                                if (read <= 0) break;
                                totalRead += read;
                            }
                            requestBody = new string(bodyBuffer, 0, totalRead);
                        }

                        // Serve routes
                        if (path == "/prev" || path == "/next" || path == "/play")
                        {
                            if (path == "/play")
                            {
                                ControlPowerPointVideo();
                            }
                            else
                            {
                                ControlPowerPoint(path == "/next");
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\"}");
                        }
                        else if (path == "/volume")
                        {
                            string valStr = "";
                            int valIdx = parts[1].IndexOf("val=");
                            if (valIdx >= 0)
                            {
                                valStr = parts[1].Substring(valIdx + 4);
                            }
                            
                            if (!string.IsNullOrEmpty(valStr))
                            {
                                float vol = 0.5f;
                                if (float.TryParse(valStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out vol))
                                {
                                    ControlPowerPointVolume(vol);
                                }
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\"}");
                        }
                        else if (path == "/api/cast_photo")
                        {
                            // Receive high-res snapshot from phone camera
                            if (!string.IsNullOrEmpty(requestBody))
                            {
                                string imgData = ExtractJsonString(requestBody, "image");
                                if (!string.IsNullOrEmpty(imgData))
                                {
                                    CurrentCastImage = imgData;
                                    CastMode = "photo";
                                    CastVersion++;
                                    
                                    if (Globals.ThisAddIn != null)
                                    {
                                        Globals.ThisAddIn.ShowLiveCast(ServerUrl + "cast_view");
                                    }
                                }
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\",\"version\":" + CastVersion + "}");
                        }
                        else if (path == "/api/cast_frame")
                        {
                            // Receive live stream frame from phone camera
                            if (!string.IsNullOrEmpty(requestBody))
                            {
                                string imgData = ExtractJsonString(requestBody, "image");
                                if (!string.IsNullOrEmpty(imgData))
                                {
                                    bool isFirst = (CastMode != "stream");
                                    CurrentCastImage = imgData;
                                    CastMode = "stream";
                                    CastVersion++;
                                    
                                    if (isFirst && Globals.ThisAddIn != null)
                                    {
                                        Globals.ThisAddIn.ShowLiveCast(ServerUrl + "cast_view");
                                    }
                                }
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\",\"version\":" + CastVersion + "}");
                        }
                        else if (path == "/api/start_mirror")
                        {
                            // Connect to external mobile screen stream (e.g. ScreenStream URL)
                            string mirrorUrl = ExtractJsonString(requestBody, "url");
                            if (!string.IsNullOrEmpty(mirrorUrl))
                            {
                                CastStreamUrl = mirrorUrl;
                                CastMode = "mirror";
                                if (Globals.ThisAddIn != null)
                                {
                                    Globals.ThisAddIn.ShowLiveCast(mirrorUrl);
                                }
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\"}");
                        }
                        else if (path == "/api/stop_cast" || path == "/stop_cast")
                        {
                            CastMode = "none";
                            CurrentCastImage = "";
                            if (Globals.ThisAddIn != null)
                            {
                                Globals.ThisAddIn.CloseLiveCast();
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\"}");
                        }
                        else if (path == "/api/rotate_cast")
                        {
                            CastRotation = (CastRotation + 90) % 360;
                            SendJsonResponse(stream, "{\"status\":\"success\",\"rotation\":" + CastRotation + "}");
                        }
                        else if (path == "/api/get_cast_data")
                        {
                            string json = string.Format(
                                "{{\"mode\":\"{0}\",\"version\":{1},\"rotation\":{2},\"url\":\"{3}\",\"image\":\"{4}\"}}",
                                CastMode, CastVersion, CastRotation, EscapeJson(CastStreamUrl), EscapeJson(CurrentCastImage)
                            );
                            SendJsonResponse(stream, json);
                        }
                        else if (path == "/cast_view")
                        {
                            string html = GetCastViewHtml();
                            SendHtmlResponse(stream, html);
                        }
                        else if (path == "/console" || path == "/console.html" || path == "/console/")
                        {
                            string html = GetControlPageHtml();
                            SendHtmlResponse(stream, html);
                        }
                        else if (!string.IsNullOrEmpty(CurrentProxyTarget) && path != "/" && path != "/index.html")
                        {
                            ProxyRequest(stream, method, CurrentProxyTarget, parts[1]);
                        }
                        else if (path == "/" || path == "/index.html")
                        {
                            if (!string.IsNullOrEmpty(CurrentProxyTarget))
                            {
                                ProxyRequest(stream, method, CurrentProxyTarget, parts[1]);
                            }
                            else
                            {
                                string headers = "HTTP/1.1 302 Found\r\n" +
                                                 "Location: /console\r\n" +
                                                 "Content-Length: 0\r\n" +
                                                 "Connection: close\r\n\r\n";
                                byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                                stream.Write(headerBytes, 0, headerBytes.Length);
                            }
                        }
                        else
                        {
                            string headers = "HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\nConnection: close\r\n\r\n";
                            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                            stream.Write(headerBytes, 0, headerBytes.Length);
                        }

                        stream.Flush();
                        try
                        {
                            client.Client.Shutdown(SocketShutdown.Send);
                            Thread.Sleep(50);
                        }
                        catch { }
                    }
                }
                catch { }
            }
        }

        private static void SendJsonResponse(Stream stream, string json)
        {
            byte[] bodyBytes = Encoding.UTF8.GetBytes(json);
            string headers = "HTTP/1.1 200 OK\r\n" +
                             "Content-Type: application/json; charset=utf-8\r\n" +
                             "Access-Control-Allow-Origin: *\r\n" +
                             "Content-Length: " + bodyBytes.Length + "\r\n" +
                             "Connection: close\r\n\r\n";
            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
            stream.Write(headerBytes, 0, headerBytes.Length);
            stream.Write(bodyBytes, 0, bodyBytes.Length);
        }

        private static void SendHtmlResponse(Stream stream, string html)
        {
            byte[] bodyBytes = Encoding.UTF8.GetBytes(html);
            string headers = "HTTP/1.1 200 OK\r\n" +
                             "Content-Type: text/html; charset=utf-8\r\n" +
                             "Access-Control-Allow-Origin: *\r\n" +
                             "Content-Length: " + bodyBytes.Length + "\r\n" +
                             "Connection: close\r\n\r\n";
            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
            stream.Write(headerBytes, 0, headerBytes.Length);
            stream.Write(bodyBytes, 0, bodyBytes.Length);
        }

        private static string ExtractJsonString(string json, string key)
        {
            if (string.IsNullOrEmpty(json)) return "";
            string search = "\"" + key + "\":";
            int idx = json.IndexOf(search);
            if (idx < 0) return "";
            idx += search.Length;
            while (idx < json.Length && (json[idx] == ' ' || json[idx] == '\t' || json[idx] == '\r' || json[idx] == '\n')) idx++;
            if (idx >= json.Length) return "";
            if (json[idx] == '"')
            {
                idx++;
                int end = idx;
                while (end < json.Length)
                {
                    if (json[end] == '"' && json[end - 1] != '\\') break;
                    end++;
                }
                return json.Substring(idx, end - idx).Replace("\\\"", "\"").Replace("\\\\", "\\");
            }
            return "";
        }

        private static string EscapeJson(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n");
        }

        private void ControlPowerPoint(bool next)
        {
            Globals.ThisAddIn.SafeInvoke(delegate()
            {
                try
                {
                    ResetActiveVideoFocus();
                    var app = Globals.ThisAddIn.Application;
                    if (app.SlideShowWindows.Count > 0)
                    {
                        var view = app.SlideShowWindows[1].View;
                        if (view.State == PowerPoint.PpSlideShowState.ppSlideShowPaused)
                        {
                            try { view.State = PowerPoint.PpSlideShowState.ppSlideShowRunning; } catch { }
                        }

                        if (next)
                        {
                            view.Next();
                        }
                        else
                        {
                            view.Previous();
                        }
                    }
                }
                catch { }
            });
        }

        private void ControlPowerPointVolume(float volume)
        {
            Globals.ThisAddIn.SafeInvoke(delegate()
            {
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app.SlideShowWindows.Count > 0)
                    {
                        var slide = app.SlideShowWindows[1].View.Slide;
                        AdjustMediaVolumeRecursive(slide.Shapes, volume);
                    }
                }
                catch { }
            });
        }

        private void AdjustMediaVolumeRecursive(object shapesObj, float volume)
        {
            if (shapesObj == null) return;
            if (shapesObj is PowerPoint.Shapes)
            {
                var shapes = (PowerPoint.Shapes)shapesObj;
                foreach (PowerPoint.Shape shape in shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                    {
                        try { AdjustMediaVolumeRecursive(shape.GroupItems, volume); } catch { }
                    }
                    else
                    {
                        try
                        {
                            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia)
                            {
                                shape.MediaFormat.Volume = volume;
                            }
                        }
                        catch { }
                    }
                }
            }
        }

        private void ControlPowerPointVideo()
        {
            Globals.ThisAddIn.SafeInvoke(delegate()
            {
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app.SlideShowWindows.Count > 0)
                    {
                        var ssw = app.SlideShowWindows[1];
                        var view = ssw.View;
                        var slide = view.Slide;
                        
                        var videoShapes = new List<PowerPoint.Shape>();
                        CollectMediaShapesRecursive(slide.Shapes, videoShapes);
                        
                        if (videoShapes.Count == 0) return;

                        PowerPoint.Shape targetShape = null;
                        if (!string.IsNullOrEmpty(_activeVideoShapeName))
                        {
                            foreach (var s in videoShapes)
                            {
                                if (s.Name == _activeVideoShapeName)
                                {
                                    targetShape = s;
                                    break;
                                }
                            }
                        }
                        if (targetShape == null)
                        {
                            targetShape = videoShapes[0];
                            _activeVideoShapeName = targetShape.Name;
                        }

                        bool isPlaying = false;
                        _videoPlayingStates.TryGetValue(targetShape.Name, out isPlaying);

                        try
                        {
                            IntPtr hwnd = (IntPtr)(uint)ssw.HWND;
                            if (hwnd != IntPtr.Zero)
                            {
                                PostMessage(hwnd, WM_KEYDOWN, (IntPtr)VK_SPACE, IntPtr.Zero);
                                PostMessage(hwnd, WM_KEYUP, (IntPtr)VK_SPACE, IntPtr.Zero);
                                _videoPlayingStates[targetShape.Name] = !isPlaying;
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            });
        }

        private void CollectMediaShapesRecursive(object shapesObj, List<PowerPoint.Shape> result)
        {
            if (shapesObj == null) return;
            if (shapesObj is PowerPoint.Shapes)
            {
                var shapes = (PowerPoint.Shapes)shapesObj;
                foreach (PowerPoint.Shape shape in shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                    {
                        try { CollectMediaShapesRecursive(shape.GroupItems, result); } catch { }
                    }
                    else
                    {
                        try
                        {
                            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia)
                            {
                                result.Add(shape);
                            }
                        }
                        catch { }
                    }
                }
            }
        }

        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP = 0x0101;
        private const int VK_SPACE = 0x20;

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        private void ProxyRequest(Stream clientStream, string method, string targetBaseUrl, string relativePath)
        {
            try
            {
                string targetUrl = targetBaseUrl.TrimEnd('/') + "/" + relativePath.TrimStart('/');
                var req = (HttpWebRequest)WebRequest.Create(targetUrl);
                req.Method = method;
                req.Timeout = 10000;
                req.UserAgent = "YuLink-Proxy/1.0";
                
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var respStream = resp.GetResponseStream())
                {
                    string headers = string.Format("HTTP/1.1 {0} {1}\r\nContent-Type: {2}\r\nAccess-Control-Allow-Origin: *\r\nConnection: close\r\n\r\n",
                        (int)resp.StatusCode, resp.StatusDescription, resp.ContentType);
                    byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                    clientStream.Write(headerBytes, 0, headerBytes.Length);
                    
                    byte[] buffer = new byte[8192];
                    int bytesRead;
                    while ((bytesRead = respStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        clientStream.Write(buffer, 0, bytesRead);
                    }
                }
            }
            catch (Exception ex)
            {
                try
                {
                    byte[] err = Encoding.UTF8.GetBytes("HTTP/1.1 502 Bad Gateway\r\nContent-Type: text/plain\r\n\r\nProxy Error: " + ex.Message);
                    clientStream.Write(err, 0, err.Length);
                }
                catch { }
            }
        }

        [DllImport("wininet.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);

        private void SetWindowsSystemProxy(string proxyServer, bool enabled)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings", true))
                {
                    if (key != null)
                    {
                        key.SetValue("ProxyEnable", enabled ? 1 : 0);
                        if (enabled && !string.IsNullOrEmpty(proxyServer))
                        {
                            key.SetValue("ProxyServer", proxyServer);
                        }
                    }
                }
                InternetSetOption(IntPtr.Zero, 39, IntPtr.Zero, 0);
                InternetSetOption(IntPtr.Zero, 37, IntPtr.Zero, 0);
            }
            catch { }
        }

        // =========================================================================
        // Full Screen Visualizer Screen Receiver HTML (Large Screen in PPT)
        // =========================================================================
        private string GetCastViewHtml()
        {
            return @"<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
    <meta charset=""utf-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>YuLink 无线展台 · 大屏接收端</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; user-select: none; }
        body {
            background-color: #0b0c10;
            color: #ffffff;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            width: 100vw;
            height: 100vh;
            overflow: hidden;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        #viewContainer {
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
            background: radial-gradient(circle at 50% 50%, #1a1c23 0%, #08090c 100%);
        }
        #castImage {
            max-width: 96vw;
            max-height: 94vh;
            object-fit: contain;
            border-radius: 12px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.8);
            transition: transform 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        }
        .hud-header {
            position: absolute;
            top: 16px;
            left: 24px;
            display: flex;
            align-items: center;
            gap: 12px;
            padding: 8px 16px;
            background: rgba(20, 22, 28, 0.75);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1px solid rgba(255, 255, 255, 0.12);
            border-radius: 20px;
            z-index: 100;
            box-shadow: 0 4px 16px rgba(0,0,0,0.3);
        }
        .hud-dot {
            width: 8px;
            height: 8px;
            background: #30d158;
            border-radius: 50%;
            box-shadow: 0 0 8px #30d158;
            animation: pulse 2s infinite;
        }
        @keyframes pulse { 0% { opacity: 0.5; } 50% { opacity: 1; } 100% { opacity: 0.5; } }
        .hud-title { font-size: 13px; font-weight: 600; letter-spacing: 0.5px; color: #f0f2f5; }
        
        .hud-tools {
            position: absolute;
            top: 16px;
            right: 24px;
            display: flex;
            gap: 8px;
            z-index: 100;
        }
        .hud-btn {
            background: rgba(255, 255, 255, 0.1);
            backdrop-filter: blur(20px);
            border: 1px solid rgba(255, 255, 255, 0.15);
            color: #fff;
            padding: 8px 14px;
            border-radius: 18px;
            font-size: 12px;
            font-weight: 500;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s ease;
        }
        .hud-btn:hover { background: rgba(255, 255, 255, 0.2); transform: translateY(-1px); }
        .hud-btn.close:hover { background: rgba(255, 69, 58, 0.8); border-color: rgba(255, 69, 58, 0.8); }
        
        .empty-hint {
            text-align: center;
            color: #8e8e93;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 12px;
        }
        .empty-hint svg { width: 64px; height: 64px; fill: #636366; }
    </style>
</head>
<body>
    <div id=""viewContainer"">
        <div class=""hud-header"">
            <div class=""hud-dot""></div>
            <div class=""hud-title"" id=""castStatus"">YuLink 无线展台 · 等待手机投屏</div>
        </div>

        <div class=""hud-tools"">
            <button class=""hud-btn"" onclick=""rotateImage()"">🔄 旋转 90°</button>
            <button class=""hud-btn"" onclick=""zoomIn()"">🔍 放大</button>
            <button class=""hud-btn"" onclick=""zoomOut()"">🔎 缩小</button>
            <button class=""hud-btn close"" onclick=""stopCast()"">❌ 退出展台</button>
        </div>

        <div id=""emptyState"" class=""empty-hint"">
            <svg viewBox=""0 0 24 24""><path d=""M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm-1-13h2v6h-2zm0 8h2v2h-2z""/></svg>
            <div style=""font-size: 16px; font-weight: 500;"">请使用手机在遥控器中点击【拍照投屏】或【开启实时展台】</div>
            <div style=""font-size: 13px; color: #636366;"">投屏后可直接使用画笔在投影屏幕上圈画批改</div>
        </div>

        <img id=""castImage"" style=""display: none;"" alt=""Cast Preview"" />
    </div>

    <script>
        let currentVersion = -1;
        let currentRotation = 0;
        let currentScale = 1.0;
        const img = document.getElementById('castImage');
        const emptyState = document.getElementById('emptyState');
        const statusEl = document.getElementById('castStatus');

        function updateTransform() {
            img.style.transform = 'rotate(' + currentRotation + 'deg) scale(' + currentScale + ')';
        }

        function rotateImage() {
            currentRotation = (currentRotation + 90) % 360;
            updateTransform();
        }

        function zoomIn() {
            currentScale = Math.min(3.0, currentScale + 0.2);
            updateTransform();
        }

        function zoomOut() {
            currentScale = Math.max(0.5, currentScale - 0.2);
            updateTransform();
        }

        function stopCast() {
            fetch('/api/stop_cast', { method: 'POST' });
        }

        async function pollCastData() {
            try {
                const res = await fetch('/api/get_cast_data');
                if (res.ok) {
                    const data = await res.json();
                    if (data.mode === 'none') {
                        emptyState.style.display = 'flex';
                        img.style.display = 'none';
                        statusEl.innerText = 'YuLink 无线展台 · 等待手机投屏';
                    } else if (data.image && data.version !== currentVersion) {
                        currentVersion = data.version;
                        img.src = data.image;
                        img.style.display = 'block';
                        emptyState.style.display = 'none';
                        if (data.mode === 'photo') {
                            statusEl.innerText = '📸 实物作业拍照讲评 (可使用画笔圈画批注)';
                        } else if (data.mode === 'stream') {
                            statusEl.innerText = '🎥 实时动态实物展台直播中';
                        }
                    }
                }
            } catch (e) {}
            setTimeout(pollCastData, 100);
        }

        pollCastData();
    </script>
</body>
</html>";
        }

        // =========================================================================
        // Mobile All-In-One Remote Controller & Camera Visualizer HTML
        // =========================================================================
        private string GetControlPageHtml()
        {
            return @"<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
    <meta charset=""utf-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover"">
    <title>YuLink 移动智能中枢</title>
    <style>
        :root {
            --bg-color: #0c0d12;
            --card-bg: rgba(28, 30, 38, 0.7);
            --card-border: rgba(255, 255, 255, 0.1);
            --accent-green: #30d158;
            --accent-blue: #0a84ff;
            --accent-red: #ff453a;
            --text-main: #f5f5f7;
            --text-sub: #86868b;
        }

        * {
            box-sizing: border-box;
            -webkit-tap-highlight-color: transparent;
            margin: 0;
            padding: 0;
            user-select: none;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, ""SF Pro Text"", ""Segoe UI"", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-main);
            min-height: 100dvh;
            display: flex;
            flex-direction: column;
            align-items: center;
            overflow-x: hidden;
            padding-bottom: 30px;
            background-image: radial-gradient(circle at 50% -10%, #1c2235, transparent 60%);
        }

        /* Top Segmented Tab Switcher */
        .tab-bar-container {
            width: 90%;
            max-width: 360px;
            margin-top: 20px;
            margin-bottom: 20px;
            background: rgba(255, 255, 255, 0.08);
            backdrop-filter: blur(25px);
            -webkit-backdrop-filter: blur(25px);
            border-radius: 24px;
            padding: 4px;
            display: flex;
            border: 1px solid var(--card-border);
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
        }

        .tab-btn {
            flex: 1;
            padding: 10px 0;
            text-align: center;
            font-size: 13px;
            font-weight: 600;
            border-radius: 20px;
            color: var(--text-sub);
            cursor: pointer;
            transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
            border: none;
            background: transparent;
        }

        .tab-btn.active {
            background: #ffffff;
            color: #000000;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.2);
        }

        /* Tab Content Panes */
        .tab-pane {
            display: none;
            width: 90%;
            max-width: 360px;
            flex-direction: column;
            align-items: center;
            animation: fadeIn 0.25s ease forwards;
        }

        .tab-pane.active {
            display: flex;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(6px); }
            to { opacity: 1; transform: translateY(0); }
        }

        /* 🎮 Tab 1: Remote Controller */
        .remote-body {
            width: 100%;
            max-width: 290px;
            padding: 36px 24px 30px 24px;
            background: linear-gradient(145deg, #22242c, #16181f);
            border-radius: 40px;
            border: 1px solid rgba(255, 255, 255, 0.12);
            box-shadow: 0 20px 40px rgba(0, 0, 0, 0.6), inset 0 1px 1px rgba(255, 255, 255, 0.2);
            display: flex;
            flex-direction: column;
            align-items: center;
            position: relative;
        }

        .led-indicator {
            width: 14px;
            height: 14px;
            background-color: #111;
            border-radius: 50%;
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: inset 0 1px 3px rgba(0,0,0,0.8);
        }

        .led-light {
            width: 6px;
            height: 6px;
            background-color: var(--accent-green);
            border-radius: 50%;
            box-shadow: 0 0 8px var(--accent-green);
            transition: all 0.15s ease;
        }

        .remote-circle-pad {
            width: 210px;
            height: 210px;
            border-radius: 50%;
            background: linear-gradient(145deg, #1b1d24, #13141a);
            position: relative;
            box-shadow: 0 8px 24px rgba(0,0,0,0.5), inset 0 1px 1px rgba(255,255,255,0.1);
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 30px;
        }

        .pad-btn {
            position: absolute;
            background: none;
            border: none;
            color: #f0f0f5;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            z-index: 2;
        }

        .pad-btn:active svg {
            transform: scale(0.9);
            fill: var(--accent-green);
        }

        .pad-btn.top { top: 12px; left: 0; width: 100%; height: 60px; }
        .pad-btn.bottom { bottom: 12px; left: 0; width: 100%; height: 60px; }
        .pad-btn.center {
            width: 82px;
            height: 82px;
            border-radius: 50%;
            background: linear-gradient(145deg, #2a2c36, #1c1e26);
            box-shadow: 0 4px 12px rgba(0,0,0,0.4), inset 0 1px 1px rgba(255,255,255,0.2);
            z-index: 3;
        }

        .pad-btn svg {
            width: 26px;
            height: 26px;
            fill: #e5e5ea;
            transition: all 0.15s ease;
        }

        .bottom-actions {
            width: 100%;
            display: flex;
            justify-content: space-around;
        }

        .pill-btn {
            width: 72px;
            height: 38px;
            border-radius: 19px;
            background: linear-gradient(145deg, #2a2c36, #1c1e26);
            border: 1px solid rgba(255,255,255,0.08);
            box-shadow: 0 4px 12px rgba(0,0,0,0.3);
            display: flex;
            align-items: center;
            justify-content: center;
            color: #8e8e93;
            cursor: pointer;
        }

        .pill-btn:active {
            transform: scale(0.95);
            background: #111;
        }

        /* 📷 Tab 2: Visualizer Camera */
        .camera-card {
            width: 100%;
            background: var(--card-bg);
            border-radius: 28px;
            border: 1px solid var(--card-border);
            padding: 16px;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 16px;
            box-shadow: 0 12px 32px rgba(0,0,0,0.4);
        }

        .preview-wrapper {
            width: 100%;
            height: 240px;
            background: #000;
            border-radius: 20px;
            overflow: hidden;
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
            border: 1px solid rgba(255,255,255,0.1);
        }

        #cameraVideo {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }

        .camera-actions {
            width: 100%;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
        }

        .action-btn {
            padding: 14px 10px;
            border-radius: 18px;
            border: none;
            font-size: 14px;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            transition: all 0.2s ease;
        }

        .action-btn:active { transform: scale(0.96); }
        .action-btn.primary {
            background: linear-gradient(135deg, #30d158, #248a3d);
            color: #ffffff;
            grid-column: span 2;
            padding: 16px;
            font-size: 16px;
            box-shadow: 0 4px 16px rgba(48, 209, 88, 0.4);
        }
        .action-btn.secondary {
            background: rgba(255, 255, 255, 0.1);
            color: #f5f5f7;
            border: 1px solid rgba(255, 255, 255, 0.15);
        }
        .action-btn.danger {
            background: rgba(255, 69, 58, 0.15);
            color: #ff453a;
            border: 1px solid rgba(255, 69, 58, 0.3);
            grid-column: span 2;
        }

        /* 📱 Tab 3: Screen Mirror */
        .mirror-card {
            width: 100%;
            background: var(--card-bg);
            border-radius: 28px;
            border: 1px solid var(--card-border);
            padding: 20px;
            display: flex;
            flex-direction: column;
            gap: 16px;
        }

        .input-box {
            width: 100%;
            padding: 12px 16px;
            background: rgba(0, 0, 0, 0.4);
            border: 1px solid rgba(255, 255, 255, 0.15);
            border-radius: 14px;
            color: #fff;
            font-size: 14px;
            outline: none;
        }

        .input-box:focus { border-color: var(--accent-blue); }

        .tip-text {
            font-size: 12px;
            color: var(--text-sub);
            line-height: 1.5;
        }
    </style>
</head>
<body>

    <!-- Top Tab Bar -->
    <div class=""tab-bar-container"">
        <button class=""tab-btn active"" onclick=""switchTab('remote')"">🎮 遥控</button>
        <button class=""tab-btn"" onclick=""switchTab('camera')"">📷 实物展台</button>
        <button class=""tab-btn"" onclick=""switchTab('mirror')"">📱 投屏</button>
    </div>

    <!-- 🎮 Tab 1: Remote Controller Pane -->
    <div id=""pane-remote"" class=""tab-pane active"">
        <div class=""remote-body"">
            <div class=""led-indicator"">
                <div id=""remoteLed"" class=""led-light""></div>
            </div>

            <div class=""remote-circle-pad"">
                <button class=""pad-btn top"" onclick=""sendCmd('/prev')"">
                    <svg viewBox=""0 0 24 24""><path d=""M7.41 15.41L12 10.83l4.59 4.58L18 14l-6-6-6 6z""/></svg>
                </button>
                <button class=""pad-btn center"" onclick=""sendCmd('/next')"">
                    <svg viewBox=""0 0 24 24""><circle cx=""12"" cy=""12"" r=""5""/></svg>
                </button>
                <button class=""pad-btn bottom"" onclick=""sendCmd('/next')"">
                    <svg viewBox=""0 0 24 24""><path d=""M7.41 8.59L12 13.17l4.59-4.58L18 10l-6 6-6-6z""/></svg>
                </button>
            </div>

            <div class=""bottom-actions"">
                <button class=""pill-btn"" onclick=""sendCmd('/play')"">
                    <svg style=""width:18px;height:18px;fill:#8e8e93;"" viewBox=""0 0 24 24""><path d=""M8 5v14l11-7z""/></svg>
                </button>
                <button class=""pill-btn"" onclick=""sendCmd('/next')"">
                    <svg style=""width:18px;height:18px;fill:#8e8e93;"" viewBox=""0 0 24 24""><path d=""M6 18l8.5-6L6 6v12zM16 6v12h2V6h-2z""/></svg>
                </button>
            </div>
        </div>
    </div>

    <!-- 📷 Tab 2: Visualizer Camera Pane -->
    <div id=""pane-camera"" class=""tab-pane"">
        <div class=""camera-card"">
            <div class=""preview-wrapper"">
                <video id=""cameraVideo"" autoplay playsinline muted></video>
                <canvas id=""captureCanvas"" style=""display:none;""></canvas>
            </div>

            <div class=""camera-actions"">
                <button class=""action-btn primary"" onclick=""snapAndCast()"">
                    📸 拍照并投屏讲评
                </button>
                <button id=""btnStream"" class=""action-btn secondary"" onclick=""toggleLiveStream()"">
                    🎥 开启实时展台
                </button>
                <button class=""action-btn secondary"" onclick=""switchCamera()"">
                    🔄 翻转镜头
                </button>
                <button class=""action-btn secondary"" onclick=""rotateScreen()"">
                    🔃 大屏旋转90°
                </button>
                <button class=""action-btn secondary"" onclick=""toggleTorch()"">
                    💡 补光灯
                </button>
                <button class=""action-btn danger"" onclick=""stopVisualizer()"">
                    ⏹️ 退出展台 · 恢复幻灯片
                </button>
            </div>
        </div>
    </div>

    <!-- 📱 Tab 3: Screen Mirror Pane -->
    <div id=""pane-mirror"" class=""tab-pane"">
        <div class=""mirror-card"">
            <div style=""font-size: 15px; font-weight: 600;"">📱 手机屏幕镜像</div>
            
            <div class=""tip-text"">
                <b>安卓一键投屏推荐：</b><br>
                1. 手机打开开源无广告 <b>ScreenStream</b> App 并开启推流；<br>
                2. 输入下方地址并点击连接，PPT 将同屏镜像手机全屏。
            </div>

            <input id=""mirrorUrlInput"" class=""input-box"" type=""text"" placeholder=""http://192.168.x.x:8080"" />

            <button class=""action-btn primary"" style=""grid-column: span 1;"" onclick=""startMirror()"">
                🚀 连接并投屏到 PPT
            </button>

            <button class=""action-btn secondary"" onclick=""tryWebShare()"">
                🌐 尝试网页直接投屏
            </button>

            <button class=""action-btn danger"" onclick=""stopVisualizer()"">
                ⏹️ 退出投屏 · 恢复幻灯片
            </button>
        </div>
    </div>

    <script>
        // Haptic feedback helper
        function vibrate() {
            if (navigator.vibrate) navigator.vibrate(25);
        }

        // Tab Switcher
        function switchTab(name) {
            vibrate();
            document.querySelectorAll('.tab-btn').forEach((b, i) => {
                b.classList.remove('active');
            });
            document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));

            if (name === 'remote') {
                document.querySelectorAll('.tab-btn')[0].classList.add('active');
                document.getElementById('pane-remote').classList.add('active');
            } else if (name === 'camera') {
                document.querySelectorAll('.tab-btn')[1].classList.add('active');
                document.getElementById('pane-camera').classList.add('active');
                startCamera();
            } else if (name === 'mirror') {
                document.querySelectorAll('.tab-btn')[2].classList.add('active');
                document.getElementById('pane-mirror').classList.add('active');
            }
        }

        // Remote command sender
        async function sendCmd(path) {
            vibrate();
            const led = document.getElementById('remoteLed');
            if (led) {
                led.style.backgroundColor = '#ff9500';
                led.style.boxShadow = '0 0 10px #ff9500';
            }
            try {
                await fetch(path);
            } catch (e) {}
            setTimeout(() => {
                if (led) {
                    led.style.backgroundColor = 'var(--accent-green)';
                    led.style.boxShadow = '0 0 8px var(--accent-green)';
                }
            }, 200);
        }

        // Camera Visualizer Logic
        let mediaStream = null;
        let currentFacingMode = 'environment';
        let isStreaming = false;
        let streamTimer = null;
        const video = document.getElementById('cameraVideo');
        const canvas = document.getElementById('captureCanvas');

        async function startCamera() {
            if (mediaStream) {
                mediaStream.getTracks().forEach(t => t.stop());
            }
            try {
                mediaStream = await navigator.mediaDevices.getUserMedia({
                    video: {
                        facingMode: { ideal: currentFacingMode },
                        width: { ideal: 1920 },
                        height: { ideal: 1080 }
                    },
                    audio: false
                });
                video.srcObject = mediaStream;
            } catch (e) {
                alert('无法调用手机摄像头，请确保在浏览器设置中授予了相机权限。\n错误: ' + e.message);
            }
        }

        function switchCamera() {
            vibrate();
            currentFacingMode = (currentFacingMode === 'environment' ? 'user' : 'environment');
            startCamera();
        }

        let torchOn = false;
        async function toggleTorch() {
            vibrate();
            if (mediaStream) {
                const track = mediaStream.getVideoTracks()[0];
                if (track && track.getCapabilities && track.getCapabilities().torch) {
                    torchOn = !torchOn;
                    await track.applyConstraints({ advanced: [{ torch: torchOn }] });
                } else {
                    alert('当前摄像头不支持或未开启手电筒功能。');
                }
            }
        }

        async function snapAndCast() {
            vibrate();
            if (!video.videoWidth) return;
            canvas.width = video.videoWidth;
            canvas.height = video.videoHeight;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
            const base64 = canvas.toDataURL('image/jpeg', 0.88);

            try {
                const res = await fetch('/api/cast_photo', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ image: base64 })
                });
                if (res.ok) {
                    alert('📸 照片已秒级投射到 PPT 大屏！可在屏幕上使用画笔直接圈画批改。');
                }
            } catch (e) {
                alert('上传照片失败: ' + e.message);
            }
        }

        function toggleLiveStream() {
            vibrate();
            const btn = document.getElementById('btnStream');
            isStreaming = !isStreaming;
            if (isStreaming) {
                btn.innerText = '⏹️ 停止实时直播';
                btn.style.background = '#ff453a';
                sendLiveFrame();
            } else {
                btn.innerText = '🎥 开启实时展台';
                btn.style.background = 'rgba(255, 255, 255, 0.1)';
                if (streamTimer) clearTimeout(streamTimer);
            }
        }

        async function sendLiveFrame() {
            if (!isStreaming) return;
            if (video.videoWidth) {
                canvas.width = Math.min(1280, video.videoWidth);
                canvas.height = Math.round(canvas.width * (video.videoHeight / video.videoWidth));
                const ctx = canvas.getContext('2d');
                ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
                const base64 = canvas.toDataURL('image/jpeg', 0.65);

                try {
                    await fetch('/api/cast_frame', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ image: base64 })
                    });
                } catch (e) {}
            }
            if (isStreaming) {
                streamTimer = setTimeout(sendLiveFrame, 120);
            }
        }

        async function rotateScreen() {
            vibrate();
            try {
                await fetch('/api/rotate_cast', { method: 'POST' });
            } catch (e) {}
        }

        async function stopVisualizer() {
            vibrate();
            if (isStreaming) toggleLiveStream();
            try {
                await fetch('/api/stop_cast', { method: 'POST' });
            } catch (e) {}
        }

        // Screen Mirror
        async function startMirror() {
            vibrate();
            let url = document.getElementById('mirrorUrlInput').value.trim();
            if (!url) {
                alert('请输入手机推流地址 (如 http://192.168.1.x:8080)');
                return;
            }
            if (!url.startsWith('http://') && !url.startsWith('https://')) {
                url = 'http://' + url;
            }
            try {
                await fetch('/api/start_mirror', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ url: url })
                });
                alert('🚀 正在连接投屏并在 PPT 页面呈现！');
            } catch (e) {
                alert('连接投屏失败: ' + e.message);
            }
        }

        async function tryWebShare() {
            if (navigator.mediaDevices && navigator.mediaDevices.getDisplayMedia) {
                try {
                    const stream = await navigator.mediaDevices.getDisplayMedia({ video: true });
                    video.srcObject = stream;
                    switchTab('camera');
                    toggleLiveStream();
                } catch (e) {
                    alert('屏幕共享已取消或不受支持: ' + e.message);
                }
            } else {
                alert('当前手机浏览器不支持原生 getDisplayMedia，请使用 ScreenStream 或拍照展台模式。');
            }
        }
    </script>
</body>
</html>";
        }
    }
}
