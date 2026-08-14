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
        public static string CastMode = "none"; // none, photo, stream
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

        public static string GetBestLocalIpAddress()
        {
            try
            {
                var candidates = new List<KeyValuePair<int, string>>();
                var nics = NetworkInterface.GetAllNetworkInterfaces();

                foreach (var nic in nics)
                {
                    if (nic.OperationalStatus != OperationalStatus.Up || 
                        nic.NetworkInterfaceType == NetworkInterfaceType.Loopback)
                    {
                        continue;
                    }

                    string name = (nic.Name + " " + nic.Description).ToLower();
                    // Exclude virtual / proxy / container / VPN adapters
                    if (name.Contains("meta") || name.Contains("clash") || name.Contains("wsl") || 
                        name.Contains("vethernet") || name.Contains("vmware") || name.Contains("virtualbox") || 
                        name.Contains("tap") || name.Contains("docker") || name.Contains("tailscale") || 
                        name.Contains("zerotier") || name.Contains("sing-box") || name.Contains("npcap"))
                    {
                        continue;
                    }

                    var ipProps = nic.GetIPProperties();
                    foreach (var ip in ipProps.UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == AddressFamily.InterNetwork && 
                            !IPAddress.IsLoopback(ip.Address))
                        {
                            string ipStr = ip.Address.ToString();
                            if (ipStr.StartsWith("198.18.") || ipStr.StartsWith("169.254.")) continue;

                            int score = 0;
                            if (nic.NetworkInterfaceType == NetworkInterfaceType.Wireless80211 || 
                                name.Contains("wlan") || name.Contains("wi-fi") || name.Contains("wireless"))
                            {
                                score += 100; // Prioritize Wi-Fi
                            }
                            else if (nic.NetworkInterfaceType == NetworkInterfaceType.Ethernet || 
                                     name.Contains("ethernet") || name.Contains("以太网"))
                            {
                                score += 50;
                            }

                            if (ipProps.GatewayAddresses.Count > 0) score += 20;

                            if (ipStr.StartsWith("192.168.")) score += 30;
                            else if (ipStr.StartsWith("10.")) score += 20;
                            else if (ipStr.StartsWith("172.")) score += 10;

                            candidates.Add(new KeyValuePair<int, string>(score, ipStr));
                        }
                    }
                }

                if (candidates.Count > 0)
                {
                    candidates.Sort((a, b) => b.Key.CompareTo(a.Key));
                    return candidates[0].Value;
                }
            }
            catch { }
            return "127.0.0.1";
        }

        public static List<KeyValuePair<string, string>> GetAllAvailableIps()
        {
            var list = new List<KeyValuePair<string, string>>();
            try
            {
                var nics = NetworkInterface.GetAllNetworkInterfaces();
                foreach (var nic in nics)
                {
                    if (nic.OperationalStatus != OperationalStatus.Up || 
                        nic.NetworkInterfaceType == NetworkInterfaceType.Loopback) continue;

                    var ipProps = nic.GetIPProperties();
                    foreach (var ip in ipProps.UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == AddressFamily.InterNetwork && 
                            !IPAddress.IsLoopback(ip.Address))
                        {
                            string ipStr = ip.Address.ToString();
                            if (ipStr.StartsWith("169.254.")) continue;
                            list.Add(new KeyValuePair<string, string>(ipStr, nic.Name + " (" + ipStr + ")"));
                        }
                    }
                }
            }
            catch { }
            if (list.Count == 0) list.Add(new KeyValuePair<string, string>("127.0.0.1", "Localhost (127.0.0.1)"));
            return list;
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
                // Auto detect best active LAN IPv4
                string localIp = GetBestLocalIpAddress();
                
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
                                "{{\"mode\":\"{0}\",\"version\":{1},\"rotation\":{2},\"image\":\"{3}\"}}",
                                CastMode, CastVersion, CastRotation, EscapeJson(CurrentCastImage)
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
                                string html = GetControlPageHtml();
                                SendHtmlResponse(stream, html);
                            }
                        }
                        else
                        {
                            string headers = "HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\nConnection: close\r\n\r\n";
                            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                            stream.Write(headerBytes, 0, headerBytes.Length);
                        }

                        stream.Flush();
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
            font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', 'Helvetica Neue', Arial, sans-serif;
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
            background: radial-gradient(circle at 50% 50%, #1c1d22 0%, #0c0d10 100%);
        }
        #castImage {
            max-width: 96vw;
            max-height: 94vh;
            object-fit: contain;
            border-radius: 16px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.7);
            transition: transform 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        }
        .hud-header {
            position: absolute;
            top: 20px;
            left: 24px;
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 8px 18px;
            background: rgba(255, 255, 255, 0.9);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border-radius: 20px;
            z-index: 100;
            box-shadow: 0 4px 20px rgba(0,0,0,0.15);
        }
        .hud-dot {
            width: 8px;
            height: 8px;
            background: #34c759;
            border-radius: 50%;
            box-shadow: 0 0 8px #34c759;
        }
        .hud-title { font-size: 13px; font-weight: 600; color: #1d1d1f; letter-spacing: -0.2px; }
        
        .hud-tools {
            position: absolute;
            top: 20px;
            right: 24px;
            display: flex;
            gap: 10px;
            z-index: 100;
        }
        .hud-btn {
            background: rgba(255, 255, 255, 0.92);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1px solid rgba(0, 0, 0, 0.08);
            color: #1d1d1f;
            padding: 8px 16px;
            border-radius: 18px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .hud-btn:hover { background: #ffffff; transform: scale(1.03); }
        .hud-btn.close { color: #ff3b30; }
        .hud-btn.close:hover { background: #ff3b30; color: #ffffff; }
        
        .empty-hint {
            text-align: center;
            color: #86868b;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 12px;
        }
        .empty-hint svg { width: 56px; height: 56px; fill: #86868b; }
    </style>
</head>
<body>
    <div id=""viewContainer"">
        <div class=""hud-header"">
            <div class=""hud-dot""></div>
            <div class=""hud-title"" id=""castStatus"">YuLink 无线展台 · 等待手机拍照投屏</div>
        </div>

        <div class=""hud-tools"">
            <button class=""hud-btn"" onclick=""rotateImage()"">🔄 旋转 90°</button>
            <button class=""hud-btn"" onclick=""zoomIn()"">🔍 放大</button>
            <button class=""hud-btn"" onclick=""zoomOut()"">🔎 缩小</button>
            <button class=""hud-btn close"" onclick=""stopCast()"">❌ 退出展台</button>
        </div>

        <div id=""emptyState"" class=""empty-hint"">
            <svg viewBox=""0 0 24 24""><path d=""M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm-1-13h2v6h-2zm0 8h2v2h-2z""/></svg>
            <div style=""font-size: 16px; font-weight: 600; color: #f5f5f7;"">请使用手机在【无线展台】中点击拍照投屏</div>
            <div style=""font-size: 13px; color: #86868b;"">投屏后可直接使用画笔在投影屏幕上圈画批改</div>
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
                        statusEl.innerText = 'YuLink 无线展台 · 等待手机拍照投屏';
                    } else if (data.image && data.version !== currentVersion) {
                        currentVersion = data.version;
                        img.src = data.image;
                        img.style.display = 'block';
                        emptyState.style.display = 'none';
                        if (data.mode === 'photo') {
                            statusEl.innerText = '📸 实物作业拍照讲评 (可使用画笔批改)';
                        } else if (data.mode === 'stream') {
                            statusEl.innerText = '🎥 实时动态实物展台直播中';
                        }
                    }
                }
            } catch (e) {}
            setTimeout(pollCastData, 120);
        }

        pollCastData();
    </script>
</body>
</html>";
        }

        // =========================================================================
        // Apple Modern Pure White Minimalist Mobile Web Controller HTML
        // =========================================================================
        private string GetControlPageHtml()
        {
            return @"<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
    <meta charset=""utf-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover"">
    <title>YuLink 遥控中心</title>
    <style>
        :root {
            --bg-page: #f5f5f7;
            --card-bg: #ffffff;
            --card-border: rgba(0, 0, 0, 0.06);
            --text-primary: #1d1d1f;
            --text-secondary: #86868b;
            --btn-neutral-bg: #f2f2f7;
            --btn-neutral-active: #e5e5ea;
            --apple-blue: #0071e3;
            --apple-blue-active: #0077ed;
            --apple-green: #34c759;
            --apple-red: #ff3b30;
            --apple-red-bg: #fff2f1;
        }

        * {
            box-sizing: border-box;
            -webkit-tap-highlight-color: transparent;
            margin: 0;
            padding: 0;
            user-select: none;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, ""SF Pro Text"", ""SF Pro Display"", ""Segoe UI"", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-page);
            color: var(--text-primary);
            min-height: 100dvh;
            display: flex;
            flex-direction: column;
            align-items: center;
            overflow-x: hidden;
            padding-bottom: 30px;
            -webkit-font-smoothing: antialiased;
        }

        /* Top Minimal Header */
        .app-header {
            width: 100%;
            max-width: 360px;
            padding: 24px 20px 14px 20px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .brand-title {
            font-size: 20px;
            font-weight: 700;
            letter-spacing: -0.4px;
            color: var(--text-primary);
        }

        .status-badge {
            display: flex;
            align-items: center;
            gap: 6px;
            background: #ffffff;
            padding: 6px 12px;
            border-radius: 14px;
            font-size: 12px;
            font-weight: 500;
            color: var(--text-secondary);
            border: 1px solid var(--card-border);
            box-shadow: 0 1px 4px rgba(0, 0, 0, 0.04);
        }

        .status-dot {
            width: 6px;
            height: 6px;
            background: var(--apple-green);
            border-radius: 50%;
            box-shadow: 0 0 6px var(--apple-green);
        }

        /* Apple iOS Segmented Control Tab Switcher */
        .tab-bar-container {
            width: 90%;
            max-width: 340px;
            margin-bottom: 24px;
            background: #e3e3e8;
            border-radius: 16px;
            padding: 3px;
            display: flex;
            position: relative;
        }

        .tab-btn {
            flex: 1;
            padding: 10px 0;
            text-align: center;
            font-size: 14px;
            font-weight: 600;
            border-radius: 13px;
            color: var(--text-secondary);
            cursor: pointer;
            transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
            border: none;
            background: transparent;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
        }

        .tab-btn.active {
            background: #ffffff;
            color: var(--text-primary);
            box-shadow: 0 3px 10px rgba(0, 0, 0, 0.12), 0 1px 2px rgba(0, 0, 0, 0.06);
        }

        /* Tab Content Panes */
        .tab-pane {
            display: none;
            width: 90%;
            max-width: 340px;
            flex-direction: column;
            align-items: center;
            animation: paneFadeIn 0.25s cubic-bezier(0.16, 1, 0.3, 1) forwards;
        }

        .tab-pane.active {
            display: flex;
        }

        @keyframes paneFadeIn {
            from { opacity: 0; transform: scale(0.98); }
            to { opacity: 1; transform: scale(1); }
        }

        /* 🎮 Tab 1: Clean Apple Remote Card */
        .remote-card {
            width: 100%;
            background: var(--card-bg);
            border-radius: 32px;
            border: 1px solid var(--card-border);
            padding: 28px 24px 28px 24px;
            display: flex;
            flex-direction: column;
            align-items: center;
            box-shadow: 0 8px 30px rgba(0, 0, 0, 0.04), 0 1px 3px rgba(0, 0, 0, 0.02);
            gap: 20px;
        }

        /* Primary Action Touch Control (Large Apple Next Button) */
        .next-hero-btn {
            width: 100%;
            height: 120px;
            background: var(--apple-blue);
            border-radius: 24px;
            border: none;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            gap: 8px;
            color: #ffffff;
            cursor: pointer;
            transition: all 0.15s ease;
            box-shadow: 0 8px 24px rgba(0, 113, 227, 0.28);
        }

        .next-hero-btn:active {
            transform: scale(0.97);
            background: var(--apple-blue-active);
            box-shadow: 0 4px 12px rgba(0, 113, 227, 0.2);
        }

        .next-hero-btn svg {
            width: 36px;
            height: 36px;
            fill: #ffffff;
        }

        .next-hero-btn span {
            font-size: 17px;
            font-weight: 600;
            letter-spacing: -0.3px;
        }

        /* Navigation Grid */
        .nav-button-grid {
            width: 100%;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
        }

        .flat-btn {
            height: 64px;
            background: var(--btn-neutral-bg);
            border-radius: 20px;
            border: none;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            font-size: 15px;
            font-weight: 600;
            color: var(--text-primary);
            cursor: pointer;
            transition: all 0.15s ease;
        }

        .flat-btn:active {
            transform: scale(0.96);
            background: var(--btn-neutral-active);
        }

        .flat-btn svg {
            width: 20px;
            height: 20px;
            fill: var(--text-primary);
        }

        .tip-status-text {
            font-size: 12px;
            font-weight: 500;
            color: var(--text-secondary);
            text-align: center;
            margin-top: 4px;
        }

        /* 📷 Tab 2: Visualizer Camera Card */
        .visualizer-card {
            width: 100%;
            background: var(--card-bg);
            border-radius: 32px;
            border: 1px solid var(--card-border);
            padding: 20px;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 16px;
            box-shadow: 0 8px 30px rgba(0, 0, 0, 0.04), 0 1px 3px rgba(0, 0, 0, 0.02);
        }

        .camera-viewport {
            width: 100%;
            height: 240px;
            background: #111216;
            border-radius: 22px;
            overflow: hidden;
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
        }

        #cameraVideo {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }

        .camera-action-grid {
            width: 100%;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
        }

        .snap-hero-btn {
            grid-column: span 2;
            height: 60px;
            background: var(--apple-green);
            border-radius: 20px;
            border: none;
            color: #ffffff;
            font-size: 16px;
            font-weight: 600;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            cursor: pointer;
            transition: all 0.15s ease;
            box-shadow: 0 6px 20px rgba(52, 199, 89, 0.3);
        }

        .snap-hero-btn:active {
            transform: scale(0.97);
            background: #2ebd52;
        }

        .snap-hero-btn svg {
            width: 22px;
            height: 22px;
            fill: #ffffff;
        }

        .exit-btn {
            grid-column: span 2;
            height: 52px;
            background: var(--apple-red-bg);
            border: 1px solid rgba(255, 59, 48, 0.15);
            border-radius: 18px;
            color: var(--apple-red);
            font-size: 14px;
            font-weight: 600;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            cursor: pointer;
            transition: all 0.15s ease;
        }

        .exit-btn:active {
            transform: scale(0.97);
            background: rgba(255, 59, 48, 0.15);
        }
    </style>
</head>
<body>

    <!-- Top Minimalist Header -->
    <div class=""app-header"">
        <div class=""brand-title"">YuLink</div>
        <div class=""status-badge"">
            <div id=""ledDot"" class=""status-dot""></div>
            <span>已连接 PPT</span>
        </div>
    </div>

    <!-- Apple Segmented Control Tab Switcher -->
    <div class=""tab-bar-container"">
        <button class=""tab-btn active"" onclick=""switchTab('remote')"">🎮 遥控翻页</button>
        <button class=""tab-btn"" onclick=""switchTab('camera')"">📷 实物展台</button>
    </div>

    <!-- 🎮 Tab 1: Minimalist Remote Controller -->
    <div id=""pane-remote"" class=""tab-pane active"">
        <div class=""remote-card"">
            
            <!-- Large Hero Next Button -->
            <button class=""next-hero-btn"" onclick=""sendCmd('/next')"">
                <svg viewBox=""0 0 24 24""><path d=""M6 18l8.5-6L6 6v12zM16 6v12h2V6h-2z""/></svg>
                <span>下一页 (Next)</span>
            </button>

            <!-- Secondary Navigation Controls -->
            <div class=""nav-button-grid"">
                <button class=""flat-btn"" onclick=""sendCmd('/prev')"">
                    <svg viewBox=""0 0 24 24""><path d=""M6 6h2v12H6zm3.5 6l8.5 6V6z""/></svg>
                    <span>上一页</span>
                </button>
                <button class=""flat-btn"" onclick=""sendCmd('/play')"">
                    <svg viewBox=""0 0 24 24""><path d=""M8 5v14l11-7z""/></svg>
                    <span>播放 / 暂停</span>
                </button>
            </div>

            <div class=""tip-status-text"">点击按键伴随触感震动反馈</div>
        </div>
    </div>

    <!-- 📷 Tab 2: Minimalist Visualizer Camera -->
    <div id=""pane-camera"" class=""tab-pane"">
        <div class=""visualizer-card"">
            
            <div class=""camera-viewport"">
                <video id=""cameraVideo"" autoplay playsinline muted></video>
                <canvas id=""captureCanvas"" style=""display:none;""></canvas>
            </div>

            <div class=""camera-action-grid"">
                <!-- Hero Snap Button -->
                <button class=""snap-hero-btn"" onclick=""snapAndCast()"">
                    <svg viewBox=""0 0 24 24""><circle cx=""12"" cy=""12"" r=""3.2""/><path d=""M9 2L7.17 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2h-3.17L15 2H9zm3 15c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5z""/></svg>
                    <span>📸 拍照并投屏讲评</span>
                </button>

                <button class=""flat-btn"" onclick=""switchCamera()"">
                    <span>🔄 翻转镜头</span>
                </button>
                <button id=""btnStream"" class=""flat-btn"" onclick=""toggleLiveStream()"">
                    <span>🎥 实时展台</span>
                </button>
                <button class=""flat-btn"" onclick=""rotateScreen()"">
                    <span>🔃 大屏旋转90°</span>
                </button>
                <button class=""flat-btn"" onclick=""toggleTorch()"">
                    <span>💡 补光灯</span>
                </button>

                <button class=""exit-btn"" onclick=""stopVisualizer()"">
                    <span>⏹️ 退出展台 · 恢复幻灯片</span>
                </button>
            </div>

        </div>
    </div>

    <script>
        // Haptic feedback
        function vibrate() {
            if (navigator.vibrate) navigator.vibrate(20);
        }

        // Tab Switcher
        function switchTab(name) {
            vibrate();
            const btns = document.querySelectorAll('.tab-btn');
            btns.forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));

            if (name === 'remote') {
                btns[0].classList.add('active');
                document.getElementById('pane-remote').classList.add('active');
            } else if (name === 'camera') {
                btns[1].classList.add('active');
                document.getElementById('pane-camera').classList.add('active');
                startCamera();
            }
        }

        // Remote command sender
        async function sendCmd(path) {
            vibrate();
            const led = document.getElementById('ledDot');
            if (led) {
                led.style.backgroundColor = '#0071e3';
                led.style.boxShadow = '0 0 8px #0071e3';
            }
            try {
                await fetch(path);
            } catch (e) {}
            setTimeout(() => {
                if (led) {
                    led.style.backgroundColor = 'var(--apple-green)';
                    led.style.boxShadow = '0 0 6px var(--apple-green)';
                }
            }, 180);
        }

        // Camera Visualizer
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
                alert('无法调用手机摄像头，请在浏览器权限设置中允许使用相机。');
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
                    alert('当前摄像头不支持闪光灯。');
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
                    alert('📸 照片已投射到 PPT 大屏！可在屏幕上使用画笔直接圈画批改。');
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
                btn.querySelector('span').innerText = '⏹️ 停止直播';
                btn.style.background = '#ffe5e5';
                btn.style.color = '#ff3b30';
                sendLiveFrame();
            } else {
                btn.querySelector('span').innerText = '🎥 实时展台';
                btn.style.background = 'var(--btn-neutral-bg)';
                btn.style.color = 'var(--text-primary)';
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
    </script>
</body>
</html>";
        }
    }
}
