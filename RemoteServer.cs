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
    // =========================================================================
    // Windows Native Core Audio Master Volume Interop (Zero External DLLs)
    // =========================================================================
    internal static class WindowsAudioHelper
    {
        [ComImport]
        [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        private class MMDeviceEnumeratorComObject { }

        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceEnumerator
        {
            int NotImpl1();
            [PreserveSig]
            int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppDevice);
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDevice
        {
            [PreserveSig]
            int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        }

        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IAudioEndpointVolume
        {
            [PreserveSig]
            int RegisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig]
            int UnregisterControlChangeNotify(IntPtr pNotify);
            [PreserveSig]
            int GetChannelCount(out uint pnChannelCount);
            [PreserveSig]
            int SetMasterVolumeLevel(float fLevelDB, ref Guid pguidEventContext);
            [PreserveSig]
            int SetMasterVolumeLevelScalar(float fLevel, ref Guid pguidEventContext);
            [PreserveSig]
            int GetMasterVolumeLevel(out float pfLevelDB);
            [PreserveSig]
            int GetMasterVolumeLevelScalar(out float pfLevel);
            [PreserveSig]
            int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, ref Guid pguidEventContext);
            [PreserveSig]
            int GetMute(out bool pbMute);
        }

        public static float GetMasterVolume()
        {
            try
            {
                var enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumeratorComObject());
                IMMDevice speakers;
                enumerator.GetDefaultAudioEndpoint(0, 1, out speakers);
                var IID_IAudioEndpointVolume = typeof(IAudioEndpointVolume).GUID;
                object o;
                speakers.Activate(ref IID_IAudioEndpointVolume, 0, IntPtr.Zero, out o);
                var audioEndpointVolume = (IAudioEndpointVolume)o;

                float currentVol;
                audioEndpointVolume.GetMasterVolumeLevelScalar(out currentVol);
                return currentVol;
            }
            catch
            {
                return 0.5f;
            }
        }

        public static void SetMasterVolume(float volume)
        {
            try
            {
                if (volume < 0f) volume = 0f;
                if (volume > 1f) volume = 1f;

                var enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumeratorComObject());
                IMMDevice speakers;
                enumerator.GetDefaultAudioEndpoint(0, 1, out speakers);
                var IID_IAudioEndpointVolume = typeof(IAudioEndpointVolume).GUID;
                object o;
                speakers.Activate(ref IID_IAudioEndpointVolume, 0, IntPtr.Zero, out o);
                var audioEndpointVolume = (IAudioEndpointVolume)o;

                Guid guid = Guid.Empty;
                audioEndpointVolume.SetMasterVolumeLevelScalar(volume, ref guid);
            }
            catch { }
        }
    }

    public class CastPhotoItem
    {
        public int Id { get; set; }
        public string Image { get; set; }
        public byte[] RawBytes { get; set; }
        public string ContentType { get; set; }
        public string Title { get; set; }
        public int Rotation { get; set; }
    }

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
        
        // Multi-Homework Photo Queue System
        public static List<CastPhotoItem> CastQueue = new List<CastPhotoItem>();
        public static int ActiveCastIndex = -1;
        public static int CastItemCounter = 0;
        public static int CastVersion = 0;
        public static string CastMode = "none"; // none, photo, stream

        private static bool _proxyEnabled = false;
        private static int _proxyPort = 7890;

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
                                score += 100;
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
                string localIp = GetBestLocalIpAddress();
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
                    
                    string clientIp = "127.0.0.1";
                    try
                    {
                        var endPoint = (IPEndPoint)client.Client.RemoteEndPoint;
                        clientIp = endPoint.Address.ToString();
                        if (clientIp.StartsWith("::ffff:")) clientIp = clientIp.Substring(7);
                    }
                    catch { }

                    using (var stream = client.GetStream())
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        string requestLine = reader.ReadLine();
                        if (string.IsNullOrEmpty(requestLine)) return;

                        string[] parts = requestLine.Split(' ');
                        if (parts.Length < 2) return;
                        
                        string method = parts[0].ToUpper();
                        string rawUrl = parts[1];
                        string path = rawUrl.ToLower();
                        int queryIdx = path.IndexOf('?');
                        if (queryIdx >= 0)
                        {
                            path = path.Substring(0, queryIdx);
                        }

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

                        string requestBody = "";
                        if (contentLength > 0 && contentLength < 35000000)
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

                        // =========================================================================
                        // API Routing
                        // =========================================================================
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
                        else if (path == "/first_slide")
                        {
                            GotoFirstSlide();
                            SendJsonResponse(stream, "{\"status\":\"success\"}");
                        }
                        else if (path == "/black")
                        {
                            TogglePowerPointBlackScreen();
                            SendJsonResponse(stream, "{\"status\":\"success\"}");
                        }
                        else if (path == "/api/get_system_volume")
                        {
                            float sysVol = WindowsAudioHelper.GetMasterVolume();
                            SendJsonResponse(stream, "{\"volume\":" + sysVol.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) + "}");
                        }
                        else if (path == "/volume")
                        {
                            string valStr = "";
                            int valIdx = rawUrl.IndexOf("val=");
                            if (valIdx >= 0)
                            {
                                valStr = rawUrl.Substring(valIdx + 4);
                            }
                            
                            if (!string.IsNullOrEmpty(valStr))
                            {
                                float vol = 0.5f;
                                if (float.TryParse(valStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out vol))
                                {
                                    WindowsAudioHelper.SetMasterVolume(vol);
                                    ControlPowerPointVolume(vol);
                                }
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\"}");
                        }
                        else if (path == "/api/set_proxy")
                        {
                            bool enable = rawUrl.Contains("enable=1") || rawUrl.Contains("enabled=true");
                            int port = 7890;
                            int portIdx = rawUrl.IndexOf("port=");
                            if (portIdx >= 0)
                            {
                                string portStr = rawUrl.Substring(portIdx + 5);
                                int endIdx = portStr.IndexOf('&');
                                if (endIdx >= 0) portStr = portStr.Substring(0, endIdx);
                                int.TryParse(portStr, out port);
                            }
                            if (port <= 0 || port > 65535) port = 7890;

                            _proxyEnabled = enable;
                            _proxyPort = port;

                            if (enable)
                            {
                                string proxyAddr = clientIp + ":" + port;
                                SetWindowsSystemProxy(proxyAddr, true);
                            }
                            else
                            {
                                SetWindowsSystemProxy("", false);
                            }

                            SendJsonResponse(stream, "{\"status\":\"success\",\"enabled\":" + (enable ? "true" : "false") + ",\"port\":" + port + ",\"clientIp\":\"" + clientIp + "\"}");
                        }
                        else if (path == "/api/get_proxy_status")
                        {
                            SendJsonResponse(stream, "{\"enabled\":" + (_proxyEnabled ? "true" : "false") + ",\"port\":" + _proxyPort + ",\"clientIp\":\"" + clientIp + "\"}");
                        }
                        else if (path == "/api/get_photo")
                        {
                            int id = -1;
                            int idPos = rawUrl.IndexOf("id=");
                            if (idPos >= 0)
                            {
                                string idStr = rawUrl.Substring(idPos + 3);
                                int endIdx = idStr.IndexOf('&');
                                if (endIdx >= 0) idStr = idStr.Substring(0, endIdx);
                                int.TryParse(idStr, out id);
                            }

                            CastPhotoItem target = null;
                            if (id > 0)
                            {
                                foreach (var it in CastQueue)
                                {
                                    if (it.Id == id) { target = it; break; }
                                }
                            }
                            if (target == null && ActiveCastIndex >= 0 && ActiveCastIndex < CastQueue.Count)
                            {
                                target = CastQueue[ActiveCastIndex];
                            }

                            if (target != null && target.RawBytes != null && target.RawBytes.Length > 0)
                            {
                                SendBinaryResponse(stream, target.RawBytes, target.ContentType ?? "image/jpeg");
                            }
                            else
                            {
                                string notFound = "HTTP/1.1 404 Not Found\r\nContent-Length: 0\r\n\r\n";
                                byte[] nfBytes = Encoding.UTF8.GetBytes(notFound);
                                stream.Write(nfBytes, 0, nfBytes.Length);
                            }
                        }
                        else if (path == "/api/cast_photo")
                        {
                            if (!string.IsNullOrEmpty(requestBody))
                            {
                                string imgData = ExtractJsonString(requestBody, "image");
                                if (!string.IsNullOrEmpty(imgData))
                                {
                                    CastItemCounter++;
                                    byte[] rawBytes = null;
                                    string contentType = "image/jpeg";
                                    try
                                    {
                                        int commaIdx = imgData.IndexOf(',');
                                        if (commaIdx >= 0)
                                        {
                                            string header = imgData.Substring(0, commaIdx);
                                            if (header.Contains("image/png")) contentType = "image/png";
                                            else if (header.Contains("image/webp")) contentType = "image/webp";
                                            string base64Payload = imgData.Substring(commaIdx + 1);
                                            rawBytes = Convert.FromBase64String(base64Payload);
                                        }
                                        else
                                        {
                                            rawBytes = Convert.FromBase64String(imgData);
                                        }
                                    }
                                    catch { }

                                    var item = new CastPhotoItem
                                    {
                                        Id = CastItemCounter,
                                        Image = imgData,
                                        RawBytes = rawBytes,
                                        ContentType = contentType,
                                        Title = "作业 " + (CastQueue.Count + 1),
                                        Rotation = 0
                                    };
                                    CastQueue.Add(item);
                                    ActiveCastIndex = CastQueue.Count - 1;
                                    CastMode = "photo";
                                    CastVersion++;
                                    
                                    if (Globals.ThisAddIn != null)
                                    {
                                        Globals.ThisAddIn.ShowLiveCast(ServerUrl + "cast_view");
                                    }
                                }
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\",\"version\":" + CastVersion + ",\"index\":" + ActiveCastIndex + ",\"total\":" + CastQueue.Count + "}");
                        }
                        else if (path == "/api/select_cast")
                        {
                            int idx = 0;
                            int idxPos = rawUrl.IndexOf("index=");
                            if (idxPos >= 0)
                            {
                                string idxStr = rawUrl.Substring(idxPos + 6);
                                int endIdx = idxStr.IndexOf('&');
                                if (endIdx >= 0) idxStr = idxStr.Substring(0, endIdx);
                                int.TryParse(idxStr, out idx);
                            }
                            if (idx >= 0 && idx < CastQueue.Count)
                            {
                                ActiveCastIndex = idx;
                                CastMode = "photo";
                                CastVersion++;
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\",\"version\":" + CastVersion + ",\"activeIndex\":" + ActiveCastIndex + "}");
                        }
                        else if (path == "/api/delete_cast")
                        {
                            int idx = 0;
                            int idxPos = rawUrl.IndexOf("index=");
                            if (idxPos >= 0)
                            {
                                string idxStr = rawUrl.Substring(idxPos + 6);
                                int endIdx = idxStr.IndexOf('&');
                                if (endIdx >= 0) idxStr = idxStr.Substring(0, endIdx);
                                int.TryParse(idxStr, out idx);
                            }
                            if (idx >= 0 && idx < CastQueue.Count)
                            {
                                CastQueue.RemoveAt(idx);
                                if (CastQueue.Count == 0)
                                {
                                    CastMode = "none";
                                    ActiveCastIndex = -1;
                                }
                                else if (ActiveCastIndex >= CastQueue.Count)
                                {
                                    ActiveCastIndex = CastQueue.Count - 1;
                                }
                                CastVersion++;
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\",\"version\":" + CastVersion + ",\"total\":" + CastQueue.Count + ",\"activeIndex\":" + ActiveCastIndex + "}");
                        }
                        else if (path == "/api/clear_cast")
                        {
                            CastQueue.Clear();
                            ActiveCastIndex = -1;
                            CastMode = "none";
                            CastVersion++;
                            SendJsonResponse(stream, "{\"status\":\"success\"}");
                        }
                        else if (path == "/api/stop_cast" || path == "/stop_cast")
                        {
                            CastMode = "none";
                            CastVersion++;
                            if (Globals.ThisAddIn != null)
                            {
                                Globals.ThisAddIn.CloseLiveCast();
                            }
                            SendJsonResponse(stream, "{\"status\":\"success\"}");
                        }
                        else if (path == "/api/rotate_cast")
                        {
                            if (ActiveCastIndex >= 0 && ActiveCastIndex < CastQueue.Count)
                            {
                                CastQueue[ActiveCastIndex].Rotation = (CastQueue[ActiveCastIndex].Rotation + 90) % 360;
                                CastVersion++;
                                SendJsonResponse(stream, "{\"status\":\"success\",\"rotation\":" + CastQueue[ActiveCastIndex].Rotation + "}");
                            }
                            else
                            {
                                SendJsonResponse(stream, "{\"status\":\"none\",\"rotation\":0}");
                            }
                        }
                        else if (path == "/api/get_cast_data" || path == "/api/get_cast_meta")
                        {
                            string json = BuildLightweightCastJson();
                            SendJsonResponse(stream, json);
                        }
                        else if (path == "/view_photo" || path == "/photo" || path == "/view")
                        {
                            string html = GetPhotoViewerHtml();
                            SendHtmlResponse(stream, html);
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
                            ProxyRequest(stream, method, CurrentProxyTarget, rawUrl);
                        }
                        else if (path == "/" || path == "/index.html")
                        {
                            if (!string.IsNullOrEmpty(CurrentProxyTarget))
                            {
                                ProxyRequest(stream, method, CurrentProxyTarget, rawUrl);
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

        // Lightweight Metadata JSON (< 250 bytes, zero CPU, instant 30ms latency)
        private static string BuildLightweightCastJson()
        {
            var sb = new StringBuilder();
            sb.Append("{");
            sb.AppendFormat("\"mode\":\"{0}\",\"version\":{1},\"total\":{2},\"activeIndex\":{3},", 
                CastMode, CastVersion, CastQueue.Count, ActiveCastIndex);

            if (ActiveCastIndex >= 0 && ActiveCastIndex < CastQueue.Count)
            {
                var cur = CastQueue[ActiveCastIndex];
                sb.AppendFormat("\"activeId\":{0},\"activeTitle\":\"{1}\",\"rotation\":{2},\"photoUrl\":\"/api/get_photo?id={0}&v={3}\",",
                    cur.Id, EscapeJson(cur.Title), cur.Rotation, CastVersion);
            }
            else
            {
                sb.Append("\"activeId\":-1,\"activeTitle\":\"\",\"rotation\":0,\"photoUrl\":\"\",");
            }

            sb.Append("\"items\":[");
            for (int i = 0; i < CastQueue.Count; i++)
            {
                if (i > 0) sb.Append(",");
                var it = CastQueue[i];
                sb.AppendFormat("{{\"index\":{0},\"id\":{1},\"title\":\"{2}\",\"rotation\":{3},\"url\":\"/api/get_photo?id={1}\"}}",
                    i, it.Id, EscapeJson(it.Title), it.Rotation);
            }
            sb.Append("]");
            sb.Append("}");
            return sb.ToString();
        }

        private static void SendBinaryResponse(Stream stream, byte[] data, string contentType)
        {
            string headers = "HTTP/1.1 200 OK\r\n" +
                             "Content-Type: " + contentType + "\r\n" +
                             "Access-Control-Allow-Origin: *\r\n" +
                             "Cache-Control: public, max-age=31536000\r\n" +
                             "Content-Length: " + data.Length + "\r\n" +
                             "Connection: close\r\n\r\n";
            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
            stream.Write(headerBytes, 0, headerBytes.Length);
            stream.Write(data, 0, data.Length);
        }

        private static void SendJsonResponse(Stream stream, string json)
        {
            byte[] bodyBytes = Encoding.UTF8.GetBytes(json);
            string headers = "HTTP/1.1 200 OK\r\n" +
                             "Content-Type: application/json; charset=utf-8\r\n" +
                             "Access-Control-Allow-Origin: *\r\n" +
                             "Cache-Control: no-cache, no-store\r\n" +
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
                        if (view.State == PowerPoint.PpSlideShowState.ppSlideShowPaused || 
                            view.State == PowerPoint.PpSlideShowState.ppSlideShowBlackScreen)
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

        private void GotoFirstSlide()
        {
            Globals.ThisAddIn.SafeInvoke(delegate()
            {
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app.SlideShowWindows.Count > 0)
                    {
                        var view = app.SlideShowWindows[1].View;
                        if (view.State == PowerPoint.PpSlideShowState.ppSlideShowBlackScreen)
                        {
                            view.State = PowerPoint.PpSlideShowState.ppSlideShowRunning;
                        }
                        view.First();
                    }
                }
                catch { }
            });
        }

        private void TogglePowerPointBlackScreen()
        {
            Globals.ThisAddIn.SafeInvoke(delegate()
            {
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app.SlideShowWindows.Count > 0)
                    {
                        var view = app.SlideShowWindows[1].View;
                        if (view.State == PowerPoint.PpSlideShowState.ppSlideShowBlackScreen)
                        {
                            view.State = PowerPoint.PpSlideShowState.ppSlideShowRunning;
                        }
                        else
                        {
                            view.State = PowerPoint.PpSlideShowState.ppSlideShowBlackScreen;
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
        // Student/Audience Photo Viewer HTML
        // Multi-Homework Photo Queue Support (Zero Control Buttons, Instant Loading)
        // =========================================================================
        private string GetPhotoViewerHtml()
        {
            return @"<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
    <meta charset=""utf-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes"">
    <title>实物作业展台</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            background-color: #0c0d10;
            color: #ffffff;
            font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Helvetica Neue', sans-serif;
            width: 100vw;
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: flex-start;
            overflow-x: hidden;
            padding: 16px;
        }
        .header-bar {
            width: 100%;
            max-width: 600px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 12px;
            padding: 0 4px;
        }
        .title {
            font-size: 17px;
            font-weight: 700;
            color: #f5f5f7;
            letter-spacing: -0.3px;
        }
        .queue-bar {
            width: 100%;
            max-width: 600px;
            display: flex;
            gap: 8px;
            overflow-x: auto;
            padding: 4px 0 12px 0;
        }
        .queue-pill {
            background: #1c1c1e;
            border: 1px solid rgba(255, 255, 255, 0.12);
            color: #8e8e93;
            padding: 6px 14px;
            border-radius: 14px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            white-space: nowrap;
            transition: all 0.15s ease;
        }
        .queue-pill.active {
            background: #0071e3;
            color: #ffffff;
            border-color: #0071e3;
            box-shadow: 0 2px 8px rgba(0, 113, 227, 0.4);
        }
        .img-card {
            width: 100%;
            max-width: 600px;
            background: #18191c;
            border-radius: 18px;
            overflow: hidden;
            box-shadow: 0 12px 40px rgba(0, 0, 0, 0.6);
            border: 1px solid rgba(255, 255, 255, 0.08);
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 320px;
        }
        #photoImg {
            width: 100%;
            height: auto;
            max-height: 80vh;
            object-fit: contain;
            display: none;
            border-radius: 18px;
            transition: transform 0.2s ease;
        }
        .empty-box {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 12px;
            color: #8e8e93;
            padding: 40px 20px;
            text-align: center;
        }
        .empty-box svg { width: 48px; height: 48px; fill: #636366; }
        .hint-bar {
            margin-top: 14px;
            font-size: 13px;
            color: #636366;
            text-align: center;
        }
    </style>
</head>
<body>
    <div class=""header-bar"">
        <div id=""lblHeaderTitle"" class=""title"">作业试卷讲评</div>
    </div>

    <!-- Homework Queue Selector -->
    <div id=""queueBar"" class=""queue-bar"" style=""display: none;""></div>

    <div class=""img-card"">
        <div id=""emptyHint"" class=""empty-box"">
            <svg viewBox=""0 0 24 24""><circle cx=""12"" cy=""12"" r=""3.2""/><path d=""M9 2L7.17 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2h-3.17L15 2H9zm3 15c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5z""/></svg>
            <div>教师尚未上传实物照片</div>
        </div>
        <img id=""photoImg"" alt=""Homework Photo"" />
    </div>

    <div class=""hint-bar"">双指捏合缩放 · 长按可保存原图</div>

    <script>
        let curVer = -1;
        let curPhotoId = -1;

        function renderQueue(items, activeIdx) {
            const bar = document.getElementById('queueBar');
            if (!items || items.length <= 1) {
                bar.style.display = 'none';
                return;
            }
            bar.style.display = 'flex';
            bar.innerHTML = '';
            items.forEach((it, idx) => {
                const btn = document.createElement('div');
                btn.className = 'queue-pill' + (idx === activeIdx ? ' active' : '');
                btn.innerText = it.title || ('作业 ' + (idx + 1));
                btn.onclick = () => selectHomework(idx);
                bar.appendChild(btn);
            });
        }

        async function selectHomework(idx) {
            try {
                await fetch('/api/select_cast?index=' + idx);
            } catch (e) {}
        }

        async function fetchPhoto() {
            try {
                const res = await fetch('/api/get_cast_meta');
                if (res.ok) {
                    const data = await res.json();
                    if (data.version !== curVer) {
                        curVer = data.version;
                        const img = document.getElementById('photoImg');
                        const empty = document.getElementById('emptyHint');
                        const title = document.getElementById('lblHeaderTitle');

                        if (data.mode === 'photo' && data.activeId > 0) {
                            if (data.activeId !== curPhotoId) {
                                curPhotoId = data.activeId;
                                img.src = data.photoUrl;
                            }
                            img.style.transform = 'rotate(' + (data.rotation || 0) + 'deg)';
                            img.style.display = 'block';
                            empty.style.display = 'none';
                            title.innerText = (data.activeTitle || '作业讲评') + ' (' + (data.activeIndex + 1) + '/' + data.total + ')';
                            renderQueue(data.items, data.activeIndex);
                        } else {
                            curPhotoId = -1;
                            img.style.display = 'none';
                            empty.style.display = 'flex';
                            title.innerText = '作业试卷讲评';
                            renderQueue([], 0);
                        }
                    }
                }
            } catch (e) {}
            setTimeout(fetchPhoto, 250);
        }
        fetchPhoto();
    </script>
</body>
</html>";
        }

        // =========================================================================
        // Full Screen Visualizer Screen Receiver HTML (Large Screen in PPT)
        // With Multi-Homework Badge + Vector Rotate Button (Zero-Lag Metadata Polling)
        // =========================================================================
        private string GetCastViewHtml()
        {
            return @"<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
    <meta charset=""utf-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes"">
    <title>YuLink 展台</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; user-select: none; -webkit-user-select: none; }
        body {
            background-color: #000000;
            color: #ffffff;
            font-family: -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'SF Pro Text', 'Helvetica Neue', Arial, sans-serif;
            width: 100vw;
            height: 100vh;
            overflow: hidden;
        }
        #viewport {
            width: 100vw;
            height: 100vh;
            position: relative;
            overflow: hidden;
            cursor: grab;
            display: flex;
            align-items: center;
            justify-content: center;
            touch-action: none;
        }
        #viewport.dragging {
            cursor: grabbing;
        }
        #panLayer {
            position: absolute;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            will-change: transform;
            pointer-events: none;
        }
        #rotateScaleLayer {
            display: flex;
            align-items: center;
            justify-content: center;
            will-change: transform;
            pointer-events: none;
        }
        #castImage {
            max-width: 96vw;
            max-height: 94vh;
            object-fit: contain;
            border-radius: 8px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.9);
            pointer-events: none;
            -webkit-user-drag: none;
        }
        .hud-badge {
            position: absolute;
            top: 18px;
            left: 20px;
            background: rgba(0, 0, 0, 0.65);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1px solid rgba(255, 255, 255, 0.15);
            color: #ffffff;
            padding: 6px 14px;
            border-radius: 16px;
            font-size: 13px;
            font-weight: 600;
            display: none;
            z-index: 100;
            letter-spacing: -0.2px;
            box-shadow: 0 4px 16px rgba(0,0,0,0.4);
        }
        .hud-tools {
            position: absolute;
            top: 18px;
            right: 20px;
            display: flex;
            gap: 10px;
            z-index: 100;
        }
        .hud-circle-btn {
            width: 40px;
            height: 40px;
            border-radius: 50%;
            background: rgba(255, 255, 255, 0.92);
            backdrop-filter: blur(20px);
            -webkit-backdrop-filter: blur(20px);
            border: 1px solid rgba(0, 0, 0, 0.1);
            color: #1d1d1f;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            box-shadow: 0 3px 12px rgba(0,0,0,0.25);
            transition: all 0.15s cubic-bezier(0.16, 1, 0.3, 1);
        }
        .hud-circle-btn:hover {
            transform: scale(1.08);
            background: #ffffff;
        }
        .hud-circle-btn:active {
            transform: scale(0.92);
        }
        .hud-circle-btn svg {
            width: 20px;
            height: 20px;
            fill: #1d1d1f;
        }
        
        .empty-hint {
            color: #8e8e93;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 10px;
            pointer-events: none;
        }
        .empty-hint svg { width: 48px; height: 48px; fill: #636366; }
    </style>
</head>
<body>
    <div id=""viewport"">
        <!-- Order / Queue Badge -->
        <div id=""orderBadge"" class=""hud-badge"">作业 1 (1/1)</div>

        <div class=""hud-tools"">
            <!-- Pure Apple Vector Rotate Button -->
            <button class=""hud-circle-btn"" onclick=""rotateImage()"" title=""旋转 90°"">
                <svg viewBox=""0 0 24 24""><path d=""M7.11 8.53L5.7 7.11C4.8 8.27 4.24 9.61 4.07 11h2.02c.14-.87.49-1.72 1.02-2.47zM6.09 13H4.07c.17 1.39.72 2.73 1.62 3.89l1.41-1.42c-.52-.75-.87-1.6-1.01-2.47zm1.01 5.32c1.16.9 2.51 1.44 3.9 1.61V17.9c-.87-.15-1.71-.49-2.46-1.03L7.1 18.32zM13 4.07V1L8.45 5.55 13 10V6.09c3.37.5 6 3.41 6 6.91 0 1.25-.34 2.43-.93 3.44l1.46 1.46C20.45 16.38 21 14.77 21 13c0-4.42-3.21-8.08-7.39-8.73z""/></svg>
            </button>
        </div>

        <div id=""emptyState"" class=""empty-hint"">
            <svg viewBox=""0 0 24 24""><path d=""M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8zm-1-13h2v6h-2zm0 8h2v2h-2z""/></svg>
            <div style=""font-size: 15px; font-weight: 600;"">等待拍照</div>
        </div>

        <div id=""panLayer"">
            <div id=""rotateScaleLayer"">
                <img id=""castImage"" style=""display: none;"" alt=""Cast Preview"" />
            </div>
        </div>
    </div>

    <script>
        let currentVersion = -1;
        let currentRotation = 0;
        let currentScale = 1.0;
        let posX = 0;
        let posY = 0;
        let currentPhotoId = -1;

        const viewport = document.getElementById('viewport');
        const panLayer = document.getElementById('panLayer');
        const rotateScaleLayer = document.getElementById('rotateScaleLayer');
        const img = document.getElementById('castImage');
        const emptyState = document.getElementById('emptyState');
        const orderBadge = document.getElementById('orderBadge');

        function updatePan() {
            panLayer.style.transform = 'translate3d(' + posX + 'px, ' + posY + 'px, 0)';
        }

        function updateTransform() {
            rotateScaleLayer.style.transform = 'rotate(' + currentRotation + 'deg) scale(' + currentScale + ')';
        }

        function resetView() {
            currentScale = 1.0;
            posX = 0;
            posY = 0;
            updatePan();
            updateTransform();
        }

        async function rotateImage() {
            currentRotation = (currentRotation + 90) % 360;
            updateTransform();
            try {
                await fetch('/api/rotate_cast', { method: 'POST' });
            } catch (e) {}
        }

        // Pointer Drag Engine (Press to Drag -> Release to Stop)
        let isDown = false;
        let startX = 0;
        let startY = 0;
        let initialDist = 0;
        let initialScale = 1;

        viewport.addEventListener('pointerdown', (e) => {
            if (e.target.tagName === 'BUTTON' || e.target.closest('button')) return;
            e.preventDefault();
            isDown = true;
            viewport.classList.add('dragging');
            viewport.setPointerCapture(e.pointerId);
            startX = e.clientX - posX;
            startY = e.clientY - posY;
        });

        viewport.addEventListener('pointermove', (e) => {
            if (!isDown) return;
            posX = e.clientX - startX;
            posY = e.clientY - startY;
            updatePan();
        });

        viewport.addEventListener('pointerup', (e) => {
            if (isDown) {
                isDown = false;
                viewport.classList.remove('dragging');
                try { viewport.releasePointerCapture(e.pointerId); } catch (err) {}
            }
        });

        viewport.addEventListener('pointercancel', (e) => {
            isDown = false;
            viewport.classList.remove('dragging');
        });

        // Mouse Wheel Zoom
        viewport.addEventListener('wheel', (e) => {
            e.preventDefault();
            const delta = e.deltaY < 0 ? 1.15 : 0.85;
            currentScale = Math.max(0.4, Math.min(6.0, currentScale * delta));
            updateTransform();
        }, { passive: false });

        // Touch Pinch-to-Zoom
        viewport.addEventListener('touchstart', (e) => {
            if (e.touches.length === 2) {
                isDown = false;
                initialDist = Math.hypot(
                    e.touches[0].clientX - e.touches[1].clientX,
                    e.touches[0].clientY - e.touches[1].clientY
                );
                initialScale = currentScale;
            }
        }, { passive: true });

        viewport.addEventListener('touchmove', (e) => {
            if (e.touches.length === 2 && initialDist > 0) {
                const dist = Math.hypot(
                    e.touches[0].clientX - e.touches[1].clientX,
                    e.touches[0].clientY - e.touches[1].clientY
                );
                currentScale = Math.max(0.4, Math.min(6.0, initialScale * (dist / initialDist)));
                updateTransform();
            }
        }, { passive: true });

        viewport.addEventListener('touchend', (e) => {
            if (e.touches.length < 2) initialDist = 0;
        });

        // Polling Cast Updates (Ultra-Fast 100ms Metadata Check)
        async function pollCastData() {
            try {
                const res = await fetch('/api/get_cast_meta');
                if (res.ok) {
                    const data = await res.json();
                    if (data.mode === 'none' || data.activeId <= 0) {
                        emptyState.style.display = 'flex';
                        img.style.display = 'none';
                        orderBadge.style.display = 'none';
                        currentPhotoId = -1;
                    } else if (data.version !== currentVersion) {
                        currentVersion = data.version;
                        if (data.activeId !== currentPhotoId) {
                            currentPhotoId = data.activeId;
                            resetView();
                            img.src = data.photoUrl;
                        }
                        currentRotation = data.rotation || 0;
                        img.style.display = 'block';
                        emptyState.style.display = 'none';
                        updateTransform();

                        if (data.total > 1) {
                            orderBadge.innerText = (data.activeTitle || '作业') + ' (' + (data.activeIndex + 1) + '/' + data.total + ')';
                            orderBadge.style.display = 'block';
                        } else {
                            orderBadge.style.display = 'none';
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
        // Pure Minimalist Modern Apple Pro Web Controller HTML
        // Instant Optimistic Queue Switching + Direct Binary Upload
        // =========================================================================
        private string GetControlPageHtml()
        {
            return @"<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
    <meta charset=""utf-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover"">
    <title>YuLink 控制中心</title>
    <style>
        :root {
            --bg-page: #f2f2f7;
            --card-bg: #ffffff;
            --card-border: rgba(0, 0, 0, 0.05);
            --card-shadow: 0 8px 28px -6px rgba(0, 0, 0, 0.04), 0 2px 4px rgba(0, 0, 0, 0.02);
            --text-primary: #1d1d1f;
            --text-secondary: #8e8e93;
            --btn-neutral-bg: #f2f2f7;
            --btn-neutral-active: #e5e5ea;
            --apple-blue: #0071e3;
            --apple-blue-active: #0062c4;
            --apple-green: #34c759;
            --apple-green-active: #2ebd52;
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
            padding: 18px 16px 36px 16px;
            -webkit-font-smoothing: antialiased;
        }

        /* Top Minimalist Header */
        .app-header {
            width: 100%;
            max-width: 380px;
            padding: 6px 4px 18px 4px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .brand-title {
            font-size: 21px;
            font-weight: 700;
            letter-spacing: -0.5px;
            color: var(--text-primary);
        }

        .status-badge {
            display: flex;
            align-items: center;
            gap: 7px;
            background: #ffffff;
            padding: 6px 13px;
            border-radius: 14px;
            font-size: 12px;
            font-weight: 500;
            color: var(--text-secondary);
            border: 1px solid var(--card-border);
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03);
        }

        .status-dot {
            width: 7px;
            height: 7px;
            background: var(--apple-green);
            border-radius: 50%;
            box-shadow: 0 0 6px var(--apple-green);
        }

        /* Apple iOS Segmented Control Switcher */
        .tab-bar-container {
            width: 100%;
            max-width: 380px;
            margin-bottom: 18px;
            background: #e5e5ea;
            border-radius: 15px;
            padding: 3px;
            display: flex;
            position: relative;
        }

        .tab-btn {
            flex: 1;
            padding: 9px 0;
            text-align: center;
            font-size: 14px;
            font-weight: 600;
            border-radius: 12px;
            color: var(--text-secondary);
            cursor: pointer;
            transition: all 0.2s cubic-bezier(0.16, 1, 0.3, 1);
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
            box-shadow: 0 3px 8px rgba(0, 0, 0, 0.1), 0 1px 2px rgba(0, 0, 0, 0.04);
        }

        /* Tab Content Panes */
        .tab-pane {
            display: none;
            width: 100%;
            max-width: 380px;
            flex-direction: column;
            align-items: center;
            gap: 16px;
            animation: paneFadeIn 0.2s cubic-bezier(0.16, 1, 0.3, 1) forwards;
        }

        .tab-pane.active {
            display: flex;
        }

        @keyframes paneFadeIn {
            from { opacity: 0; transform: translateY(3px); }
            to { opacity: 1; transform: translateY(0); }
        }

        /* Pure White Card Base */
        .apple-card {
            width: 100%;
            background: var(--card-bg);
            border-radius: 26px;
            border: 1px solid var(--card-border);
            padding: 22px 18px;
            display: flex;
            flex-direction: column;
            box-shadow: var(--card-shadow);
            gap: 14px;
        }

        .card-header-title {
            font-size: 13px;
            font-weight: 600;
            letter-spacing: -0.2px;
            color: var(--text-secondary);
        }

        /* 🎮 Tab 1: Navigation Controls */
        .next-hero-btn {
            width: 100%;
            height: 98px;
            background: var(--apple-blue);
            border-radius: 20px;
            border: none;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
            color: #ffffff;
            cursor: pointer;
            transition: all 0.12s ease;
            box-shadow: 0 6px 18px rgba(0, 113, 227, 0.24);
        }

        .next-hero-btn:active {
            transform: scale(0.98);
            background: var(--apple-blue-active);
        }

        .next-hero-btn svg {
            width: 26px;
            height: 26px;
            fill: #ffffff;
        }

        .next-hero-btn span {
            font-size: 18px;
            font-weight: 600;
            letter-spacing: -0.3px;
        }

        .nav-grid-two {
            width: 100%;
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
        }

        .flat-btn {
            height: 52px;
            background: var(--btn-neutral-bg);
            border-radius: 16px;
            border: none;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 7px;
            font-size: 14px;
            font-weight: 600;
            color: var(--text-primary);
            cursor: pointer;
            transition: all 0.12s ease;
        }

        .flat-btn:active {
            transform: scale(0.97);
            background: var(--btn-neutral-active);
        }

        .flat-btn svg {
            width: 17px;
            height: 17px;
            fill: var(--text-primary);
        }

        /* iOS Control Center Volume Capsule Slider */
        .volume-header-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
        }

        .volume-badge {
            font-size: 13px;
            font-weight: 600;
            color: var(--apple-blue);
        }

        .volume-capsule-wrapper {
            position: relative;
            width: 100%;
            height: 54px;
            background: #e5e5ea;
            border-radius: 18px;
            overflow: hidden;
            cursor: pointer;
            display: flex;
            align-items: center;
            touch-action: none;
        }

        .volume-capsule-fill {
            position: absolute;
            left: 0;
            top: 0;
            height: 100%;
            width: 50%;
            background: var(--apple-blue);
            border-radius: 18px;
            pointer-events: none;
            will-change: width;
        }

        .volume-capsule-content {
            position: absolute;
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 0 16px;
            pointer-events: none;
            z-index: 2;
        }

        .volume-icon-left {
            width: 20px;
            height: 20px;
            fill: #ffffff;
            filter: drop-shadow(0 1px 2px rgba(0,0,0,0.2));
        }

        .volume-icon-right {
            width: 20px;
            height: 20px;
            fill: var(--text-secondary);
        }

        /* 📶 Minimalist Network Sharing Card */
        .proxy-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
            width: 100%;
        }

        .proxy-title {
            font-size: 15px;
            font-weight: 600;
            color: var(--text-primary);
            cursor: pointer;
        }

        /* iOS Toggle Switch */
        .ios-switch {
            position: relative;
            display: inline-block;
            width: 51px;
            height: 31px;
        }

        .ios-switch input {
            opacity: 0;
            width: 0;
            height: 0;
        }

        .ios-slider {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: #e5e5ea;
            transition: .25s cubic-bezier(0.16, 1, 0.3, 1);
            border-radius: 31px;
        }

        .ios-slider:before {
            position: absolute;
            content: """";
            height: 27px;
            width: 27px;
            left: 2px;
            bottom: 2px;
            background-color: white;
            transition: .25s cubic-bezier(0.16, 1, 0.3, 1);
            border-radius: 50%;
            box-shadow: 0 2px 5px rgba(0, 0, 0, 0.2);
        }

        input:checked + .ios-slider {
            background-color: var(--apple-green);
        }

        input:checked + .ios-slider:before {
            transform: translateX(20px);
        }

        /* 📷 Tab 2: Visualizer Camera & Homework Queue */
        .queue-scroll-container {
            display: flex;
            gap: 8px;
            overflow-x: auto;
            padding: 4px 2px;
            width: 100%;
            -webkit-overflow-scrolling: touch;
        }

        .queue-scroll-container::-webkit-scrollbar { display: none; }

        .queue-item-pill {
            display: flex;
            align-items: center;
            gap: 6px;
            background: #f2f2f7;
            padding: 7px 13px;
            border-radius: 14px;
            font-size: 13px;
            font-weight: 600;
            color: var(--text-secondary);
            cursor: pointer;
            white-space: nowrap;
            transition: all 0.12s ease;
            border: 1px solid transparent;
        }

        .queue-item-pill.active {
            background: var(--apple-blue);
            color: #ffffff;
            box-shadow: 0 3px 10px rgba(0, 113, 227, 0.3);
        }

        .snap-hero-wrapper {
            position: relative;
            width: 100%;
        }

        .snap-hero-btn {
            width: 100%;
            height: 72px;
            background: var(--apple-green);
            border-radius: 20px;
            border: none;
            color: #ffffff;
            font-size: 16px;
            font-weight: 600;
            letter-spacing: -0.3px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            cursor: pointer;
            transition: all 0.12s ease;
            box-shadow: 0 6px 18px rgba(52, 199, 89, 0.25);
        }

        .snap-hero-btn:active {
            transform: scale(0.98);
            background: var(--apple-green-active);
        }

        .snap-hero-btn svg {
            width: 22px;
            height: 22px;
            fill: #ffffff;
        }

        #nativeCameraInput {
            position: absolute;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            opacity: 0;
            cursor: pointer;
            z-index: 10;
        }

        .preview-box {
            width: 100%;
            height: 180px;
            background: #f8f8fa;
            border: 1px dashed #d1d1d6;
            border-radius: 18px;
            overflow: hidden;
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
        }

        #previewThumb {
            width: 100%;
            height: 100%;
            object-fit: contain;
            display: none;
        }

        .preview-placeholder {
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 8px;
            color: var(--text-secondary);
            font-size: 13px;
            font-weight: 500;
        }

        .preview-placeholder svg {
            width: 32px;
            height: 32px;
            fill: #aeaeb2;
        }

        .exit-btn {
            width: 100%;
            height: 48px;
            background: var(--apple-red-bg);
            border: 1px solid rgba(255, 59, 48, 0.12);
            border-radius: 16px;
            color: var(--apple-red);
            font-size: 14px;
            font-weight: 600;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            cursor: pointer;
            transition: all 0.12s ease;
        }

        .exit-btn:active {
            transform: scale(0.97);
            background: rgba(255, 59, 48, 0.16);
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

    <!-- Apple iOS Segmented Control Switcher -->
    <div class=""tab-bar-container"">
        <button class=""tab-btn active"" onclick=""switchTab('remote')"">遥控翻页</button>
        <button class=""tab-btn"" onclick=""switchTab('camera')"">实物展台</button>
    </div>

    <!-- 🎮 Tab 1: Minimalist Remote Controller Pane -->
    <div id=""pane-remote"" class=""tab-pane active"">
        
        <!-- Primary Slideshow Control Card -->
        <div class=""apple-card"">
            <div class=""card-header-title"">幻灯片放映控制</div>

            <button class=""next-hero-btn"" onclick=""sendCmd('/next')"">
                <svg viewBox=""0 0 24 24""><path d=""M6 18l8.5-6L6 6v12zM16 6v12h2V6h-2z""/></svg>
                <span>下一页 (Next)</span>
            </button>

            <div class=""nav-grid-two"">
                <button class=""flat-btn"" onclick=""sendCmd('/prev')"">
                    <svg viewBox=""0 0 24 24""><path d=""M6 6h2v12H6zm3.5 6l8.5 6V6z""/></svg>
                    <span>上一页</span>
                </button>
                <button class=""flat-btn"" onclick=""sendCmd('/play')"">
                    <svg viewBox=""0 0 24 24""><path d=""M8 5v14l11-7z""/></svg>
                    <span>播放 / 暂停</span>
                </button>
            </div>

            <!-- Practical Presentation Utilities -->
            <div class=""nav-grid-two"">
                <button class=""flat-btn"" onclick=""sendCmd('/black')"">
                    <svg viewBox=""0 0 24 24""><path d=""M21 3H3c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h18c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H3V5h18v14z""/></svg>
                    <span>一键黑屏</span>
                </button>
                <button class=""flat-btn"" onclick=""sendCmd('/first_slide')"">
                    <svg viewBox=""0 0 24 24""><path d=""M12 5V1L7 6l5 5V7c3.31 0 6 2.69 6 6s-2.69 6-6 6-6-2.69-6-6H4c0 4.42 3.58 8 8 8s8-3.58 8-8-3.58-8-8-8z""/></svg>
                    <span>返回第一页</span>
                </button>
            </div>
        </div>

        <!-- Windows Native Master Volume Card -->
        <div class=""apple-card"">
            <div class=""volume-header-row"">
                <span class=""card-header-title"">系统主音量</span>
                <span id=""lblVolume"" class=""volume-badge"">50%</span>
            </div>

            <!-- iOS Control Center Fluid Drag Capsule Slider -->
            <div id=""volCapsule"" class=""volume-capsule-wrapper"">
                <div id=""volFill"" class=""volume-capsule-fill"" style=""width: 50%;""></div>
                <div class=""volume-capsule-content"">
                    <svg class=""volume-icon-left"" viewBox=""0 0 24 24""><path d=""M7 9v6h4l5 5V4L11 9H7z""/></svg>
                    <svg class=""volume-icon-right"" viewBox=""0 0 24 24""><path d=""M3 9v6h4l5 5V4L7 9H3zm13.5 3c0-1.77-1.02-3.29-2.5-4.03v8.05c1.48-.73 2.5-2.25 2.5-4.02zM14 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77z""/></svg>
                </div>
            </div>
        </div>

        <!-- 📶 Minimalist Network Sharing Card -->
        <div class=""apple-card"">
            <div class=""proxy-row"">
                <div class=""proxy-title"" onclick=""configProxyPort()"">网络共享</div>
                <label class=""ios-switch"">
                    <input id=""proxySwitch"" type=""checkbox"" onchange=""toggleProxy(this.checked)"">
                    <span class=""ios-slider""></span>
                </label>
            </div>
        </div>

    </div>

    <!-- 📷 Tab 2: Visualizer Camera & Homework Queue Pane -->
    <div id=""pane-camera"" class=""tab-pane"">
        
        <div class=""apple-card"">
            <div class=""card-header-title"" id=""lblQueueStatus"">作业拍摄队列</div>

            <!-- Homework Queue Pills -->
            <div id=""mobileQueueList"" class=""queue-scroll-container"" style=""display: none;""></div>

            <!-- Native Optical Camera Trigger (Apple Green Hero) -->
            <div class=""snap-hero-wrapper"">
                <button class=""snap-hero-btn"">
                    <svg viewBox=""0 0 24 24""><circle cx=""12"" cy=""12"" r=""3.2""/><path d=""M9 2L7.17 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2h-3.17L15 2H9zm3 15c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5z""/></svg>
                    <span id=""btnSnapText"">拍摄新作业</span>
                </button>
                <input id=""nativeCameraInput"" type=""file"" accept=""image/*"" capture=""environment"" onchange=""handleNativePhoto(this)"" />
            </div>

            <!-- Photo Preview Box -->
            <div class=""preview-box"">
                <img id=""previewThumb"" alt=""Photo Preview"" />
                <div id=""previewPlaceholder"" class=""preview-placeholder"">
                    <svg viewBox=""0 0 24 24""><path d=""M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z""/></svg>
                    <span>拍摄后在此预览并同步大屏</span>
                </div>
            </div>

            <div class=""nav-grid-two"">
                <button class=""flat-btn"" onclick=""rotateScreen()"">
                    <svg viewBox=""0 0 24 24""><path d=""M7.11 8.53L5.7 7.11C4.8 8.27 4.24 9.61 4.07 11h2.02c.14-.87.49-1.72 1.02-2.47zM6.09 13H4.07c.17 1.39.72 2.73 1.62 3.89l1.41-1.42c-.52-.75-.87-1.6-1.01-2.47zm1.01 5.32c1.16.9 2.51 1.44 3.9 1.61V17.9c-.87-.15-1.71-.49-2.46-1.03L7.1 18.32zM13 4.07V1L8.45 5.55 13 10V6.09c3.37.5 6 3.41 6 6.91 0 1.25-.34 2.43-.93 3.44l1.46 1.46C20.45 16.38 21 14.77 21 13c0-4.42-3.21-8.08-7.39-8.73z""/></svg>
                    <span>旋转 90°</span>
                </button>
                <button class=""flat-btn"" onclick=""deleteCurrentPhoto()"">
                    <svg viewBox=""0 0 24 24""><path d=""M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z""/></svg>
                    <span>删除本张</span>
                </button>
            </div>

            <button class=""exit-btn"" onclick=""stopVisualizer()"">
                <span>退出展台 · 恢复放映</span>
            </button>
        </div>

    </div>

    <script>
        function vibrate() {
            if (navigator.vibrate) navigator.vibrate(20);
        }

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
                pollQueueData();
            }
        }

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

        // High-Precision Real-time Audio Volume Dragging Engine
        const capsule = document.getElementById('volCapsule');
        const fill = document.getElementById('volFill');
        const badge = document.getElementById('lblVolume');
        let isDraggingVol = false;
        let isSendingVol = false;
        let pendingVol = null;

        function updateVolumeUI(ratio) {
            ratio = Math.max(0, Math.min(1, ratio));
            const percent = Math.round(ratio * 100);
            fill.style.width = percent + '%';
            badge.innerText = percent + '%';
        }

        function dispatchVolumeToServer(ratio) {
            pendingVol = ratio;
            if (isSendingVol) return;

            isSendingVol = true;
            const toSend = pendingVol;
            pendingVol = null;

            fetch('/volume?val=' + toSend.toFixed(2))
                .finally(() => {
                    isSendingVol = false;
                    if (pendingVol !== null) {
                        dispatchVolumeToServer(pendingVol);
                    }
                });
        }

        function handleCapsulePointer(e) {
            const rect = capsule.getBoundingClientRect();
            const clientX = (e.touches && e.touches.length > 0) ? e.touches[0].clientX : e.clientX;
            let ratio = (clientX - rect.left) / rect.width;
            ratio = Math.max(0, Math.min(1, ratio));

            updateVolumeUI(ratio);
            dispatchVolumeToServer(ratio);
        }

        capsule.addEventListener('mousedown', (e) => {
            isDraggingVol = true;
            vibrate();
            handleCapsulePointer(e);
        });

        capsule.addEventListener('touchstart', (e) => {
            isDraggingVol = true;
            vibrate();
            handleCapsulePointer(e);
        }, { passive: false });

        window.addEventListener('mousemove', (e) => {
            if (isDraggingVol) handleCapsulePointer(e);
        });

        window.addEventListener('touchmove', (e) => {
            if (isDraggingVol) {
                e.preventDefault();
                handleCapsulePointer(e);
            }
        }, { passive: false });

        window.addEventListener('mouseup', () => { isDraggingVol = false; });
        window.addEventListener('touchend', () => { isDraggingVol = false; });

        async function fetchSystemVolume() {
            try {
                const res = await fetch('/api/get_system_volume');
                if (res.ok) {
                    const data = await res.json();
                    if (typeof data.volume === 'number') {
                        updateVolumeUI(data.volume);
                    }
                }
            } catch (e) {}
        }
        fetchSystemVolume();

        // Pure Network Sharing Toggle & Port Config
        let currentProxyPort = 7890;

        async function initProxyStatus() {
            try {
                const res = await fetch('/api/get_proxy_status');
                if (res.ok) {
                    const data = await res.json();
                    document.getElementById('proxySwitch').checked = data.enabled;
                    currentProxyPort = data.port || 7890;
                }
            } catch (e) {}
        }
        initProxyStatus();

        async function toggleProxy(enabled) {
            vibrate();
            try {
                await fetch('/api/set_proxy?enable=' + (enabled ? '1' : '0') + '&port=' + currentProxyPort);
            } catch (e) {
                alert('设置失败: ' + e.message);
            }
        }

        function configProxyPort() {
            const newPort = prompt('设置网络共享代理端口 (默认 7890):', currentProxyPort);
            if (newPort) {
                const p = parseInt(newPort);
                if (p > 0 && p <= 65535) {
                    currentProxyPort = p;
                    const isChecked = document.getElementById('proxySwitch').checked;
                    toggleProxy(isChecked);
                } else {
                    alert('端口号不合法 (1-65535)');
                }
            }
        }

        // =========================================================================
        // Homework Queue Management Engine (Mobile Control - Zero Lag)
        // =========================================================================
        let queueVersion = -1;
        let currentActiveIdx = -1;
        let localActivePhotoId = -1;

        async function pollQueueData() {
            try {
                const res = await fetch('/api/get_cast_meta');
                if (res.ok) {
                    const data = await res.json();
                    if (data.version !== queueVersion) {
                        queueVersion = data.version;
                        currentActiveIdx = data.activeIndex;
                        renderMobileQueue(data);
                    }
                }
            } catch (e) {}
            setTimeout(pollQueueData, 200);
        }

        function renderMobileQueue(data) {
            const qList = document.getElementById('mobileQueueList');
            const qStatus = document.getElementById('lblQueueStatus');
            const thumb = document.getElementById('previewThumb');
            const placeholder = document.getElementById('previewPlaceholder');
            const btnSnap = document.getElementById('btnSnapText');

            if (!data.items || data.items.length === 0) {
                qList.style.display = 'none';
                qStatus.innerText = '实物展台 · 拍照讲评作业';
                thumb.style.display = 'none';
                placeholder.style.display = 'flex';
                btnSnap.innerText = '拍照投屏讲评';
                localActivePhotoId = -1;
                return;
            }

            qList.style.display = 'flex';
            qList.innerHTML = '';
            qStatus.innerText = '已拍摄 ' + data.total + ' 份作业 (当前: ' + (data.activeIndex + 1) + ')';
            btnSnap.innerText = '+ 拍摄下一份作业';

            data.items.forEach((it, idx) => {
                const pill = document.createElement('div');
                pill.className = 'queue-item-pill' + (idx === data.activeIndex ? ' active' : '');
                pill.innerHTML = '<span>' + (it.title || ('作业 ' + (idx + 1))) + '</span>';
                pill.onclick = () => switchActiveHomework(idx);
                qList.appendChild(pill);
            });

            if (data.activeId > 0) {
                if (data.activeId !== localActivePhotoId) {
                    localActivePhotoId = data.activeId;
                    thumb.src = data.photoUrl;
                }
                thumb.style.transform = 'rotate(' + (data.rotation || 0) + 'deg)';
                thumb.style.display = 'block';
                placeholder.style.display = 'none';
            }
        }

        // Instant Optimistic UI Switch (0ms UI feedback + Instant Broadcast)
        async function switchActiveHomework(idx) {
            vibrate();
            currentActiveIdx = idx;
            const pills = document.querySelectorAll('.queue-item-pill');
            pills.forEach((p, i) => {
                if (i === idx) p.classList.add('active');
                else p.classList.remove('active');
            });
            try {
                await fetch('/api/select_cast?index=' + idx);
            } catch (e) {}
        }

        async function deleteCurrentPhoto() {
            if (currentActiveIdx < 0) return;
            if (!confirm('确定删除当前这份作业照片吗？')) return;
            vibrate();
            try {
                await fetch('/api/delete_cast?index=' + currentActiveIdx);
            } catch (e) {}
        }

        function handleNativePhoto(input) {
            if (input.files && input.files[0]) {
                const file = input.files[0];
                const reader = new FileReader();
                reader.onload = async function(e) {
                    const rawBase64 = e.target.result;
                    vibrate();

                    // Optimistically show preview immediately
                    const thumb = document.getElementById('previewThumb');
                    const placeholder = document.getElementById('previewPlaceholder');
                    thumb.src = rawBase64;
                    thumb.style.display = 'block';
                    placeholder.style.display = 'none';

                    try {
                        const res = await fetch('/api/cast_photo', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ image: rawBase64 })
                        });
                        if (res.ok) {
                            input.value = '';
                        }
                    } catch (err) {
                        alert('上传失败: ' + err.message);
                    }
                };
                reader.readAsDataURL(file);
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
            try {
                await fetch('/api/stop_cast', { method: 'POST' });
            } catch (e) {}
        }

        pollQueueData();
    </script>
</body>
</html>";
        }
    }
}
