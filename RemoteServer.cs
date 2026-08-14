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
                string localIp = GetLocalIPAddress();
                if (string.IsNullOrEmpty(localIp))
                {
                    localIp = "127.0.0.1";
                }
                
                // Bind to standard TCP socket which does not require admin privileges on Windows
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
                    client.ReceiveTimeout = 3000;
                    client.SendTimeout = 3000;
                    
                    using (var stream = client.GetStream())
                    using (var reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        // Read request line (e.g. "GET /next HTTP/1.1")
                        string requestLine = reader.ReadLine();
                        if (string.IsNullOrEmpty(requestLine)) return;

                        // Parse method and path
                        string[] parts = requestLine.Split(' ');
                        if (parts.Length < 2) return;
                        
                        string method = parts[0].ToUpper();
                        string path = parts[1].ToLower();
                        int queryIdx = path.IndexOf('?');
                        if (queryIdx >= 0)
                        {
                            path = path.Substring(0, queryIdx);
                        }

                        // Consume headers
                        string headerLine;
                        while (!string.IsNullOrEmpty(headerLine = reader.ReadLine())) { }

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
                            
                            string responseBody = "{\"status\":\"success\"}";
                            byte[] bodyBytes = Encoding.UTF8.GetBytes(responseBody);
                            
                            string headers = "HTTP/1.1 200 OK\r\n" +
                                             "Content-Type: application/json\r\n" +
                                             "Access-Control-Allow-Origin: *\r\n" +
                                             "Content-Length: " + bodyBytes.Length + "\r\n" +
                                             "Connection: close\r\n\r\n";
                            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                            
                            stream.Write(headerBytes, 0, headerBytes.Length);
                            stream.Write(bodyBytes, 0, bodyBytes.Length);
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
                            
                            string responseBody = "{\"status\":\"success\"}";
                            byte[] bodyBytes = Encoding.UTF8.GetBytes(responseBody);
                            
                            string headers = "HTTP/1.1 200 OK\r\n" +
                                             "Content-Type: application/json\r\n" +
                                             "Access-Control-Allow-Origin: *\r\n" +
                                             "Content-Length: " + bodyBytes.Length + "\r\n" +
                                             "Connection: close\r\n\r\n";
                            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                            
                            stream.Write(headerBytes, 0, headerBytes.Length);
                            stream.Write(bodyBytes, 0, bodyBytes.Length);
                        }
                        else if (path == "/share-vpn")
                        {
                            string query = parts[1];
                            string portStr = "";
                            string enableStr = "";
                            
                            int portIdx = query.IndexOf("port=");
                            if (portIdx >= 0)
                            {
                                int endIdx = query.IndexOf('&', portIdx);
                                if (endIdx >= 0)
                                    portStr = query.Substring(portIdx + 5, endIdx - (portIdx + 5));
                                else
                                    portStr = query.Substring(portIdx + 5);
                            }
                            
                            int enableIdx = query.IndexOf("enable=");
                            if (enableIdx >= 0)
                            {
                                int endIdx = query.IndexOf('&', enableIdx);
                                if (endIdx >= 0)
                                    enableStr = query.Substring(enableIdx + 7, endIdx - (enableIdx + 7));
                                else
                                    enableStr = query.Substring(enableIdx + 7);
                            }

                            bool enable = enableStr == "true";
                            int port = 7890;
                            int.TryParse(portStr, out port);

                            string clientIp = "";
                            try
                            {
                                var remoteEndPoint = client.Client.RemoteEndPoint as IPEndPoint;
                                if (remoteEndPoint != null)
                                {
                                    clientIp = remoteEndPoint.Address.ToString();
                                    // Strip IPv6-mapped IPv4 prefix if present (e.g. ::ffff:192.168.1.100)
                                    if (clientIp.StartsWith("::ffff:"))
                                    {
                                        clientIp = clientIp.Substring(7);
                                    }
                                }
                            }
                            catch {}

                            if (enable && !string.IsNullOrEmpty(clientIp))
                            {
                                string proxyAddress = string.Format("{0}:{1}", clientIp, port);
                                SetWindowsSystemProxy(proxyAddress, true);
                            }
                            else
                            {
                                SetWindowsSystemProxy("", false);
                            }

                            string responseBody = "{\"status\":\"success\"}";
                            byte[] bodyBytes = Encoding.UTF8.GetBytes(responseBody);
                            
                            string headers = "HTTP/1.1 200 OK\r\n" +
                                             "Content-Type: application/json\r\n" +
                                             "Access-Control-Allow-Origin: *\r\n" +
                                             "Content-Length: " + bodyBytes.Length + "\r\n" +
                                             "Connection: close\r\n\r\n";
                            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                            stream.Write(headerBytes, 0, headerBytes.Length);
                            stream.Write(bodyBytes, 0, bodyBytes.Length);
                        }
                        else if (path == "/console" || path == "/console.html" || path == "/console/")
                        {
                            string html = GetControlPageHtml();
                            byte[] bodyBytes = Encoding.UTF8.GetBytes(html);
                            
                            string headers = "HTTP/1.1 200 OK\r\n" +
                                             "Content-Type: text/html; charset=utf-8\r\n" +
                                             "Content-Length: " + bodyBytes.Length + "\r\n" +
                                             "Connection: close\r\n\r\n";
                            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                            
                            stream.Write(headerBytes, 0, headerBytes.Length);
                            stream.Write(bodyBytes, 0, bodyBytes.Length);
                        }
                        else if (!string.IsNullOrEmpty(CurrentProxyTarget) && path != "/" && path != "/index.html")
                        {
                            // Catch-all transparent reverse proxy for all assets, scripts, APIs, etc.
                            ProxyRequest(stream, method, CurrentProxyTarget, parts[1]);
                        }
                        else if (path == "/" || path == "/index.html")
                        {
                            if (!string.IsNullOrEmpty(CurrentProxyTarget))
                            {
                                // If target is set, proxy the root requests directly to the local site
                                ProxyRequest(stream, method, CurrentProxyTarget, parts[1]);
                            }
                            else
                            {
                                // If no target is set, redirect root requests to the remote control console
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
                            string headers = "HTTP/1.1 404 Not Found\r\n" +
                                             "Content-Length: 0\r\n" +
                                             "Connection: close\r\n\r\n";
                            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                            stream.Write(headerBytes, 0, headerBytes.Length);
                        }

                        // Gracefully flush and shutdown output to prevent RST packets
                        stream.Flush();
                        try
                        {
                            client.Client.Shutdown(SocketShutdown.Send);
                            Thread.Sleep(100);
                        }
                        catch { }
                    }
                }
                catch { }
            }
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
                        
                        // Automatically resume presentation state to running if it was paused,
                        // ensuring that slide transition logic continues uninterrupted.
                        if (view.State == PowerPoint.PpSlideShowState.ppSlideShowPaused)
                        {
                            view.State = PowerPoint.PpSlideShowState.ppSlideShowRunning;
                            Globals.ThisAddIn.LogToFile("ControlPowerPoint: Autoresumed paused slideshow state before transition.");
                        }
                        
                        if (next)
                            view.Next();
                        else
                            view.Previous();
                    }
                    else
                    {
                        var activeWindow = app.ActiveWindow;
                        if (activeWindow != null && activeWindow.View != null)
                        {
                            var view = activeWindow.View;
                            int currentSlideIndex = ((Microsoft.Office.Interop.PowerPoint.Slide)view.Slide).SlideIndex;
                            int slidesCount = app.ActivePresentation.Slides.Count;
                            
                            if (next && currentSlideIndex < slidesCount)
                            {
                                view.GotoSlide(currentSlideIndex + 1);
                            }
                            else if (!next && currentSlideIndex > 1)
                            {
                                view.GotoSlide(currentSlideIndex - 1);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Control PPT COM failed: " + ex.Message);
                }
            });
        }

        private class MediaShapeInfo
        {
            public PowerPoint.Shape Shape;
            public float AbsoluteLeft;
            public float AbsoluteTop;
            public bool IsVideo;
        }

        private List<MediaShapeInfo> FindAllMediaShapes(PowerPoint.Slide slide)
        {
            var result = new List<MediaShapeInfo>();
            if (slide != null && slide.Shapes != null)
            {
                FindMediaShapesRecursive(slide.Shapes, result, 0, 0);
            }
            return result;
        }

        private void FindMediaShapesRecursive(PowerPoint.Shapes shapes, List<MediaShapeInfo> result, float parentLeft, float parentTop)
        {
            if (shapes == null) return;
            foreach (PowerPoint.Shape shape in shapes)
            {
                ProcessSingleShape(shape, result, parentLeft, parentTop);
            }
        }

        private void FindMediaShapesRecursive(PowerPoint.GroupShapes shapes, List<MediaShapeInfo> result, float parentLeft, float parentTop)
        {
            if (shapes == null) return;
            foreach (PowerPoint.Shape shape in shapes)
            {
                ProcessSingleShape(shape, result, parentLeft, parentTop);
            }
        }

        private void ProcessSingleShape(PowerPoint.Shape shape, List<MediaShapeInfo> result, float parentLeft, float parentTop)
        {
            try
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                {
                    try
                    {
                        // Accumulate group Left/Top offsets recursively. GroupItems are coordinates local to their parent group bounds.
                        FindMediaShapesRecursive(shape.GroupItems, result, parentLeft + shape.Left, parentTop + shape.Top);
                    }
                    catch { }
                }
                else
                {
                    bool isMedia = false;
                    bool isVideo = false;
                    try
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia)
                        {
                            isMedia = true;
                        }
                    }
                    catch { }

                    try
                    {
                        if (shape.MediaType == PowerPoint.PpMediaType.ppMediaTypeMovie)
                        {
                            isMedia = true;
                            isVideo = true;
                        }
                        else if (shape.MediaType == PowerPoint.PpMediaType.ppMediaTypeSound)
                        {
                            isMedia = true;
                        }
                    }
                    catch { }

                    if (isMedia)
                    {
                        result.Add(new MediaShapeInfo
                        {
                            Shape = shape,
                            AbsoluteLeft = parentLeft + shape.Left,
                            AbsoluteTop = parentTop + shape.Top,
                            IsVideo = isVideo
                        });
                    }
                }
            }
            catch { }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ScreenToClient(IntPtr hWnd, ref POINT lpPoint);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr ChildWindowFromPointEx(IntPtr hwndParent, POINT pt, uint uFlags);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit)]
        private struct INPUT_UNION
        {
            [System.Runtime.InteropServices.FieldOffset(0)]
            public MOUSEINPUT mi;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;
            public INPUT_UNION ui;
        }

        private const uint INPUT_MOUSE = 0;
        private const uint MOUSEEVENTF_MOVE = 0x0001;
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_ABSOLUTE = 0x8000;
        private const uint MOUSEEVENTF_VIRTUALDESK = 0x4000;
        private const uint WM_LBUTTONDOWN = 0x0201;
        private const uint WM_LBUTTONUP = 0x0202;
        private const uint MK_LBUTTON = 0x0001;
        private const uint CWP_SKIPINVISIBLE = 0x0001;

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        // Make a POINT compatible lParam for WM_LBUTTON* messages (low 16 bits = x, high 16 bits = y, both in client coords)
        private static IntPtr MakeLParam(int x, int y)
        {
            return (IntPtr)((y << 16) | (x & 0xFFFF));
        }

        // P/Invoke for keyboard input
        private const uint INPUT_KEYBOARD     = 1;
        private const uint KEYEVENTF_KEYUP    = 0x0002;
        private const uint KEYEVENTF_EXTENDEDKEY = 0x0001;

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint   dwFlags;
            public uint   time;
            public IntPtr dwExtraInfo;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit)]
        private struct INPUT_UNION2
        {
            [System.Runtime.InteropServices.FieldOffset(0)]
            public MOUSEINPUT  mi;
            [System.Runtime.InteropServices.FieldOffset(0)]
            public KEYBDINPUT  ki;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct INPUT2
        {
            public uint      type;
            public INPUT_UNION2 ui;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT2[] pInputs, int cbSize);

        [System.Runtime.InteropServices.DllImport("wininet.dll", SetLastError = true)]
        private static extern bool InternetSetOption(IntPtr hInternet, int dwOption, IntPtr lpBuffer, int dwBufferLength);

        private const int INTERNET_OPTION_SETTINGS_CHANGED = 39;
        private const int INTERNET_OPTION_REFRESH = 37;

        public static void SetWindowsSystemProxy(string proxyAddress, bool enable)
        {
            try
            {
                using (Microsoft.Win32.RegistryKey registry = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings", true))
                {
                    if (registry != null)
                    {
                        registry.SetValue("ProxyEnable", enable ? 1 : 0);
                        if (enable && !string.IsNullOrEmpty(proxyAddress))
                        {
                            registry.SetValue("ProxyServer", proxyAddress);
                        }
                    }
                }

                // Broadcast change
                InternetSetOption(IntPtr.Zero, INTERNET_OPTION_SETTINGS_CHANGED, IntPtr.Zero, 0);
                InternetSetOption(IntPtr.Zero, INTERNET_OPTION_REFRESH, IntPtr.Zero, 0);
                
                Globals.ThisAddIn.LogToFile(string.Format("SetWindowsSystemProxy: proxyAddress={0}, enable={1}", proxyAddress, enable));
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.LogToFile("SetWindowsSystemProxy failed: " + ex.Message);
            }
        }

        private void SendAltP(IntPtr slideshowHwnd)
        {
            // Alt+P is PowerPoint's OFFICIAL built-in media play/pause TOGGLE shortcut.
            // Unlike mouse clicks, it is media-context-aware and NEVER advances the slide.
            // Source: PowerPoint shortcut documentation + OSCPoint / presenter remote projects.
            //
            // We must bring the slideshow window to foreground first so it receives the key.
            SetForegroundWindow(slideshowHwnd);
            System.Threading.Thread.Sleep(60); // wait for focus switch

            const ushort VK_ALT = 0x12;
            const ushort VK_P   = 0x50;

            INPUT2[] inputs = new INPUT2[4];

            // Alt DOWN
            inputs[0].type    = INPUT_KEYBOARD;
            inputs[0].ui.ki.wVk      = VK_ALT;
            inputs[0].ui.ki.dwFlags  = 0;

            // P DOWN
            inputs[1].type    = INPUT_KEYBOARD;
            inputs[1].ui.ki.wVk      = VK_P;
            inputs[1].ui.ki.dwFlags  = 0;

            // P UP
            inputs[2].type    = INPUT_KEYBOARD;
            inputs[2].ui.ki.wVk      = VK_P;
            inputs[2].ui.ki.dwFlags  = KEYEVENTF_KEYUP;

            // Alt UP
            inputs[3].type    = INPUT_KEYBOARD;
            inputs[3].ui.ki.wVk      = VK_ALT;
            inputs[3].ui.ki.dwFlags  = KEYEVENTF_KEYUP;

            uint result = SendInput(4, inputs, System.Runtime.InteropServices.Marshal.SizeOf(typeof(INPUT2)));
            Globals.ThisAddIn.LogToFile(string.Format("SendAltP: hwnd={0} result={1}", slideshowHwnd, result));
        }

        private bool PostMessageClickAt(IntPtr rootHwnd, int screenX, int screenY)
        {
            // Strategy: Find the deepest child window at the target screen coordinates.
            // Then post WM_LBUTTONDOWN/UP directly to that child window using client coordinates.
            // This bypasses the global hit-test system entirely.
            // NOTE: AdvanceOnClick is already disabled on slides with video, so even if we
            // post to rootHwnd it will NOT trigger a slide advance.
            POINT screenPt = new POINT { X = screenX, Y = screenY };

            // Find the topmost child window at this screen position
            IntPtr targetHwnd = WindowFromPoint(screenPt);
            if (targetHwnd == IntPtr.Zero)
                targetHwnd = rootHwnd;

            Globals.ThisAddIn.LogToFile(string.Format("PostMessageClickAt: rootHwnd={0}, targetChildHwnd={1}", rootHwnd, targetHwnd));

            // Convert screen coords to client coords of the target window
            POINT clientPt = new POINT { X = screenX, Y = screenY };
            ScreenToClient(targetHwnd, ref clientPt);

            IntPtr lParam = MakeLParam(clientPt.X, clientPt.Y);
            IntPtr wParam = (IntPtr)MK_LBUTTON;

            // Post the click directly into the window's message queue
            PostMessage(targetHwnd, WM_LBUTTONDOWN, wParam, lParam);
            System.Threading.Thread.Sleep(40);
            PostMessage(targetHwnd, WM_LBUTTONUP, IntPtr.Zero, lParam);

            Globals.ThisAddIn.LogToFile(string.Format("PostMessageClickAt: Posted to hwnd={0} clientPt=({1},{2})", targetHwnd, clientPt.X, clientPt.Y));
            return true;
        }


        private void SendInputClickAt(int screenX, int screenY)
        {
            int virtualLeft   = System.Windows.Forms.SystemInformation.VirtualScreen.Left;
            int virtualTop    = System.Windows.Forms.SystemInformation.VirtualScreen.Top;
            int virtualWidth  = System.Windows.Forms.SystemInformation.VirtualScreen.Width;
            int virtualHeight = System.Windows.Forms.SystemInformation.VirtualScreen.Height;

            int normX = ((screenX - virtualLeft) * 65536) / virtualWidth;
            int normY = ((screenY - virtualTop)  * 65536) / virtualHeight;

            INPUT2[] inputs = new INPUT2[3];

            // 1. MOVE cursor to exact video center (sets physical cursor position for PPT hit-test)
            inputs[0].type = INPUT_MOUSE;
            inputs[0].ui.mi.dx      = normX;
            inputs[0].ui.mi.dy      = normY;
            inputs[0].ui.mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK;

            // 2. LEFT BUTTON DOWN
            inputs[1].type = INPUT_MOUSE;
            inputs[1].ui.mi.dx      = normX;
            inputs[1].ui.mi.dy      = normY;
            inputs[1].ui.mi.dwFlags = MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK;

            // 3. LEFT BUTTON UP
            inputs[2].type = INPUT_MOUSE;
            inputs[2].ui.mi.dx      = normX;
            inputs[2].ui.mi.dy      = normY;
            inputs[2].ui.mi.dwFlags = MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK;

            uint result = SendInput(3, inputs, System.Runtime.InteropServices.Marshal.SizeOf(typeof(INPUT2)));
            Globals.ThisAddIn.LogToFile(string.Format("SendInputClickAt: screen=({0},{1}) norm=({2},{3}) result={4}", screenX, screenY, normX, normY, result));
        }

        private void ControlPowerPointVideo()
        {
            Globals.ThisAddIn.SafeInvoke(delegate()
            {
                Globals.ThisAddIn.LogToFile("ControlVideo: Video Play/Pause control has been disabled by user request.");
            });
        }

        private void ControlPowerPointVolume(float vol)
        {
            try
            {
                VolumeHelper.SetSystemVolume(vol);
                Globals.ThisAddIn.LogToFile("ControlPowerPointVolume: successfully set system volume to " + vol);
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.LogToFile("Control PPT Volume failed: " + ex.Message);
            }
        }

        private static void ProxyRequest(Stream clientStream, string method, string targetBase, string subPathAndQuery)
        {
            try
            {
                string targetUrl = targetBase.TrimEnd('/') + "/" + subPathAndQuery.TrimStart('/');
                var req = (HttpWebRequest)WebRequest.Create(targetUrl);
                req.Method = method;
                req.Timeout = 8000;
                req.AllowAutoRedirect = false; // Let client handle redirects

                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    string contentType = resp.ContentType ?? "application/octet-stream";
                    using (var respStream = resp.GetResponseStream())
                    {
                        byte[] buffer = new byte[8192];
                        int bytesRead;
                        using (var ms = new MemoryStream())
                        {
                            while ((bytesRead = respStream.Read(buffer, 0, buffer.Length)) > 0)
                            {
                                ms.Write(buffer, 0, bytesRead);
                            }
                            byte[] body = ms.ToArray();

                            string headers = string.Format(
                                "HTTP/1.1 {0} {1}\r\n" +
                                "Content-Type: {2}\r\n" +
                                "Access-Control-Allow-Origin: *\r\n" +
                                "Content-Length: {3}\r\n" +
                                "Connection: close\r\n\r\n",
                                (int)resp.StatusCode,
                                resp.StatusDescription,
                                contentType,
                                body.Length
                            );
                            byte[] headerBytes = Encoding.UTF8.GetBytes(headers);
                            clientStream.Write(headerBytes, 0, headerBytes.Length);
                            clientStream.Write(body, 0, body.Length);
                        }
                    }
                }
            }
            catch (WebException webEx)
            {
                try
                {
                    var resp = webEx.Response as HttpWebResponse;
                    if (resp != null)
                    {
                        string contentType = resp.ContentType ?? "text/html";
                        using (var respStream = resp.GetResponseStream())
                        {
                            byte[] buffer = new byte[8192];
                            int bytesRead;
                            using (var ms = new MemoryStream())
                            {
                                while ((bytesRead = respStream.Read(buffer, 0, buffer.Length)) > 0)
                                {
                                    ms.Write(buffer, 0, bytesRead);
                                }
                                byte[] body = ms.ToArray();
                                string headers = string.Format(
                                    "HTTP/1.1 {0} {1}\r\n" +
                                    "Content-Type: {2}\r\n" +
                                    "Content-Length: {3}\r\n" +
                                    "Connection: close\r\n\r\n",
                                    (int)resp.StatusCode,
                                    resp.StatusDescription,
                                    contentType,
                                    body.Length
                                );
                                clientStream.Write(Encoding.UTF8.GetBytes(headers), 0, headers.Length);
                                clientStream.Write(body, 0, body.Length);
                            }
                        }
                    }
                    else
                    {
                        string errBody = "<html><body><h2>Proxy Connection Error</h2><p>" + webEx.Message + "</p></body></html>";
                        byte[] errBytes = Encoding.UTF8.GetBytes(errBody);
                        string errHeaders = string.Format(
                            "HTTP/1.1 502 Bad Gateway\r\n" +
                            "Content-Type: text/html; charset=utf-8\r\n" +
                            "Content-Length: {0}\r\n" +
                            "Connection: close\r\n\r\n",
                            errBytes.Length
                        );
                        clientStream.Write(Encoding.UTF8.GetBytes(errHeaders), 0, errHeaders.Length);
                        clientStream.Write(errBytes, 0, errBytes.Length);
                    }
                }
                catch {}
            }
            catch (Exception ex)
            {
                try
                {
                    string errBody = "<html><body><h2>Proxy Server Error</h2><p>" + ex.Message + "</p></body></html>";
                    byte[] errBytes = Encoding.UTF8.GetBytes(errBody);
                    string errHeaders = string.Format(
                        "HTTP/1.1 500 Internal Server Error\r\n" +
                        "Content-Type: text/html; charset=utf-8\r\n" +
                        "Content-Length: {0}\r\n" +
                        "Connection: close\r\n\r\n",
                        errBytes.Length
                    );
                    clientStream.Write(Encoding.UTF8.GetBytes(errHeaders), 0, errHeaders.Length);
                    clientStream.Write(errBytes, 0, errBytes.Length);
                }
                catch {}
            }
        }

        private string GetLocalIPAddress()
        {
            try
            {
                // Prioritize active physical wireless/ethernet adapters with gateways
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (ni.OperationalStatus != OperationalStatus.Up)
                        continue;
                        
                    if (ni.NetworkInterfaceType != NetworkInterfaceType.Wireless80211 && 
                        ni.NetworkInterfaceType != NetworkInterfaceType.Ethernet)
                        continue;

                    string desc = ni.Description.ToLower();
                    if (desc.Contains("virtual") || desc.Contains("vmware") || 
                        desc.Contains("virtualbox") || desc.Contains("wsl") || 
                        desc.Contains("hyper-v") || desc.Contains("vbox") || 
                        desc.Contains("vpn") || desc.Contains("loopback"))
                        continue;

                    IPInterfaceProperties ipProps = ni.GetIPProperties();
                    if (ipProps.GatewayAddresses.Count == 0)
                        continue;

                    foreach (UnicastIPAddressInformation ip in ipProps.UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            return ip.Address.ToString();
                        }
                    }
                }

                // Fallback 1: Any active ethernet/wireless IP
                foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (ni.OperationalStatus != OperationalStatus.Up)
                        continue;

                    if (ni.NetworkInterfaceType != NetworkInterfaceType.Wireless80211 && 
                        ni.NetworkInterfaceType != NetworkInterfaceType.Ethernet)
                        continue;

                    foreach (UnicastIPAddressInformation ip in ni.GetIPProperties().UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == AddressFamily.InterNetwork)
                        {
                            return ip.Address.ToString();
                        }
                    }
                }
            }
            catch { }

            // Fallback 2: Old DNS method
            try
            {
                var host = Dns.GetHostEntry(Dns.GetHostName());
                foreach (var ip in host.AddressList)
                {
                    if (ip.AddressFamily == AddressFamily.InterNetwork)
                    {
                        if (!IPAddress.IsLoopback(ip))
                        {
                            return ip.ToString();
                        }
                    }
                }
            }
            catch { }

            return null;
        }

        private string GetControlPageHtml()
        {
            return @"<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
    <meta charset=""utf-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, viewport-fit=cover"">
    <title>PPT 智能遥控器</title>
    <style>
        :root {
            --bg-color: #f0f2f5;
            --remote-bg: linear-gradient(145deg, #fbfbfd, #e6e9ee);
            --text-primary: #1d1d1f;
            --text-secondary: #86868b;
            --accent-color: #30d158;
            --led-shadow: 0 0 8px #30d158;
        }

        * {
            box-sizing: border-box;
            -webkit-tap-highlight-color: transparent;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, ""SF Pro Text"", ""Segoe UI"", Roboto, Helvetica, Arial, sans-serif;
            background-color: var(--bg-color);
            color: var(--text-primary);
            margin: 0;
            padding: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            min-height: 100dvh;
            overflow: hidden;
            background-image: radial-gradient(circle at 50% -20%, #ffffff, transparent 70%);
        }

        .remote-body {
            width: 85%;
            max-width: 290px;
            padding: 40px 24px 30px 24px;
            background: var(--remote-bg);
            border-radius: 40px;
            border: 1px solid rgba(255, 255, 255, 0.7);
            box-shadow: 
                15px 15px 30px rgba(163, 177, 198, 0.5), 
                -15px -15px 30px rgba(255, 255, 255, 0.9),
                inset 1px 1px 2px rgba(255, 255, 255, 0.9),
                inset -1px -1px 2px rgba(163, 177, 198, 0.2);
            display: flex;
            flex-direction: column;
            align-items: center;
            position: relative;
        }

        .remote-body::before {
            content: '';
            position: absolute;
            top: 2px; left: 2px; right: 2px; bottom: 2px;
            border-radius: 38px;
            border: 1px solid rgba(255, 255, 255, 0.4);
            pointer-events: none;
        }

        .led-indicator {
            width: 14px;
            height: 14px;
            background-color: #222;
            border-radius: 50%;
            margin-bottom: 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 
                inset 1px 1px 2px rgba(0, 0, 0, 0.6), 
                1px 1px 1px rgba(255, 255, 255, 0.8);
        }

        .led-light {
            width: 6px;
            height: 6px;
            background-color: var(--led-color, #30d158);
            border-radius: 50%;
            position: relative;
            box-shadow: 0 0 6px var(--led-color, #30d158);
            transition: all 0.2s ease;
            --led-color: #30d158;
        }

        .led-light::before,
        .led-light::after {
            content: '';
            position: absolute;
            left: 50%; top: 50%;
            width: 8px;
            height: 8px;
            margin-left: -4px;
            margin-top: -4px;
            border-radius: 50%;
            background-color: transparent;
            border: 1.5px solid var(--led-color);
            opacity: 0;
            pointer-events: none;
            box-sizing: border-box;
        }

        .led-light::before {
            animation: ledRipple 2.5s infinite cubic-bezier(0.25, 0, 0, 1);
        }

        .led-light::after {
            animation: ledRipple 2.5s infinite cubic-bezier(0.25, 0, 0, 1);
            animation-delay: 1.25s;
        }

        .led-light.loading {
            --led-color: #ff9500;
        }

        .led-light.loading::before {
            animation: ledRipple 1.0s infinite cubic-bezier(0.25, 0, 0, 1);
        }

        .led-light.loading::after {
            animation: ledRipple 1.0s infinite cubic-bezier(0.25, 0, 0, 1);
            animation-delay: 0.5s;
        }

        .led-light.error {
            --led-color: #ff3b30;
        }

        .led-light.error::before {
            animation: ledRipple 1.8s infinite cubic-bezier(0.25, 0, 0, 1);
        }

        .led-light.error::after {
            animation: ledRipple 1.8s infinite cubic-bezier(0.25, 0, 0, 1);
            animation-delay: 0.9s;
        }

        @keyframes ledRipple {
            0% {
                transform: scale(0.8);
                opacity: 0.8;
                box-shadow: inset 0 0 2px var(--led-color), 0 0 2px var(--led-color);
            }
            50% {
                opacity: 0.4;
            }
            100% {
                transform: scale(3.5);
                opacity: 0;
                box-shadow: inset 0 0 4px var(--led-color), 0 0 8px var(--led-color);
            }
        }

        .brand-text {
            font-size: 14px;
            font-weight: 700;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            color: var(--text-secondary);
            margin-bottom: 35px;
            text-shadow: 1px 1px 1px rgba(255, 255, 255, 0.85);
        }

        .btn-group {
            width: 100%;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 25px;
        }

        .remote-btn {
            background: linear-gradient(135deg, #f5f6f9, #dce1e9);
            border: 1px solid rgba(255, 255, 255, 0.6);
            border-radius: 24px;
            color: #4a5568;
            cursor: pointer;
            outline: none;
            user-select: none;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.15s ease;
        }

        .row-buttons {
            display: flex;
            width: 100%;
            gap: 16px;
        }

        .main-row {
            justify-content: center;
            gap: 24px;
        }

        .btn-round {
            width: 100px;
            height: 100px;
            border-radius: 50%;
            box-shadow: 
                6px 6px 15px rgba(163, 177, 198, 0.7), 
                -6px -6px 15px rgba(255, 255, 255, 0.9),
                inset 1px 1px 2px rgba(255, 255, 255, 0.9);
            flex-direction: column;
            gap: 6px;
            font-size: 13px;
            font-weight: 600;
        }

        .btn-round:active {
            background: linear-gradient(135deg, #dce1e9, #f5f6f9);
            box-shadow: 
                inset 4px 4px 8px rgba(163, 177, 198, 0.7), 
                inset -4px -4px 8px rgba(255, 255, 255, 0.9);
            transform: scale(0.97);
        }

        .remote-btn svg {
            width: 30px;
            height: 30px;
            fill: #4a5568;
            filter: drop-shadow(1px 1px 0px rgba(255, 255, 255, 0.8));
        }

        /* Volume Slider Styling */
        .volume-container {
            width: 224px;
            margin-top: 25px;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 8px;
        }

        .volume-label {
            font-size: 11px;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 0.5px;
            display: flex;
            justify-content: space-between;
            width: 100%;
        }

        .volume-slider-wrapper {
            position: relative;
            width: 100%;
            height: 24px;
            display: flex;
            align-items: center;
        }

        .volume-slider {
            -webkit-appearance: none;
            width: 100%;
            height: 12px;
            border-radius: 6px;
            background: #e6e9ee;
            outline: none;
            box-shadow: 
                inset 2px 2px 5px rgba(163, 177, 198, 0.5), 
                inset -2px -2px 5px rgba(255, 255, 255, 0.8);
            margin: 0;
            position: relative;
            z-index: 1;
        }

        .volume-slider::-webkit-slider-runnable-track {
            width: 100%;
            height: 12px;
            border-radius: 6px;
            background: transparent;
        }

        .volume-slider::-webkit-slider-thumb {
            -webkit-appearance: none;
            appearance: none;
            width: 26px;
            height: 26px;
            border-radius: 50%;
            background: radial-gradient(circle at 30% 30%, #ffffff 0%, #dcdfe6 100%);
            border: 1px solid rgba(0, 0, 0, 0.08);
            box-shadow: 
                2px 4px 6px rgba(163, 177, 198, 0.6), 
                -1px -1px 3px rgba(255, 255, 255, 0.9);
            cursor: pointer;
            transition: transform 0.1s ease;
            margin-top: -7px;
            position: relative;
            z-index: 2;
        }

        .volume-slider::-webkit-slider-thumb:active {
            transform: scale(1.1);
        }

        /* VPN Share Styling */
        .vpn-container {
            width: 224px;
            margin-top: 30px;
            display: flex;
            flex-direction: column;
            gap: 15px;
            padding: 16px;
            border-radius: 20px;
            background: var(--remote-bg);
            box-shadow: 
                3px 3px 6px rgba(163, 177, 198, 0.4), 
                -3px -3px 6px rgba(255, 255, 255, 0.8);
        }

        .vpn-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            width: 100%;
        }

        .vpn-title {
            font-size: 13px;
            font-weight: 600;
            color: var(--text-primary);
            letter-spacing: 0.5px;
            cursor: pointer;
            transition: opacity 0.2s;
            user-select: none;
        }

        .vpn-title:active {
            opacity: 0.6;
        }

        /* Neumorphic Toggle Switch */
        .switch {
            position: relative;
            display: inline-block;
            width: 60px;
            height: 32px;
        }

        .switch input {
            opacity: 0;
            width: 0;
            height: 0;
        }

        .slider {
            position: absolute;
            cursor: pointer;
            top: 0; left: 0; right: 0; bottom: 0;
            background-color: #e6e9ee;
            transition: .3s;
            border-radius: 32px;
            box-shadow: 
                inset 2px 2px 5px rgba(163, 177, 198, 0.5), 
                inset -2px -2px 5px rgba(255, 255, 255, 0.8);
        }

        .slider:before {
            position: absolute;
            content: """";
            height: 24px;
            width: 24px;
            left: 4px;
            bottom: 4px;
            background: radial-gradient(circle at 30% 30%, #ffffff 0%, #dcdfe6 100%);
            transition: .3s;
            border-radius: 50%;
            box-shadow: 
                2px 2px 5px rgba(163, 177, 198, 0.4), 
                -1px -1px 2px rgba(255, 255, 255, 0.9);
        }

        input:checked + .slider {
            background-color: #30d158;
            box-shadow: 
                inset 2px 2px 5px rgba(0, 0, 0, 0.2), 
                inset -2px -2px 5px rgba(255, 255, 255, 0.1);
        }

        input:checked + .slider:before {
            transform: translateX(28px);
            background: #ffffff;
        }

        /* Modal Overlay & Content */
        .modal-overlay {
            display: none;
            position: fixed;
            top: 0; left: 0; width: 100%; height: 100%;
            background: rgba(0, 0, 0, 0.4);
            backdrop-filter: blur(4px);
            z-index: 1000;
            justify-content: center;
            align-items: center;
        }

        .modal-content {
            width: 80%;
            max-width: 260px;
            background: var(--remote-bg);
            border-radius: 28px;
            padding: 24px;
            border: 1px solid rgba(255, 255, 255, 0.7);
            box-shadow: 
                15px 15px 30px rgba(0, 0, 0, 0.15), 
                -10px -10px 20px rgba(255, 255, 255, 0.9);
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 20px;
            animation: modalFadeIn 0.2s ease-out;
        }

        @keyframes modalFadeIn {
            from { transform: scale(0.9); opacity: 0; }
            to { transform: scale(1); opacity: 1; }
        }

        .modal-title {
            font-size: 14px;
            font-weight: 700;
            color: var(--text-primary);
            text-align: center;
            letter-spacing: 0.5px;
        }

        .vpn-input {
            width: 120px;
            height: 40px;
            border-radius: 12px;
            border: none;
            background: #e6e9ee;
            outline: none;
            box-shadow: 
                inset 2px 2px 5px rgba(163, 177, 198, 0.5), 
                inset -2px -2px 5px rgba(255, 255, 255, 0.8);
            text-align: center;
            font-size: 16px;
            font-weight: 600;
            color: #4a5568;
        }

        .btn-confirm {
            width: 100%;
            height: 40px;
            border-radius: 12px;
            font-weight: 600;
            font-size: 13px;
            box-shadow: 
                3px 3px 6px rgba(163, 177, 198, 0.5), 
                -3px -3px 6px rgba(255, 255, 255, 0.8),
                inset 1px 1px 1px rgba(255, 255, 255, 0.9);
        }

        .btn-confirm:active {
            background: linear-gradient(135deg, #dce1e9, #f5f6f9);
            box-shadow: 
                inset 3px 3px 6px rgba(163, 177, 198, 0.6), 
                inset -3px -3px 6px rgba(255, 255, 255, 0.9);
        }

        .status-text {
            margin-top: 25px;
            font-size: 11px;
            font-weight: 600;
            color: var(--text-secondary);
            text-transform: uppercase;
            letter-spacing: 1px;
            text-shadow: 1px 1px 1px rgba(255, 255, 255, 0.8);
        }
    </style>
</head>
<body>
    <div class=""remote-body"">
        <div class=""led-indicator"">
            <div class=""led-light"" id=""led""></div>
        </div>
        <div class=""brand-text"">Presenter</div>
        
        <div class=""btn-group"">
            <div class=""row-buttons main-row"">
                <button class=""remote-btn btn-round"" onclick=""control('prev')"" title=""上一页"">
                    <svg viewBox=""0 0 24 24"" style=""transform: rotate(180deg);"">
                        <path d=""M8 5v14l11-7z""/>
                    </svg>
                    <span>上一页</span>
                </button>
                <button class=""remote-btn btn-round"" onclick=""control('next')"" title=""下一页"">
                    <svg viewBox=""0 0 24 24"">
                        <path d=""M8 5v14l11-7z""/>
                    </svg>
                    <span>下一页</span>
                </button>
            </div>
        </div>

        <div class=""volume-container"">
            <div class=""volume-label"">
                <span>音量 (Volume)</span>
                <span id=""vol-val"">50%</span>
            </div>
            <div class=""volume-slider-wrapper"">
                <input type=""range"" class=""volume-slider"" id=""volume"" min=""0"" max=""1"" step=""0.05"" value=""0.5"" oninput=""changeVolume(this)"">
            </div>
        </div>

        <div class=""vpn-container"">
            <div class=""vpn-header"">
                <span class=""vpn-title"" onclick=""openPortModal()"">网络共享</span>
                <label class=""switch"">
                    <input type=""checkbox"" id=""vpnToggle"" onchange=""toggleVpnShare(this)"">
                    <span class=""slider""></span>
                </label>
            </div>
        </div>

        <div class=""status-text"" id=""status"">Ready</div>
    </div>

    <!-- Port Settings Modal -->
    <div id=""portModal"" class=""modal-overlay"" onclick=""closePortModal()"">
        <div class=""modal-content"" onclick=""event.stopPropagation()"">
            <div class=""modal-title"">设置共享代理端口</div>
            <input type=""number"" id=""vpnPort"" class=""vpn-input"" value=""7890"">
            <button class=""remote-btn btn-confirm"" onclick=""confirmPort()"">确认</button>
        </div>
    </div>
    
    <script>
        var lastVol = -1;
        var throttleTimeout = null;
        var lastSendTime = 0;

        function updateVolumeBackground(slider) {
            var val = (slider.value - slider.min) / (slider.max - slider.min) * 100;
            slider.style.background = 'linear-gradient(to right, #002fa7 0%, #002fa7 ' + val + '%, #e6e9ee ' + val + '%, #e6e9ee 100%)';
            document.getElementById('vol-val').innerText = Math.round(val) + '%';
        }

        function changeVolume(slider) {
            updateVolumeBackground(slider);
            var vol = parseFloat(slider.value);
            if (vol === lastVol) return;
            lastVol = vol;

            var now = Date.now();
            var remaining = 150 - (now - lastSendTime);

            if (remaining <= 0) {
                if (throttleTimeout) {
                    clearTimeout(throttleTimeout);
                    throttleTimeout = null;
                }
                sendVolume(vol);
                lastSendTime = now;
            } else {
                if (throttleTimeout) {
                    clearTimeout(throttleTimeout);
                }
                throttleTimeout = setTimeout(function() {
                    sendVolume(vol);
                    lastSendTime = Date.now();
                    throttleTimeout = null;
                }, remaining);
            }
        }

        function sendVolume(vol) {
            fetch('/volume?val=' + vol)
                .catch(function(err) {
                    console.error('Volume send failed:', err);
                });
        }

        window.onload = function() {
            var slider = document.getElementById('volume');
            updateVolumeBackground(slider);
        }

        function control(action) {
            var statusLabel = document.getElementById('status');
            var led = document.getElementById('led');
            
            statusLabel.innerText = 'Sending...';
            led.className = 'led-light loading';
            
            if (navigator.vibrate) {
                navigator.vibrate([40]);
            }
            
            fetch('/' + action)
                .then(res => res.json())
                .then(data => {
                    if(data.status === 'success') {
                        statusLabel.innerText = 'Ready';
                        led.className = 'led-light';
                    } else {
                        statusLabel.innerText = 'Error';
                        led.className = 'led-light error';
                    }
                })
                .catch(err => {
                    statusLabel.innerText = 'Offline';
                    led.className = 'led-light error';
                });
        }

        function toggleVpnShare(checkbox) {
            var port = document.getElementById('vpnPort').value || '7890';
            var statusLabel = document.getElementById('status');
            var led = document.getElementById('led');
            
            statusLabel.innerText = 'Setting...';
            led.className = 'led-light loading';
            
            var enable = checkbox.checked;
            
            fetch('/share-vpn?port=' + port + '&enable=' + enable)
                .then(res => res.json())
                .then(data => {
                    if (data.status === 'success') {
                        statusLabel.innerText = 'Ready';
                        led.className = 'led-light';
                    } else {
                        checkbox.checked = !enable; // revert
                        statusLabel.innerText = 'Error';
                        led.className = 'led-light error';
                    }
                })
                .catch(err => {
                    checkbox.checked = !enable; // revert
                    statusLabel.innerText = 'Offline';
                    led.className = 'led-light error';
                });
        }

        function openPortModal() {
            document.getElementById('portModal').style.display = 'flex';
        }

        function closePortModal() {
            document.getElementById('portModal').style.display = 'none';
        }

        function confirmPort() {
            var port = document.getElementById('vpnPort').value || '7890';
            closePortModal();
            
            var led = document.getElementById('led');
            var statusLabel = document.getElementById('status');
            var oldClass = led.className;
            var oldStatus = statusLabel.innerText;
            
            statusLabel.innerText = 'Saved Port ' + port;
            led.className = 'led-light loading';
            
            setTimeout(function() {
                led.className = oldClass;
                statusLabel.innerText = oldStatus;
            }, 1000);
        }
    </script>
</body>
</html>";
        }
    }

    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    public class MMDeviceEnumerator
    {
    }

    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDeviceEnumerator
    {
        [PreserveSig]
        int EnumAudioEndpoints(int dataFlow, int dwStateMask, out IntPtr ppDevices);
        
        [PreserveSig]
        int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppEndpoint);
    }

    [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IMMDevice
    {
        [PreserveSig]
        int Activate([In, MarshalAs(UnmanagedType.LPStruct)] Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
    }

    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IAudioEndpointVolume
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
    }

    public static class VolumeHelper
    {
        public static void SetSystemVolume(float level)
        {
            IMMDevice speakers = null;
            object endpointVolumeObj = null;
            IAudioEndpointVolume endpointVolume = null;
            IMMDeviceEnumerator enumerator = null;
            try
            {
                enumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
                // 0 = eRender, 1 = eMultimedia
                int hr = enumerator.GetDefaultAudioEndpoint(0, 1, out speakers);
                if (hr == 0 && speakers != null)
                {
                    Guid iid = new Guid("5CDF2C82-841E-4546-9722-0CF74078229A");
                    hr = speakers.Activate(iid, 23, IntPtr.Zero, out endpointVolumeObj); // CLSCTX_ALL = 23
                    if (hr == 0 && endpointVolumeObj != null)
                    {
                        endpointVolume = (IAudioEndpointVolume)endpointVolumeObj;
                        Guid eventContext = Guid.Empty;
                        endpointVolume.SetMasterVolumeLevelScalar(level, ref eventContext);
                    }
                }
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.LogToFile("SetSystemVolume failed: " + ex.ToString());
            }
            finally
            {
                if (endpointVolume != null) Marshal.ReleaseComObject(endpointVolume);
                if (speakers != null) Marshal.ReleaseComObject(speakers);
                if (enumerator != null) Marshal.ReleaseComObject(enumerator);
            }
        }
    }
}
