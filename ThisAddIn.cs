using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PPTWebBrowserAddIn
{
    public partial class ThisAddIn
    {
        public static ThisAddIn Instance;
        private WebBrowserForm _liveCastForm = null;
        private PowerPoint.Slide _currentSlide = null;
        private System.Windows.Forms.Timer _slideMonitorTimer;
        private bool _isInSlideShow = false;
        private Dictionary<PowerPoint.Slide, WebBrowserForm> _activeWebViews = new Dictionary<PowerPoint.Slide, WebBrowserForm>();
        private List<PowerPoint.Shape> _hiddenShapes = new List<PowerPoint.Shape>();
        
        private RemoteServer _remoteServer;
        private SynchronizationContext _syncContext;
        private Control _marshaler;

        public DateTime LastSlideChangeTime { get; set; }
        private Microsoft.Office.Core.MsoTriState _originalAdvanceOnClick = Microsoft.Office.Core.MsoTriState.msoTrue;
        private PowerPoint.Slide _modifiedSlide = null;

        private string _tempSlidePath = null;
        public string GetTempSlidePath()
        {
            if (_tempSlidePath == null)
            {
                _tempSlidePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ppt_current_slide.png");
            }
            return _tempSlidePath;
        }

        private int _slideVersion = 0;
        public int SlideVersion { get { return _slideVersion; } }
        private int _currentSlideIndex = 0;
        public int CurrentSlideIndex { get { return _currentSlideIndex; } }
        private int _totalSlidesCount = 0;
        public int TotalSlidesCount { get { return _totalSlidesCount; } }
        public bool IsInSlideShow { get { return _isInSlideShow; } }

        public void ExportCurrentSlideImage()
        {
            // Slide sync and sharing features disabled by user request.
        }


        public void SafeInvoke(Action action)
        {
            if (_marshaler != null && _marshaler.IsHandleCreated)
            {
                if (_marshaler.InvokeRequired)
                {
                    try
                    {
                        _marshaler.Invoke(action);
                    }
                    catch { }
                }
                else
                {
                    action();
                }
            }
            else
            {
                action();
            }
        }

        public void LogToFile(string msg)
        {
            try
            {
                string tempLog = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "PowerPointAddIn_startup.log");
                System.IO.File.AppendAllText(tempLog, "\n[" + DateTime.Now.ToString("HH:mm:ss.fff") + "] " + msg);
            }
            catch { }
        }

        [DllImport("user32.dll")]
        private static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        private bool _isInitialized = false;

        public void EnsureInitialized()
        {
            if (_isInitialized) return;
            _isInitialized = true;
            Instance = this;

            try
            {
                _marshaler = new Control();
                _marshaler.CreateControl();

                _syncContext = SynchronizationContext.Current;
                LogToFile("EnsureInitialized started at " + DateTime.Now.ToString());

                SettingsManager.Load();

                Application.SlideShowBegin += Application_SlideShowBegin;
                Application.SlideShowNextSlide += Application_SlideShowNextSlide;
                Application.SlideShowEnd += Application_SlideShowEnd;

                InitializeSlideMonitor();

                if (SettingsManager.Current.RemoteControlEnabled)
                {
                    StartRemoteServer();
                }

                LogToFile("EnsureInitialized completed successfully.");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("网页嵌入插件初始化出错！\n详情: " + ex.ToString(), "初始化错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                LogToFile("EnsureInitialized Error: " + ex.ToString());
            }
        }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            EnsureInitialized();
        }

        private void InitializeSlideMonitor()
        {
            _slideMonitorTimer = new System.Windows.Forms.Timer();
            _slideMonitorTimer.Interval = 150;
            _slideMonitorTimer.Tick += delegate(object s, EventArgs ev) { CheckCurrentSlide(); };
            _slideMonitorTimer.Start();
        }

        private void CheckCurrentSlide()
        {
            if (Application.SlideShowWindows.Count == 0)
            {
                if (_isInSlideShow)
                {
                    LogToFile("SlideShowWindows count is 0, auto-transitioning out of slide show.");
                    _isInSlideShow = false;
                    _currentSlide = null;
                    HideAllWebViews();
                }
                return;
            }

            _isInSlideShow = true;

            try
            {
                var slideShowWindow = Application.SlideShowWindows[1];
                var view = slideShowWindow.View;
                
                // NOTE: AdvanceOnClick is kept disabled for the entire time on a video slide.
                // It is only restored when the user moves to a different slide (see below).
                // Do NOT restore it on a timer — that would re-enable page-turn risk during video playback.

                // Trace view state
                LogToFile("CheckCurrentSlide tick. view.State=" + view.State);
                
                PowerPoint.Slide activeSlide = view.Slide;
                
                // Trace slide ID comparison
                int currentSlideId = _currentSlide != null ? _currentSlide.SlideID : -1;
                LogToFile(string.Format("CheckCurrentSlide state check. _currentSlideIsNull={0}, _currentSlideID={1}, activeSlideID={2}", 
                    _currentSlide == null, currentSlideId, activeSlide.SlideID));

                if (_currentSlide == null || _currentSlide.SlideID != activeSlide.SlideID)
                {
                    LogToFile("Slide changed or first check. SlideID=" + activeSlide.SlideID);
                    
                    // Restore previous slide's AdvanceOnClick if we modified it
                    RestoreSlideAdvanceOnClick();

                    _currentSlide = activeSlide;
                    LastSlideChangeTime = DateTime.Now; // Set transition start time
                    ExportCurrentSlideImage();

                    // If the new slide has media shapes or web views, disable AdvanceOnClick to prevent accidental page turns from clicks
                    if (SlideHasMedia(activeSlide) || SlideHasWebView(activeSlide))
                    {
                        try
                        {
                            _originalAdvanceOnClick = activeSlide.SlideShowTransition.AdvanceOnClick;
                            activeSlide.SlideShowTransition.AdvanceOnClick = Microsoft.Office.Core.MsoTriState.msoFalse;
                            _modifiedSlide = activeSlide;
                            LogToFile("Disabled AdvanceOnClick for web/video slide to prevent accidental page turns.");
                        }
                        catch (Exception ex)
                        {
                            LogToFile("Failed to disable AdvanceOnClick: " + ex.Message);
                        }
                    }

                    ShowWebViewForSlide(activeSlide, (IntPtr)(uint)slideShowWindow.HWND);
                }
                else
                {
                    RepositionWebViewForSlide(activeSlide, (IntPtr)(uint)slideShowWindow.HWND);
                }
            }
            catch (Exception ex)
            {
                LogToFile("CheckCurrentSlide error: " + ex.ToString());
            }
        }

        private bool SlideHasWebView(PowerPoint.Slide slide)
        {
            try
            {
                if (slide == null || slide.Shapes == null) return false;
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Name.StartsWith("WebViewContainer_"))
                        return true;
                }
            }
            catch { }
            return false;
        }

        private bool SlideHasMedia(PowerPoint.Slide slide)
        {
            try
            {
                if (slide == null || slide.Shapes == null) return false;
                return HasMediaRecursive(slide.Shapes);
            }
            catch
            {
                return false;
            }
        }

        private bool HasMediaRecursive(object shapesObj)
        {
            if (shapesObj == null) return false;
            
            if (shapesObj is PowerPoint.Shapes)
            {
                var shapes = (PowerPoint.Shapes)shapesObj;
                foreach (PowerPoint.Shape shape in shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                    {
                        try
                        {
                            if (HasMediaRecursive(shape.GroupItems))
                                return true;
                        }
                        catch { }
                    }
                    else
                    {
                        if (IsMediaShape(shape)) return true;
                    }
                }
            }
            else if (shapesObj is PowerPoint.GroupShapes)
            {
                var shapes = (PowerPoint.GroupShapes)shapesObj;
                foreach (PowerPoint.Shape shape in shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                    {
                        try
                        {
                            if (HasMediaRecursive(shape.GroupItems))
                                return true;
                        }
                        catch { }
                    }
                    else
                    {
                        if (IsMediaShape(shape)) return true;
                    }
                }
            }
            return false;
        }

        private bool IsMediaShape(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia)
                    return true;
            }
            catch { }

            try
            {
                if (shape.MediaType == PowerPoint.PpMediaType.ppMediaTypeMovie ||
                    shape.MediaType == PowerPoint.PpMediaType.ppMediaTypeSound)
                    return true;
            }
            catch { }
            
            return false;
        }

        public void RestoreSlideAdvanceOnClick()
        {
            if (_modifiedSlide != null)
            {
                try
                {
                    _modifiedSlide.SlideShowTransition.AdvanceOnClick = _originalAdvanceOnClick;
                    LogToFile(string.Format("Restored AdvanceOnClick for slide ID {0} to {1}.", _modifiedSlide.SlideID, _originalAdvanceOnClick));
                }
                catch (Exception ex)
                {
                    LogToFile("Failed to restore AdvanceOnClick: " + ex.Message);
                }
                finally
                {
                    _modifiedSlide = null;
                }
            }
        }

        private void ShowWebViewForSlide(PowerPoint.Slide slide, IntPtr slideshowHwnd)
        {
            LogToFile("ShowWebViewForSlide. SlideID=" + slide.SlideID + " HWND=" + slideshowHwnd);
            HideAllWebViews();

            try
            {
                LogToFile("Scanning shapes in slide. Count=" + slide.Shapes.Count);
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    LogToFile("  Found shape: Name=" + shape.Name);
                    if (shape.Name.StartsWith("WebViewContainer_"))
                    {
                        string url = shape.Tags["WebViewUrl"];
                        LogToFile("  WebViewContainer found! Name=" + shape.Name + " TagURL=" + url);
                        
                        if (string.IsNullOrEmpty(url))
                        {
                            LogToFile("  Warning: WebViewUrl tag is empty.");
                            continue;
                        }
                        
                        LogToFile("  Instantiating WebBrowserForm for URL: " + url);
                        var form = new WebBrowserForm(url);
                        
                        LogToFile("  Showing form overlay...");
                        form.Show(new WindowWrapper(slideshowHwnd));
                        _activeWebViews[slide] = form;
                        
                        LogToFile("  Repositioning form...");
                        RepositionForm(form, shape, slideshowHwnd);

                        try
                        {
                            shape.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                            _hiddenShapes.Add(shape);
                        }
                        catch (Exception ex)
                        {
                            LogToFile("  Failed to hide shape: " + ex.Message);
                        }

                        LogToFile("  WebViewForm display complete.");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                LogToFile("ShowWebViewForSlide error: " + ex.ToString());
            }
        }

        private void RepositionWebViewForSlide(PowerPoint.Slide slide, IntPtr slideshowHwnd)
        {
            WebBrowserForm form;
            if (_activeWebViews.TryGetValue(slide, out form))
            {
                try
                {
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.Name.StartsWith("WebViewContainer_"))
                        {
                            RepositionForm(form, shape, slideshowHwnd);
                            break;
                        }
                    }
                }
                catch { }
            }
        }

        private void RepositionForm(WebBrowserForm form, PowerPoint.Shape shape, IntPtr hwnd)
        {
            if (form.IsMaximized)
            {
                RECT winRect;
                if (GetClientRect(hwnd, out winRect))
                {
                    int maxW = winRect.Right - winRect.Left;
                    int maxH = winRect.Bottom - winRect.Top;
                    POINT maxPt = new POINT { X = 0, Y = 0 };
                    ClientToScreen(hwnd, ref maxPt);
                    if (form.Location.X != maxPt.X || form.Location.Y != maxPt.Y || form.Size.Width != maxW || form.Size.Height != maxH)
                    {
                        form.Location = new Point(maxPt.X, maxPt.Y);
                        form.Size = new Size(maxW, maxH);
                    }
                }
                return;
            }

            RECT clientRect;
            if (!GetClientRect(hwnd, out clientRect)) return;

            int winWidth = clientRect.Right - clientRect.Left;
            int winHeight = clientRect.Bottom - clientRect.Top;

            float slideWidthPts = Application.ActivePresentation.PageSetup.SlideWidth;
            float slideHeightPts = Application.ActivePresentation.PageSetup.SlideHeight;

            float arSlide = slideWidthPts / slideHeightPts;
            float arWin = (float)winWidth / winHeight;

            float slideScaledW, slideScaledH;
            float offsetX = 0, offsetY = 0;

            if (arWin > arSlide)
            {
                slideScaledH = winHeight;
                slideScaledW = winHeight * arSlide;
                offsetX = (winWidth - slideScaledW) / 2.0f;
            }
            else
            {
                slideScaledW = winWidth;
                slideScaledH = winWidth / arSlide;
                offsetY = (winHeight - slideScaledH) / 2.0f;
            }

            float leftPct = shape.Left / slideWidthPts;
            float topPct = shape.Top / slideHeightPts;
            float widthPct = shape.Width / slideWidthPts;
            float heightPct = shape.Height / slideHeightPts;

            int clientX = (int)(offsetX + leftPct * slideScaledW);
            int clientY = (int)(offsetY + topPct * slideScaledH);
            int clientW = (int)(widthPct * slideScaledW);
            int clientH = (int)(heightPct * slideScaledH);

            POINT pt = new POINT { X = clientX, Y = clientY };
            ClientToScreen(hwnd, ref pt);

            if (form.Location.X != pt.X || form.Location.Y != pt.Y || form.Size.Width != clientW || form.Size.Height != clientH)
            {
                form.Location = new Point(pt.X, pt.Y);
                form.Size = new Size(clientW, clientH);
            }
        }

        private void RestoreHiddenShapes()
        {
            foreach (var shape in _hiddenShapes)
            {
                try
                {
                    shape.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                }
                catch { }
            }
            _hiddenShapes.Clear();
        }

        private void HideAllWebViews()
        {
            LogToFile("HideAllWebViews called. ActiveCount=" + _activeWebViews.Count);
            RestoreHiddenShapes();
            foreach (var kvp in _activeWebViews)
            {
                try
                {
                    kvp.Value.Close();
                    kvp.Value.Dispose();
                }
                catch { }
            }
            _activeWebViews.Clear();
        }


        public void ToggleRemoteConsole()
        {
            if (_remoteServer == null)
            {
                StartRemoteServer();
                SettingsManager.Current.RemoteControlEnabled = true;
                SettingsManager.Save();
                
                if (_remoteServer != null && !string.IsNullOrEmpty(_remoteServer.ServerUrl))
                {
                    MessageBox.Show("局域网遥控服务已成功启动！\n服务地址: " + _remoteServer.ServerUrl + "\n\n您现在可以点击【二维码】按钮扫码连接手机了。", "遥控已开启", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                StopRemoteServer();
                SettingsManager.Current.RemoteControlEnabled = false;
                SettingsManager.Save();
                MessageBox.Show("局域网遥控服务已关闭。", "遥控已关闭", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void StartRemoteServer()
        {
            _remoteServer = new RemoteServer(SettingsManager.Current.ConsolePort, _syncContext);
            _remoteServer.Start();
        }

        private void StopRemoteServer()
        {
            if (_remoteServer != null)
            {
                _remoteServer.Stop();
                _remoteServer = null;
            }
        }

        public void ShowLiveCast(string url)
        {
            SafeInvoke(new Action(() => {
                ShowLiveCastInternal(url);
            }));
        }

        private void ShowLiveCastInternal(string url)
        {
            try
            {
                IntPtr slideshowHwnd = IntPtr.Zero;
                try
                {
                    if (Application.SlideShowWindows != null && Application.SlideShowWindows.Count > 0)
                    {
                        slideshowHwnd = (IntPtr)(uint)Application.SlideShowWindows[1].HWND;
                    }
                }
                catch { }

                if (_liveCastForm == null || _liveCastForm.IsDisposed)
                {
                    _liveCastForm = new WebBrowserForm(url);
                    _liveCastForm.IsMaximized = false;
                    if (slideshowHwnd != IntPtr.Zero)
                    {
                        _liveCastForm.Show(new WindowWrapper(slideshowHwnd));
                    }
                    else
                    {
                        _liveCastForm.Show();
                    }
                }
                else
                {
                    _liveCastForm.NavigateTo(url);
                }

                if (slideshowHwnd != IntPtr.Zero)
                {
                    RECT winRect;
                    if (GetClientRect(slideshowHwnd, out winRect))
                    {
                        int maxW = winRect.Right - winRect.Left;
                        int maxH = winRect.Bottom - winRect.Top;
                        POINT maxPt = new POINT { X = 0, Y = 0 };
                        ClientToScreen(slideshowHwnd, ref maxPt);
                        _liveCastForm.Location = new Point(maxPt.X, maxPt.Y);
                        _liveCastForm.Size = new Size(maxW, maxH);
                    }
                }
                _liveCastForm.BringToFront();
            }
            catch (Exception ex)
            {
                LogToFile("ShowLiveCast error: " + ex.Message);
            }
        }

        public void CloseLiveCast()
        {
            SafeInvoke(new Action(() => {
                CloseLiveCastInternal();
            }));
        }

        private void CloseLiveCastInternal()
        {
            try
            {
                if (_liveCastForm != null && !_liveCastForm.IsDisposed)
                {
                    _liveCastForm.Close();
                    _liveCastForm.Dispose();
                    _liveCastForm = null;
                }
            }
            catch { }
        }

        public void ShowQRCodeWindow()
        {
            if (_remoteServer != null && !string.IsNullOrEmpty(_remoteServer.ServerUrl))
            {
                using (var qrForm = new QRCodeForm(_remoteServer.ServerUrl.TrimEnd('/') + "/console"))
                {
                    qrForm.ShowDialog();
                }
            }
            else
            {
                MessageBox.Show("遥控服务器未启用。请先开启局域网遥控器！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public void OpenSettingsDialog()
        {
            bool wasRunning = (_remoteServer != null);
            int oldPort = SettingsManager.Current.ConsolePort;

            using (var settingsForm = new SettingsForm())
            {
                if (settingsForm.ShowDialog() == DialogResult.OK)
                {
                    if (wasRunning && SettingsManager.Current.ConsolePort != oldPort)
                    {
                        StopRemoteServer();
                        StartRemoteServer();
                    }
                }
            }
        }

        /// Returns the RemoteServer's LAN URL (e.g. "http://192.168.1.5:8765/") or null if not running.
        public string GetRemoteServerUrl()
        {
            if (_remoteServer != null && !string.IsNullOrEmpty(_remoteServer.ServerUrl))
                return _remoteServer.ServerUrl;
            return null;
        }

        public void EnsureRemoteServerStarted()
        {
            if (_remoteServer == null)
            {
                try
                {
                    StartRemoteServer();
                    SettingsManager.Current.RemoteControlEnabled = true;
                    SettingsManager.Save();
                }
                catch { }
            }
        }

        public void NavigateSlide(bool next)
        {
            SafeInvoke(delegate()
            {
                try
                {
                    if (Application != null && Application.SlideShowWindows != null && Application.SlideShowWindows.Count > 0)
                    {
                        var view = Application.SlideShowWindows[1].View;
                        if (view != null)
                        {
                            if (view.State == PowerPoint.PpSlideShowState.ppSlideShowPaused)
                            {
                                view.State = PowerPoint.PpSlideShowState.ppSlideShowRunning;
                            }
                            if (next) view.Next();
                            else view.Previous();
                        }
                    }
                }
                catch { }
            });
        }

        public void RepositionFormDirect(WebBrowserForm form)
        {
            if (form == null || form.IsDisposed) return;
            SafeInvoke(delegate()
            {
                try
                {
                    var screen = Screen.FromControl(form);
                    if (form.IsMaximized)
                    {
                        form.Location = screen.Bounds.Location;
                        form.Size = screen.Bounds.Size;
                    }
                    else
                    {
                        int w = Math.Min(1100, (int)(screen.WorkingArea.Width * 0.82));
                        int h = Math.Min(720, (int)(screen.WorkingArea.Height * 0.82));
                        int x = screen.WorkingArea.Left + (screen.WorkingArea.Width - w) / 2;
                        int y = screen.WorkingArea.Top + (screen.WorkingArea.Height - h) / 2;
                        form.Location = new Point(x, y);
                        form.Size = new Size(w, h);
                    }
                    form.BringToFront();
                }
                catch { }
            });
        }

        public void TriggerReposition()
        {
            if (_syncContext != null)
            {
                _syncContext.Post(new SendOrPostCallback(delegate(object state)
                {
                    CheckCurrentSlide();
                }), null);
            }
            else
            {
                CheckCurrentSlide();
            }
        }

        private void Application_SlideShowBegin(PowerPoint.SlideShowWindow Wn)
        {
            LogToFile("Application_SlideShowBegin called. HWND=" + Wn.HWND);
            _isInSlideShow = true;
            LastSlideChangeTime = DateTime.Now; // Set transition start time
            CheckCurrentSlide();
        }

        private void Application_SlideShowNextSlide(PowerPoint.SlideShowWindow Wn)
        {
            LogToFile("Application_SlideShowNextSlide called.");
            CheckCurrentSlide();
        }

        private void Application_SlideShowEnd(PowerPoint.Presentation Pres)
        {
            LogToFile("Application_SlideShowEnd called.");
            _isInSlideShow = false;
            _currentSlide = null;
            RestoreSlideAdvanceOnClick();
            HideAllWebViews();
        }

        public bool ControlWebViewVideo()
        {
            bool success = false;
            try
            {
                WebBrowserForm form = null;
                if (_currentSlide != null && _activeWebViews.TryGetValue(_currentSlide, out form))
                {
                    form.ExecuteJavaScript(@"(function() {
                        var videos = document.querySelectorAll('video');
                        if (videos.length > 0) {
                            for (var i = 0; i < videos.length; i++) {
                                var v = videos[i];
                                if (v.paused) {
                                    v.play();
                                } else {
                                    v.pause();
                                }
                            }
                            return 'controlled';
                        }
                        var audios = document.querySelectorAll('audio');
                        if (audios.length > 0) {
                            for (var i = 0; i < audios.length; i++) {
                                var a = audios[i];
                                if (a.paused) {
                                    a.play();
                                } else {
                                    a.pause();
                                }
                            }
                            return 'controlled';
                        }
                        return 'none';
                    })();");
                    success = true;
                }
            }
            catch (Exception ex)
            {
                LogToFile("ControlWebViewVideo failed: " + ex.Message);
            }
            return success;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            HideAllWebViews();
            StopRemoteServer();
        }

        public class WindowWrapper : IWin32Window
        {
            private IntPtr _handle;
            public WindowWrapper(IntPtr handle) { _handle = handle; }
            public IntPtr Handle { get { return _handle; } }
        }
    }
}
