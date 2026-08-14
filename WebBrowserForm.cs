using System;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace PPTWebBrowserAddIn
{
    public partial class WebBrowserForm : Form
        {
        private WebView2 _webView;
        private string _targetUrl;
        private Panel _navPanel;
        private Button _btnBack;
        private Button _btnForward;
        private Button _btnRefresh;
        private TextBox _txtAddress;
        private Button _btnZoomOut;
        private Button _btnZoomIn;
        private Button _btnFullScreen;
        private ToolTip _toolTip;
        private static readonly double[] ZoomLevels = new double[] { 
            0.25, 0.33, 0.50, 0.67, 0.75, 0.90, 1.00, 1.10, 1.25, 1.50, 1.75, 2.00, 2.50, 3.00, 4.00, 5.00 
        };
        private int _zoomLevelIndex = 6; // default is index 6 (1.00 = 100% zoom)
        private double _customZoomOffset = 1.0;
        private string _tempUserDataFolder;

        private Panel addressPanel;
        private bool _isHeaderDragging = false;
        private Point _headerDragStart;
        private Point _formDragStartLoc;

        // Interactive Edge & Corner Resizing Elements
        private Panel _resizeGripBR;
        private Panel _borderRight;
        private Panel _borderBottom;
        private Panel _borderLeft;
        private bool _isBorderResizing = false;
        private int _resizeMode = 0; // 1: Right, 2: Bottom, 3: BottomRight, 4: Left
        private Point _resizingStartMouse;
        private Size _resizingStartSize;
        private Point _resizingStartLoc;
        private Button btnLiquidGlass; // The red traffic light button, triggers the toolkit drawer
        
        // Toolkit Drawer variables
        private bool _isToolkitOpen = false;
        private Panel _toolkitTray;
        private Button _btnInk;
        private Button _btnScreenshot;
        private Button _btnShare;
        private bool _isInkDrawing = false;

        [System.Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn(
            int nLeftRect,
            int nTopRect,
            int nRightRect,
            int nBottomRect,
            int nWidthEllipse,
            int nHeightEllipse
        );

        [System.Runtime.InteropServices.DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool DeleteObject(IntPtr hObject);

        private void ApplyRoundedCorners()
        {
            try
            {
                int cornerRadius = 18; // Apple-like smooth roundness
                IntPtr hRgn = CreateRoundRectRgn(0, 0, this.Width, this.Height, cornerRadius, cornerRadius);
                this.Region = Region.FromHrgn(hRgn);
                DeleteObject(hRgn);
            }
            catch { }
        }

        private void MakeCircular(Button btn)
        {
            try
            {
                var path = new System.Drawing.Drawing2D.GraphicsPath();
                path.AddEllipse(0, 0, btn.Width, btn.Height);
                btn.Region = new Region(path);
            }
            catch { }
        }

        private void LayoutControls()
        {
            try
            {
                int formWidth = this.ClientSize.Width;
                int formHeight = this.ClientSize.Height;

                // Adjust right-aligned buttons location dynamically (Zoom In, Zoom Out, Share, Ink)
                if (_btnZoomIn != null) _btnZoomIn.Location = new Point(formWidth - 36, 10);
                if (_btnZoomOut != null) _btnZoomOut.Location = new Point(formWidth - 66, 10);
                                if (_btnInk != null) _btnInk.Location = new Point(formWidth - 126, 10);

                // Adjust address bar panel location & width
                if (addressPanel != null)
                {
                    addressPanel.Location = new Point(146, 8);
                    addressPanel.Width = Math.Max(100, formWidth - 146 - 110);

                    var path = new System.Drawing.Drawing2D.GraphicsPath();
                    int r = 8;
                    path.AddArc(0, 0, r, r, 180, 90);
                    path.AddArc(addressPanel.Width - r, 0, r, r, 270, 90);
                    path.AddArc(addressPanel.Width - r, addressPanel.Height - r, r, r, 0, 90);
                    path.AddArc(0, addressPanel.Height - r, r, r, 90, 90);
                    addressPanel.Region = new Region(path);

                    if (_txtAddress != null)
                    {
                        _txtAddress.Location = new Point(10, 6);
                        _txtAddress.Width = addressPanel.Width - 20;
                    }
                }

                // Layout Toolkit Drawer inside _navPanel
                if (_toolkitTray != null)
                {
                    _toolkitTray.Location = new Point(0, 44);
                    _toolkitTray.Width = formWidth;
                    _toolkitTray.Height = 32;

                    int buttonWidth = 80;
                    int spacing = 15;
                    int totalWidth = (buttonWidth * 3) + (spacing * 2);
                    int startX = (formWidth - totalWidth) / 2;

                    if (_btnInk != null) _btnInk.Location = new Point(startX, 4);
                    if (_btnScreenshot != null) _btnScreenshot.Location = new Point(startX + buttonWidth + spacing, 4);
                    if (_btnShare != null) _btnShare.Location = new Point(startX + (buttonWidth + spacing) * 2, 4);
                }

                // Position Resize Grip and Border Strips (visible only when not maximized)
                bool showBorders = !this.IsMaximized;
                if (_resizeGripBR != null)
                {
                    _resizeGripBR.Visible = showBorders;
                    _resizeGripBR.Location = new Point(formWidth - 22, formHeight - 22);
                    _resizeGripBR.BringToFront();
                }
                if (_borderRight != null)
                {
                    _borderRight.Visible = showBorders;
                    _borderRight.Location = new Point(formWidth - 6, 44);
                    _borderRight.Size = new Size(6, Math.Max(0, formHeight - 44 - 22));
                    _borderRight.BringToFront();
                }
                if (_borderBottom != null)
                {
                    _borderBottom.Visible = showBorders;
                    _borderBottom.Location = new Point(0, formHeight - 6);
                    _borderBottom.Size = new Size(Math.Max(0, formWidth - 22), 6);
                    _borderBottom.BringToFront();
                }
                if (_borderLeft != null)
                {
                    _borderLeft.Visible = showBorders;
                    _borderLeft.Location = new Point(0, 44);
                    _borderLeft.Size = new Size(6, Math.Max(0, formHeight - 44 - 6));
                    _borderLeft.BringToFront();
                }
            }
            catch { }
        }

        private void ToggleToolkitDrawer()
        {
            if (_navPanel == null) return;
            _isToolkitOpen = !_isToolkitOpen;
            if (_isToolkitOpen)
            {
                _navPanel.Height = 76; // Expand for toolkit
                if (_toolkitTray != null) _toolkitTray.Visible = true;
            }
            else
            {
                _navPanel.Height = 44; // Collapse to standard Safari height
                if (_toolkitTray != null) _toolkitTray.Visible = false;
            }
            LayoutControls();
        }

        public void NavigateTo(string url)
        {
            _targetUrl = url;
            if (_txtAddress != null)
            {
                _txtAddress.Text = url;
            }
            if (_webView != null && _webView.CoreWebView2 != null)
            {
                _webView.CoreWebView2.Navigate(url);
            }
        }

        private async void ToggleInkDrawing()
        {
            if (_webView == null || _webView.CoreWebView2 == null) return;
            try
            {
                _isInkDrawing = !_isInkDrawing;
                if (_isInkDrawing)
                {
                    await InjectInkDrawingUI();
                }
                else
                {
                    string js = @"
                        (function() {
                            let container = document.getElementById('ppt-web-floating-menu-container');
                            if (container) container.parentNode.removeChild(container);

                            let style = document.getElementById('ppt-web-floating-menu-style');
                            if (style) style.parentNode.removeChild(style);

                            let canvas = document.getElementById('ppt-web-ink-canvas');
                            if (canvas) canvas.parentNode.removeChild(canvas);
                        })();
                    ";
                    await _webView.CoreWebView2.ExecuteScriptAsync(js);
                }
                if (_btnInk != null) _btnInk.Invalidate();
            }
            catch { }
        }

        private async System.Threading.Tasks.Task InjectInkDrawingUI()
        {
            if (_webView == null || _webView.CoreWebView2 == null) return;
            try
            {
                string js = @"
                        (function() {
                            let oldMenu = document.getElementById('ppt-web-floating-menu-container');
                            if (oldMenu) return;

                            const css = [
                                '#ppt-web-floating-menu-container {',
                                '    position: fixed;',
                                '    bottom: 90px;',
                                '    right: 36px;',
                                '    z-index: 1000000;',
                                '    user-select: none;',
                                '    font-family: -apple-system, BlinkMacSystemFont, ""SF Pro Display"", ""Segoe UI"", Roboto, sans-serif;',
                                '    width: 58px;',
                                '    height: 58px;',
                                '}',
                                '.ppt-skeuo-circle {',
                                '    width: 58px;',
                                '    height: 58px;',
                                '    border-radius: 50%;',
                                '    background: radial-gradient(circle at 35% 25%, rgba(255, 255, 255, 0.98), rgba(245, 247, 252, 0.75) 45%, rgba(220, 226, 240, 0.55) 80%, rgba(205, 215, 235, 0.75));',
                                '    backdrop-filter: blur(32px) saturate(220%);',
                                '    -webkit-backdrop-filter: blur(32px) saturate(220%);',
                                '    border: 1.5px solid rgba(255, 255, 255, 0.92);',
                                '    box-shadow: 0 16px 36px -4px rgba(0, 24, 64, 0.22), 0 4px 12px rgba(0, 0, 0, 0.08), inset 0 2.5px 5px rgba(255, 255, 255, 0.95), inset 0 -2px 4px rgba(0, 20, 50, 0.08);',
                                '    display: flex;',
                                '    align-items: center;',
                                '    justify-content: center;',
                                '    cursor: grab;',
                                '    transition: transform 0.22s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.22s ease;',
                                '}',
                                '.ppt-skeuo-circle:hover {',
                                '    transform: scale(1.08) translateY(-2px);',
                                '    box-shadow: 0 22px 44px -4px rgba(0, 32, 80, 0.28), 0 6px 16px rgba(0, 0, 0, 0.1), inset 0 3px 6px rgba(255, 255, 255, 1);',
                                '}',
                                '.ppt-skeuo-circle:active {',
                                '    cursor: grabbing;',
                                '    transform: scale(0.94);',
                                '}',
                                '.ppt-skeuo-circle svg {',
                                '    width: 28px;',
                                '    height: 28px;',
                                '    fill: #1d1d1f;',
                                '    filter: drop-shadow(0 1px 2px rgba(255, 255, 255, 0.8));',
                                '}',
                                '.ppt-floating-tray {',
                                '    position: absolute;',
                                '    top: 50%;',
                                '    display: flex;',
                                '    align-items: center;',
                                '    height: 56px;',
                                '    white-space: nowrap;',
                                '    background: linear-gradient(135deg, rgba(255, 255, 255, 0.88), rgba(244, 247, 253, 0.76));',
                                '    backdrop-filter: blur(32px) saturate(220%);',
                                '    -webkit-backdrop-filter: blur(32px) saturate(220%);',
                                '    border-radius: 28px;',
                                '    border: 1.5px solid rgba(255, 255, 255, 0.92);',
                                '    box-shadow: 0 20px 48px -6px rgba(0, 24, 64, 0.2), 0 6px 18px rgba(0, 0, 0, 0.08), inset 0 1.5px 3px rgba(255, 255, 255, 0.95);',
                                '    padding: 0 16px;',
                                '    opacity: 0;',
                                '    pointer-events: none;',
                                '    right: calc(100% + 14px);',
                                '    left: auto;',
                                '    transform: translateY(-50%) scaleX(0.85);',
                                '    transform-origin: right center;',
                                '    transition: opacity 0.25s cubic-bezier(0.16, 1, 0.3, 1), transform 0.25s cubic-bezier(0.16, 1, 0.3, 1);',
                                '}',
                                '.ppt-floating-tray.expand-right {',
                                '    right: auto;',
                                '    left: calc(100% + 14px);',
                                '    transform: translateY(-50%) scaleX(0.85);',
                                '    transform-origin: left center;',
                                '}',
                                '.ppt-floating-tray.expanded {',
                                '    opacity: 1;',
                                '    transform: translateY(-50%) scaleX(1);',
                                '    pointer-events: auto;',
                                '}',
                                '.ppt-tray-divider {',
                                '    width: 1.5px;',
                                '    height: 26px;',
                                '    background: rgba(180, 185, 200, 0.6);',
                                '    margin: 0 10px;',
                                '}',
                                '.ppt-color-dot {',
                                '    width: 22px;',
                                '    height: 22px;',
                                '    border-radius: 50%;',
                                '    margin: 0 6px;',
                                '    cursor: pointer;',
                                '    box-shadow: inset 0 1px 3px rgba(0,0,0,0.3), 0 2px 6px rgba(0,0,0,0.15);',
                                '    transition: transform 0.15s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.15s ease;',
                                '}',
                                '.ppt-color-dot:hover { transform: scale(1.24); }',
                                '.ppt-color-dot.active {',
                                '    box-shadow: 0 0 0 3px #007aff, 0 0 0 4.5px rgba(255,255,255,0.8), inset 0 1px 3px rgba(0,0,0,0.3);',
                                '    transform: scale(1.1);',
                                '}',
                                '.ppt-size-dot {',
                                '    display: flex;',
                                '    align-items: center;',
                                '    justify-content: center;',
                                '    width: 30px;',
                                '    height: 30px;',
                                '    margin: 0 4px;',
                                '    cursor: pointer;',
                                '    border-radius: 50%;',
                                '    transition: background 0.15s ease, transform 0.15s ease;',
                                '}',
                                '.ppt-size-dot:hover { background: rgba(0, 0, 0, 0.07); transform: scale(1.12); }',
                                '.ppt-size-dot.active { background: rgba(0, 122, 255, 0.16); transform: scale(1.1); }',
                                '.ppt-size-dot::after {',
                                '    content: """";',
                                '    display: block;',
                                '    background: #1d1d1f;',
                                '    border-radius: 50%;',
                                '}',
                                '.ppt-size-dot[data-size=""2""]::after { width: 4px; height: 4px; }',
                                '.ppt-size-dot[data-size=""5""]::after { width: 8px; height: 8px; }',
                                '.ppt-size-dot[data-size=""10""]::after { width: 14px; height: 14px; }',
                                '.ppt-tray-btn {',
                                '    width: 36px;',
                                '    height: 36px;',
                                '    border-radius: 10px;',
                                '    display: flex;',
                                '    align-items: center;',
                                '    justify-content: center;',
                                '    cursor: pointer;',
                                '    margin: 0 4px;',
                                '    transition: background 0.15s ease, transform 0.12s ease;',
                                '}',
                                '.ppt-tray-btn:hover { background: rgba(0, 0, 0, 0.08); transform: scale(1.12); }',
                                '.ppt-tray-btn:active { transform: scale(0.92); }',
                                '.ppt-tray-btn.active { background: rgba(255, 59, 48, 0.18); }',
                                '.ppt-tray-btn svg { width: 20px; height: 20px; fill: #1d1d1f; }',
                                '.ppt-tray-btn.active svg { fill: #ff3b30; }'
                            ].join('\n');

                            const style = document.createElement('style');
                            style.id = 'ppt-web-floating-menu-style';
                            style.innerHTML = css;
                            document.head.appendChild(style);

                            // High-DPI Retina Canvas & Bézier Smooth Ink Drawing Engine
                            let canvas = document.getElementById('ppt-web-ink-canvas');
                            if (!canvas) {
                                canvas = document.createElement('canvas');
                                canvas.id = 'ppt-web-ink-canvas';
                                canvas.style.position = 'fixed';
                                canvas.style.top = '0';
                                canvas.style.left = '0';
                                canvas.style.width = '100vw';
                                canvas.style.height = '100vh';
                                canvas.style.zIndex = '999999';
                                canvas.style.pointerEvents = 'auto';
                                document.body.appendChild(canvas);
                            }
                            canvas.style.display = 'block';

                            const dpr = window.devicePixelRatio || 1;
                            canvas.width = window.innerWidth * dpr;
                            canvas.height = window.innerHeight * dpr;

                            const ctx = canvas.getContext('2d');
                            ctx.scale(dpr, dpr);

                            let inkColor = 'rgba(255, 59, 48, 0.9)';
                            let inkHex = '#ff3b30';
                            let inkWidth = 4.5;
                            let isEraser = false;

                            // SVG Stylus Pen & Eraser Custom Cursors
                            const getPenCursor = (hex) => {
                                const encHex = encodeURIComponent(hex || '#ff3b30');
                                const svg = '<svg xmlns=""http://www.w3.org/2000/svg"" width=""32"" height=""32"" viewBox=""0 0 32 32""><path d=""M2 30L8 28L26 10C27.1 8.9 27.1 7.1 26 6L24 4C22.9 2.9 21.1 2.9 20 4L2 22L0 32L10 30z"" fill=""%23ffffff"" stroke=""%231d1d1f"" stroke-width=""1.6"" stroke-linejoin=""round""/><path d=""M20 4L26 10"" stroke=""%231d1d1f"" stroke-width=""1.6""/><path d=""M6 24L8 26"" stroke=""' + encHex + '"" stroke-width=""2.5""/><circle cx=""2"" cy=""30"" r=""2"" fill=""' + encHex + '""/></svg>';
                                return 'url(""data:image/svg+xml;utf8,' + svg + '"") 2 30, crosshair';
                            };

                            const eraserCursor = 'url(""data:image/svg+xml;utf8,<svg xmlns=\'http://www.w3.org/2000/svg\' width=\'32\' height=\'32\' viewBox=\'0 0 32 32\'><rect x=\'6\' y=\'8\' width=\'20\' height=\'14\' rx=\'3\' transform=\'rotate(35 16 15)\' fill=\'%23ffffff\' stroke=\'%231d1d1f\' stroke-width=\'1.8\'/><path d=\'M11 11L18 21\' transform=\'rotate(35 16 15)\' stroke=\'%23ff3b30\' stroke-width=\'2.5\'/></svg>"") 4 26, auto';

                            const updateCursor = () => {
                                canvas.style.cursor = isEraser ? eraserCursor : getPenCursor(inkHex);
                            };
                            updateCursor();

                            let drawing = false;
                            let points = [];
                            let smoothedWidth = inkWidth;
                            let history = [];

                            const pushState = () => {
                                if (history.length >= 30) history.shift();
                                history.push(ctx.getImageData(0, 0, canvas.width, canvas.height));
                            };

                            ctx.lineCap = 'round';
                            ctx.lineJoin = 'round';

                            canvas.addEventListener('mousedown', (e) => {
                                pushState();
                                drawing = true;
                                points = [{ x: e.clientX, y: e.clientY, time: Date.now() }];
                                smoothedWidth = inkWidth;

                                ctx.beginPath();
                                ctx.arc(e.clientX, e.clientY, (isEraser ? inkWidth * 2 : inkWidth / 2), 0, Math.PI * 2);
                                ctx.fillStyle = isEraser ? 'rgba(0,0,0,1)' : inkColor;
                                ctx.globalCompositeOperation = isEraser ? 'destination-out' : 'source-over';
                                ctx.fill();
                            });

                            canvas.addEventListener('mousemove', (e) => {
                                if (!drawing) return;
                                const curPt = { x: e.clientX, y: e.clientY, time: Date.now() };
                                points.push(curPt);

                                if (points.length < 3) {
                                    const b = points[0];
                                    ctx.beginPath();
                                    ctx.arc(b.x, b.y, (isEraser ? inkWidth * 2 : inkWidth / 2), 0, Math.PI * 2);
                                    ctx.fill();
                                    return;
                                }

                                const p0 = points[points.length - 3];
                                const p1 = points[points.length - 2];
                                const p2 = points[points.length - 1];

                                const xc1 = (p0.x + p1.x) / 2;
                                const yc1 = (p0.y + p1.y) / 2;
                                const xc2 = (p1.x + p2.x) / 2;
                                const yc2 = (p1.y + p2.y) / 2;

                                const dist = Math.hypot(p2.x - p1.x, p2.y - p1.y);
                                const timeDelta = Math.max(1, p2.time - p1.time);
                                const speed = dist / timeDelta;
                                const targetW = isEraser 
                                    ? inkWidth * 4.5 
                                    : Math.max(inkWidth * 0.75, Math.min(inkWidth * 1.3, inkWidth / (speed * 0.35 + 0.85)));
                                
                                smoothedWidth = smoothedWidth * 0.65 + targetW * 0.35;

                                ctx.beginPath();
                                ctx.moveTo(xc1, yc1);
                                ctx.quadraticCurveTo(p1.x, p1.y, xc2, yc2);

                                ctx.strokeStyle = isEraser ? 'rgba(0,0,0,1)' : inkColor;
                                ctx.lineWidth = smoothedWidth;
                                ctx.globalCompositeOperation = isEraser ? 'destination-out' : 'source-over';
                                ctx.stroke();
                            });

                            canvas.addEventListener('mouseup', () => { drawing = false; points = []; });
                            canvas.addEventListener('mouseleave', () => { drawing = false; points = []; });

                            window.addEventListener('resize', () => {
                                const tempCanvas = document.createElement('canvas');
                                tempCanvas.width = canvas.width;
                                tempCanvas.height = canvas.height;
                                const tempCtx = tempCanvas.getContext('2d');
                                tempCtx.drawImage(canvas, 0, 0);

                                const d = window.devicePixelRatio || 1;
                                canvas.width = window.innerWidth * d;
                                canvas.height = window.innerHeight * d;

                                ctx.scale(d, d);
                                ctx.drawImage(tempCanvas, 0, 0, tempCanvas.width / d, tempCanvas.height / d);
                                ctx.strokeStyle = inkColor;
                                ctx.lineWidth = inkWidth;
                                ctx.lineCap = 'round';
                                ctx.lineJoin = 'round';
                            });

                            const container = document.createElement('div');
                            container.id = 'ppt-web-floating-menu-container';

                            const tray = document.createElement('div');
                            tray.className = 'ppt-floating-tray';
                            tray.innerHTML = [
                                '<div class=""ppt-color-dot active"" style=""background:#ff3b30"" data-color=""rgba(255,59,48,0.9)""></div>',
                                '<div class=""ppt-color-dot"" style=""background:#007aff"" data-color=""rgba(0,122,255,0.9)""></div>',
                                '<div class=""ppt-color-dot"" style=""background:#34c759"" data-color=""rgba(52,199,89,0.9)""></div>',
                                '<div class=""ppt-color-dot"" style=""background:#1c1c1e"" data-color=""rgba(28,28,30,0.9)""></div>',
                                '<div class=""ppt-desktop-divider ppt-tray-divider""></div>',
                                '<div class=""ppt-size-dot"" data-size=""2""></div>',
                                '<div class=""ppt-size-dot active"" data-size=""5""></div>',
                                '<div class=""ppt-size-dot"" data-size=""10""></div>',
                                '<div class=""ppt-desktop-divider ppt-tray-divider""></div>',
                                '<div class=""ppt-tray-btn"" id=""ppt-btn-undo"" title=""撤销"">',
                                '    <svg viewBox=""0 0 24 24""><path d=""M12.5 8c-2.65 0-5.05.99-6.9 2.6L2 7v9h9l-3.62-3.62c1.39-1.16 3.16-1.88 5.12-1.88 3.54 0 6.55 2.31 7.6 5.5l2.37-.78C21.08 11.03 17.15 8 12.5 8z""/></svg>',
                                '</div>',
                                '<div class=""ppt-tray-btn"" id=""ppt-btn-eraser"" title=""橡皮擦"">',
                                '    <svg viewBox=""0 0 24 24""><path d=""M16.24 7.56l4.2 4.2c.78.78.78 2.05 0 2.83L14.7 20.3c-.78.78-2.05.78-2.83 0l-8.48-8.48c-.78-.78-.78-2.05 0-2.83L9.12 3.25c.78-.78 2.05-.78 2.83 0l4.29 4.31zm-7.07 7.07l2.83 2.83 4.24-4.24-2.83-2.83-4.24 4.24z""/></svg>',
                                '</div>',
                                '<div class=""ppt-tray-btn"" id=""ppt-btn-clear"" title=""清空"">',
                                '    <svg viewBox=""0 0 24 24""><path d=""M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z""/></svg>',
                                '</div>',
                                '<div class=""ppt-tray-btn"" id=""ppt-btn-share"" title=""扫码共享网页"">',
                                '    <svg viewBox=""0 0 24 24""><path d=""M4 4h6v6H4V4zm2 2v2h2V6H6zm0 7h2v2H6v-2zm2 2H6v2h2v-2zm8-11h6v6h-6V4zm2 2v2h2V6h-2zM4 14h6v6H4v-6zm2 2v2h2v-2H6zm10-12h2v2h-2V4zm0 8h2v2h-2v-2zm2 2h2v2h-2v-2zm-2 2h2v2h-2v-2zm2 2h2v2h-2v-2zm-4-4h2v2h-2v-2zm0-4h2v2h-2v-2zm2-2h2v2h-2V8zm-2-2h2v2h-2V6zm8 8h2v2h-2v-2zm0 4h2v2h-2v-2z""/></svg>',
                                '</div>'
                            ].join('');

                            const mainButton = document.createElement('div');
                            mainButton.className = 'ppt-skeuo-circle';
                            mainButton.title = '展开标注工具';
                            mainButton.innerHTML = '<svg viewBox=""0 0 24 24""><path d=""M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z""/></svg>';

                            container.appendChild(tray);
                            container.appendChild(mainButton);
                            document.body.appendChild(container);

                            let isDragging = false;
                            let dragStartX = 0, dragStartY = 0;
                            let buttonStartX = 0, buttonStartY = 0;

                            mainButton.addEventListener('mousedown', (e) => {
                                isDragging = false;
                                dragStartX = e.clientX;
                                dragStartY = e.clientY;
                                const rect = container.getBoundingClientRect();
                                buttonStartX = rect.left;
                                buttonStartY = rect.top;
                                document.addEventListener('mousemove', onMouseMove);
                                document.addEventListener('mouseup', onMouseUp);
                                e.preventDefault();
                            });

                            const onMouseMove = (e) => {
                                const dx = e.clientX - dragStartX;
                                const dy = e.clientY - dragStartY;
                                if (Math.abs(dx) > 3 || Math.abs(dy) > 3) isDragging = true;
                                if (isDragging) {
                                    container.style.right = 'auto';
                                    container.style.left = (buttonStartX + dx) + 'px';
                                    container.style.top = (buttonStartY + dy) + 'px';
                                }
                            };

                            const updateTrayDirection = () => {
                                const ballRect = container.getBoundingClientRect();
                                const ballCenterX = ballRect.left + ballRect.width / 2;
                                const isOnRight = ballCenterX > window.innerWidth / 2;
                                if (isOnRight) {
                                    tray.classList.remove('expand-right');
                                } else {
                                    tray.classList.add('expand-right');
                                }
                            };

                            const onMouseUp = () => {
                                document.removeEventListener('mousemove', onMouseMove);
                                document.removeEventListener('mouseup', onMouseUp);
                                if (!isDragging) {
                                    updateTrayDirection();
                                    tray.classList.toggle('expanded');
                                }
                                isDragging = false;
                            };

                            tray.querySelectorAll('.ppt-color-dot').forEach(dot => {
                                dot.addEventListener('click', (e) => {
                                    tray.querySelectorAll('.ppt-color-dot').forEach(d => d.classList.remove('active'));
                                    dot.classList.add('active');
                                    inkColor = dot.getAttribute('data-color');
                                    inkHex = dot.style.background || '#ff3b30';
                                    isEraser = false;
                                    tray.querySelector('#ppt-btn-eraser').classList.remove('active');
                                    updateCursor();
                                });
                            });

                            tray.querySelectorAll('.ppt-size-dot').forEach(dot => {
                                dot.addEventListener('click', (e) => {
                                    tray.querySelectorAll('.ppt-size-dot').forEach(d => d.classList.remove('active'));
                                    dot.classList.add('active');
                                    inkWidth = parseInt(dot.getAttribute('data-size')) || 4.5;
                                });
                            });

                            tray.querySelector('#ppt-btn-undo').addEventListener('click', () => {
                                if (history.length > 0) {
                                    const prev = history.pop();
                                    ctx.putImageData(prev, 0, 0);
                                }
                            });

                            tray.querySelector('#ppt-btn-eraser').addEventListener('click', () => {
                                isEraser = !isEraser;
                                tray.querySelector('#ppt-btn-eraser').classList.toggle('active', isEraser);
                                updateCursor();
                            });

                            tray.querySelector('#ppt-btn-clear').addEventListener('click', () => {
                                pushState();
                                ctx.clearRect(0, 0, canvas.width, canvas.height);
                            });

                            tray.querySelector('#ppt-btn-share').addEventListener('click', () => {
                                if (window.chrome && window.chrome.webview) {
                                    window.chrome.webview.postMessage({ action: 'share' });
                                }
                            });
                        })();
                    ";
                await _webView.CoreWebView2.ExecuteScriptAsync(js);
            }
            catch { }
        }
        private async void CaptureWebPreview()
        {
            if (_webView == null || _webView.CoreWebView2 == null) return;
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string filename = Path.Combine(desktop, "网页截图_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".png");
                
                using (var fs = new FileStream(filename, FileMode.Create))
                {
                    await _webView.CoreWebView2.CapturePreviewAsync(
                        Microsoft.Web.WebView2.Core.CoreWebView2CapturePreviewImageFormat.Png,
                        fs
                    );
                }
                MessageBox.Show("网页截图已成功保存至您的桌面！\n文件名: " + Path.GetFileName(filename), "截图成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("截图失败: " + ex.Message, "截图错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // For localhost URLs: route through the plugin's built-in reverse proxy so
        // any device on the same LAN can access the page — even if the local server
        // is only bound to 127.0.0.1.
        private static string ReplaceLANUrl(string url)
        {
            if (string.IsNullOrEmpty(url)) return url;
            Uri uri;
            try { uri = new Uri(url); } catch { return url; }

            bool isLocal = uri.Host.Equals("localhost", StringComparison.OrdinalIgnoreCase)
                        || uri.Host == "127.0.0.1"
                        || uri.Host == "::1";
            if (!isLocal) return url;

            // Start remote control server if not running
            try
            {
                if (Globals.ThisAddIn != null)
                {
                    Globals.ThisAddIn.EnsureRemoteServerStarted();
                }
            }
            catch { }

            // Set the CurrentProxyTarget so the RemoteServer's catchall proxies to it
            string portPart = uri.IsDefaultPort ? "" : ":" + uri.Port;
            RemoteServer.CurrentProxyTarget = uri.Scheme + "://localhost" + portPart;

            // Use the remote server URL directly for sharing!
            string proxyBase = null;
            try
            {
                if (Globals.ThisAddIn != null)
                {
                    proxyBase = Globals.ThisAddIn.GetRemoteServerUrl();
                }
            }
            catch { }

            if (!string.IsNullOrEmpty(proxyBase))
            {
                // Return proxy URL with path and query (e.g. http://192.168.1.5:8765/slides/index.html?param=1)
                return proxyBase.TrimEnd('/') + uri.PathAndQuery;
            }

            // Fallback: plain LAN IP substitution (requires local server to bind 0.0.0.0)
            string lanIp = GetLanIp();
            if (string.IsNullOrEmpty(lanIp)) return url;
            UriBuilder builder = new UriBuilder(uri);
            builder.Host = lanIp;
            return builder.ToString();
        }

        private static string GetLanIp()
        {
            try
            {
                foreach (System.Net.NetworkInformation.NetworkInterface ni in
                         System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces())
                {
                    if (ni.OperationalStatus != System.Net.NetworkInformation.OperationalStatus.Up) continue;
                    var t = ni.NetworkInterfaceType;
                    if (t != System.Net.NetworkInformation.NetworkInterfaceType.Wireless80211 &&
                        t != System.Net.NetworkInformation.NetworkInterfaceType.Ethernet) continue;
                    string desc = ni.Description.ToLower();
                    if (desc.Contains("virtual") || desc.Contains("vmware") ||
                        desc.Contains("hyper-v") || desc.Contains("vpn")) continue;
                    var ipProps = ni.GetIPProperties();
                    if (ipProps.GatewayAddresses.Count == 0) continue;
                    foreach (var ip in ipProps.UnicastAddresses)
                    {
                        if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                            return ip.Address.ToString();
                    }
                }
            }
            catch { }
            return null;
        }

                private void ShareViaQrCode()
        {
            string currentUrl = _targetUrl;
            if (_webView != null && _webView.CoreWebView2 != null && !string.IsNullOrEmpty(_webView.Source.ToString()))
            {
                currentUrl = _webView.Source.ToString();
            }

            string shareUrl = ReplaceLANUrl(currentUrl);
            if (shareUrl.Contains("cast_view")) { shareUrl = shareUrl.Replace("cast_view", "view_photo"); }

            try
            {
                Form qrForm = new Form();
                qrForm.Size = new Size(300, 360);
                qrForm.MinimumSize = new Size(220, 260);
                qrForm.Text = "扫码共享网页";
                qrForm.FormBorderStyle = FormBorderStyle.None;
                qrForm.StartPosition = FormStartPosition.CenterScreen;
                qrForm.TopMost = true;
                qrForm.ShowInTaskbar = false;
                qrForm.BackColor = Color.FromArgb(246, 246, 246);
                qrForm.KeyPreview = true;

                // Close on ESC or clicking outside
                qrForm.Deactivate += (s, e) => {
                    try { qrForm.Close(); qrForm.Dispose(); } catch {}
                };
                qrForm.KeyDown += (s, e) => {
                    if (e.KeyCode == Keys.Escape) {
                        try { qrForm.Close(); qrForm.Dispose(); } catch {}
                    }
                };

                // Title bar
                Panel titlePanel = new Panel();
                titlePanel.Height = 36;
                titlePanel.Dock = DockStyle.Top;
                titlePanel.BackColor = Color.FromArgb(246, 246, 246);
                titlePanel.Cursor = Cursors.SizeAll;

                // Close button (Red circle)
                Button btnClose = new Button();
                btnClose.Size = new Size(13, 13);
                btnClose.Location = new Point(12, 11);
                btnClose.FlatStyle = FlatStyle.Flat;
                btnClose.FlatAppearance.BorderSize = 0;
                btnClose.BackColor = Color.FromArgb(255, 95, 86);
                btnClose.Cursor = Cursors.Hand;
                MakeCircular(btnClose);
                btnClose.Click += (s, e) => { 
                    try { qrForm.Close(); qrForm.Dispose(); } catch {}
                };
                titlePanel.Controls.Add(btnClose);

                // Title label
                Label lblTitle = new Label();
                lblTitle.Text = "扫码共享网页";
                lblTitle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
                lblTitle.ForeColor = Color.FromArgb(74, 74, 78);
                lblTitle.Size = new Size(120, 20);
                lblTitle.Location = new Point((qrForm.Width - lblTitle.Width) / 2, 8);
                lblTitle.TextAlign = ContentAlignment.MiddleCenter;
                lblTitle.Cursor = Cursors.SizeAll;
                titlePanel.Controls.Add(lblTitle);
                qrForm.Controls.Add(titlePanel);

                // Dragging Support for Title Bar & Label
                Action<MouseEventArgs> handleDrag = (e) => {
                    if (e.Button == MouseButtons.Left) {
                        ReleaseCapture();
                        SendMessage(qrForm.Handle, 0xA1, 2, 0); // WM_NCLBUTTONDOWN, HTCAPTION
                    }
                };
                titlePanel.MouseDown += (s, e) => handleDrag(e);
                lblTitle.MouseDown += (s, e) => handleDrag(e);

                // QR Image PictureBox
                PictureBox pb = new PictureBox();
                pb.Size = new Size(240, 240);
                pb.Location = new Point(30, 46);
                pb.SizeMode = PictureBoxSizeMode.StretchImage;
                pb.BackColor = Color.White;
                qrForm.Controls.Add(pb);

                // Hint label
                Label lblHint = new Label();
                lblHint.Text = "按住标题栏拖动 · 拖动右下角调整大小";
                lblHint.Font = new Font("Segoe UI", 8.5f);
                lblHint.ForeColor = Color.FromArgb(140, 140, 145);
                lblHint.Size = new Size(280, 20);
                lblHint.Location = new Point(10, 328);
                lblHint.TextAlign = ContentAlignment.MiddleCenter;
                qrForm.Controls.Add(lblHint);

                // Corner Resize Grip
                Panel qrGrip = new Panel();
                qrGrip.Size = new Size(18, 18);
                qrGrip.Location = new Point(qrForm.ClientSize.Width - 18, qrForm.ClientSize.Height - 18);
                qrGrip.BackColor = Color.Transparent;
                qrGrip.Cursor = Cursors.SizeNWSE;
                qrGrip.Paint += (s, e) => {
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    using (var p = new Pen(Color.FromArgb(150, 150, 155), 1.2f))
                    {
                        e.Graphics.DrawLine(p, 13, 5, 5, 13);
                        e.Graphics.DrawLine(p, 13, 9, 9, 13);
                    }
                };

                bool isQrResizing = false;
                Point qrResizeStartMouse = Point.Empty;
                Size qrResizeStartSize = Size.Empty;

                qrGrip.MouseDown += (s, e) => {
                    if (e.Button == MouseButtons.Left) {
                        isQrResizing = true;
                        qrResizeStartMouse = Cursor.Position;
                        qrResizeStartSize = qrForm.Size;
                    }
                };
                qrGrip.MouseMove += (s, e) => {
                    if (isQrResizing && (Control.MouseButtons & MouseButtons.Left) != 0) {
                        Point cur = Cursor.Position;
                        int dx = cur.X - qrResizeStartMouse.X;
                        int dy = cur.Y - qrResizeStartMouse.Y;
                        int nw = Math.Max(220, qrResizeStartSize.Width + dx);
                        int nh = Math.Max(260, qrResizeStartSize.Height + dy);
                        qrForm.Size = new Size(nw, nh);
                    } else {
                        isQrResizing = false;
                    }
                };
                qrGrip.MouseUp += (s, e) => { isQrResizing = false; };
                qrForm.Controls.Add(qrGrip);
                qrGrip.BringToFront();

                // Dynamic Layout on Resize
                Action updateQrLayout = () => {
                    int w = qrForm.ClientSize.Width;
                    int h = qrForm.ClientSize.Height;
                    titlePanel.Width = w;
                    lblTitle.Location = new Point((w - lblTitle.Width) / 2, 8);
                    
                    int availableH = h - 36 - 32;
                    int availableW = w - 32;
                    int qrSide = Math.Max(120, Math.Min(availableW, availableH));
                    
                    pb.Size = new Size(qrSide, qrSide);
                    pb.Location = new Point((w - qrSide) / 2, 36 + (availableH - qrSide) / 2);
                    
                    lblHint.Location = new Point(10, h - 26);
                    lblHint.Width = w - 20;
                    
                    qrGrip.Location = new Point(w - 18, h - 18);
                    qrGrip.BringToFront();
                    
                    try {
                        IntPtr hRgn = CreateRoundRectRgn(0, 0, w, h, 16, 16);
                        qrForm.Region = Region.FromHrgn(hRgn);
                        DeleteObject(hRgn);
                    } catch {}
                };

                qrForm.Resize += (s, e) => updateQrLayout();

                // Mouse Wheel Zoom for QR Window
                MouseEventHandler wheelHandler = (s, e) => {
                    int delta = e.Delta > 0 ? 30 : -30;
                    int nw = Math.Max(220, Math.Min(800, qrForm.Width + delta));
                    int nh = Math.Max(260, Math.Min(900, qrForm.Height + (int)(delta * 1.2)));
                    qrForm.Size = new Size(nw, nh);
                };
                qrForm.MouseWheel += wheelHandler;
                pb.MouseWheel += wheelHandler;

                // Generate QR code locally offline using QRCoder
                using (var qrGenerator = new QRCoder.QRCodeGenerator())
                using (var qrCodeData = qrGenerator.CreateQrCode(shareUrl, QRCoder.QRCodeGenerator.ECCLevel.Q))
                using (var qrCode = new QRCoder.QRCode(qrCodeData))
                using (var qrCodeImage = qrCode.GetGraphic(20))
                {
                    pb.Image = new Bitmap(qrCodeImage);
                }

                updateQrLayout();
                qrForm.Show(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成二维码失败: " + ex.Message, "分享错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool ReleaseCapture();
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        public bool IsMaximized { get; set; }

        public void ToggleMaximize()
        {
            this.IsMaximized = !this.IsMaximized;
            Screen screen = Screen.FromControl(this);
            
            if (this.IsMaximized)
            {
                // Directly fill 100% full screen
                this.WindowState = FormWindowState.Normal;
                this.Location = screen.Bounds.Location;
                this.Size = screen.Bounds.Size;
                this.Region = null;
            }
            else
            {
                // Restore to comfortable centered window
                this.WindowState = FormWindowState.Normal;
                int w = Math.Min(1100, (int)(screen.WorkingArea.Width * 0.82));
                int h = Math.Min(720, (int)(screen.WorkingArea.Height * 0.82));
                int x = screen.WorkingArea.Left + (screen.WorkingArea.Width - w) / 2;
                int y = screen.WorkingArea.Top + (screen.WorkingArea.Height - h) / 2;
                this.Location = new Point(x, y);
                this.Size = new Size(w, h);
            }

            if (Globals.ThisAddIn != null)
            {
                Globals.ThisAddIn.RepositionFormDirect(this);
            }

            UpdateZoomFactor();
            ApplyRoundedCorners();
            LayoutControls();
            this.BringToFront();
        }

        public void MinimizeWindow()
        {
            Screen screen = Screen.FromControl(this);
            if (this.Size.Width <= 450)
            {
                // Already minimized -> restore to windowed mode
                this.IsMaximized = false;
                int w = Math.Min(1100, (int)(screen.WorkingArea.Width * 0.82));
                int h = Math.Min(720, (int)(screen.WorkingArea.Height * 0.82));
                int x = screen.WorkingArea.Left + (screen.WorkingArea.Width - w) / 2;
                int y = screen.WorkingArea.Top + (screen.WorkingArea.Height - h) / 2;
                this.Location = new Point(x, y);
                this.Size = new Size(w, h);
            }
            else
            {
                // Minimize to compact floating window at bottom-right (420x260)
                this.IsMaximized = false;
                int w = 420;
                int h = 260;
                int x = screen.WorkingArea.Right - w - 24;
                int y = screen.WorkingArea.Bottom - h - 24;
                this.Location = new Point(x, y);
                this.Size = new Size(w, h);
            }

            UpdateZoomFactor();
            ApplyRoundedCorners();
            LayoutControls();
            this.BringToFront();
        }

        public WebBrowserForm(string url)
        {
            InitializeComponent();
            _targetUrl = url;

            this.FormBorderStyle = FormBorderStyle.None;
            this.ShowInTaskbar = false;
            this.AutoScaleMode = AutoScaleMode.None;
            this.AutoSize = false;
            this.TopMost = true;

            this.Resize += (s, e) => { 
                UpdateZoomFactor(); 
                ApplyRoundedCorners();
                LayoutControls();
            };

            InitializeWebView();
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE
                return cp;
            }
        }

        // Chrome-style zoom using discrete zoom levels (25% to 500%)
        private void UpdateZoomFactor()
        {
            if (_webView == null || _webView.CoreWebView2 == null) return;
            try
            {
                double zoom = ZoomLevels[_zoomLevelIndex];
                _webView.ZoomFactor = zoom;

                // Dynamically update tooltips to show the current zoom percentage
                if (_toolTip != null)
                {
                    string pct = string.Format(" ({0}%)", (int)Math.Round(zoom * 100));
                    if (_btnZoomIn != null) _toolTip.SetToolTip(_btnZoomIn, "放大" + pct);
                    if (_btnZoomOut != null) _toolTip.SetToolTip(_btnZoomOut, "缩小" + pct);
                }
            }
            catch { }
        }

        private async void InitializeWebView()
        {
            try
            {
                // Create navigation panel (Safari macOS Top Bar)
                _navPanel = new Panel();
                _navPanel.Height = 44;
                _navPanel.Dock = DockStyle.Top;
                _navPanel.BackColor = Color.FromArgb(246, 246, 246);
                _navPanel.Padding = new Padding(5);

                // Smooth Window Dragging & Double-Click Maximize on NavPanel
                _navPanel.MouseDown += (s, e) => {
                    if (e.Button == MouseButtons.Left && !this.IsMaximized) {
                        _isHeaderDragging = true;
                        _headerDragStart = Cursor.Position;
                        _formDragStartLoc = this.Location;
                        ReleaseCapture();
                        SendMessage(this.Handle, 0x00A1, 2, 0); // HTCAPTION Win32 Drag
                    }
                };
                _navPanel.MouseMove += (s, e) => {
                    if (_isHeaderDragging && !this.IsMaximized && (Control.MouseButtons & MouseButtons.Left) != 0) {
                        Point cur = Cursor.Position;
                        this.Location = new Point(_formDragStartLoc.X + (cur.X - _headerDragStart.X), _formDragStartLoc.Y + (cur.Y - _headerDragStart.Y));
                    } else {
                        _isHeaderDragging = false;
                    }
                };
                _navPanel.MouseUp += (s, e) => { _isHeaderDragging = false; };
                _navPanel.DoubleClick += (s, e) => { this.ToggleMaximize(); };

                _navPanel.Paint += (s, e) => {
                    using (var pen = new Pen(Color.FromArgb(218, 218, 222), 1))
                    {
                        e.Graphics.DrawLine(pen, 0, _navPanel.Height - 1, _navPanel.Width, _navPanel.Height - 1);
                    }
                };

                _toolTip = new ToolTip();
                var toolTip = _toolTip;

                // =========================================================================
                // 🍎 Exquisite, High-Precision macOS Sonoma/Sequoia Traffic Light Dots
                // =========================================================================
                bool isTrafficHovered = false;

                // Red Close Button (macOS #FF5F56)
                btnLiquidGlass = new Button();
                btnLiquidGlass.Size = new Size(20, 20);
                btnLiquidGlass.Location = new Point(14, 12);
                btnLiquidGlass.FlatStyle = FlatStyle.Flat;
                btnLiquidGlass.FlatAppearance.BorderSize = 0;
                btnLiquidGlass.BackColor = Color.Transparent;
                btnLiquidGlass.Cursor = Cursors.Hand;
                btnLiquidGlass.Click += (s, e) => { this.Close(); };
                btnLiquidGlass.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) this.Close(); };
                btnLiquidGlass.Paint += (s, e) => {
                    var g = e.Graphics;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    Rectangle circleRect = new Rectangle(4, 4, 12, 12);
                    using (var brush = new SolidBrush(Color.FromArgb(255, 95, 86)))
                    {
                        g.FillEllipse(brush, circleRect);
                    }
                    using (var pen = new Pen(Color.FromArgb(224, 68, 62), 1f))
                    {
                        g.DrawEllipse(pen, circleRect);
                    }
                };

                // Yellow Refresh / Reload Button (macOS #FFBD2E - replaces useless minimize)
                _btnRefresh = new Button();
                _btnRefresh.Size = new Size(20, 20);
                _btnRefresh.Location = new Point(36, 12);
                _btnRefresh.FlatStyle = FlatStyle.Flat;
                _btnRefresh.FlatAppearance.BorderSize = 0;
                _btnRefresh.BackColor = Color.Transparent;
                _btnRefresh.Cursor = Cursors.Hand;
                _btnRefresh.Click += (s, e) => { 
                    try { if (_webView != null && _webView.CoreWebView2 != null) _webView.Reload(); } catch {}
                };
                _btnRefresh.MouseDown += (s, e) => { 
                    if (e.Button == MouseButtons.Left) {
                        try { if (_webView != null && _webView.CoreWebView2 != null) _webView.Reload(); } catch {}
                    }
                };
                _btnRefresh.Paint += (s, e) => {
                    var g = e.Graphics;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    Rectangle circleRect = new Rectangle(4, 4, 12, 12);
                    using (var brush = new SolidBrush(Color.FromArgb(255, 189, 46)))
                    {
                        g.FillEllipse(brush, circleRect);
                    }
                    using (var pen = new Pen(Color.FromArgb(222, 161, 35), 1f))
                    {
                        g.DrawEllipse(pen, circleRect);
                    }
                };

                // Green Fullscreen / Maximize Button (macOS #27C93F)
                _btnFullScreen = new Button();
                _btnFullScreen.Size = new Size(20, 20);
                _btnFullScreen.Location = new Point(58, 12);
                _btnFullScreen.FlatStyle = FlatStyle.Flat;
                _btnFullScreen.FlatAppearance.BorderSize = 0;
                _btnFullScreen.BackColor = Color.Transparent;
                _btnFullScreen.Cursor = Cursors.Hand;
                _btnFullScreen.Click += (s, e) => { this.ToggleMaximize(); };
                _btnFullScreen.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) this.ToggleMaximize(); };
                _btnFullScreen.Paint += (s, e) => {
                    var g = e.Graphics;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                    Rectangle circleRect = new Rectangle(4, 4, 12, 12);
                    using (var brush = new SolidBrush(Color.FromArgb(39, 201, 63)))
                    {
                        g.FillEllipse(brush, circleRect);
                    }
                    using (var pen = new Pen(Color.FromArgb(26, 171, 41), 1f))
                    {
                        g.DrawEllipse(pen, circleRect);
                    }
                };

                // Create Ink button
                _btnInk = new Button();
                _btnInk.Size = new Size(24, 24);
                _btnInk.FlatStyle = FlatStyle.Flat;
                _btnInk.FlatAppearance.BorderSize = 0;
                _btnInk.BackColor = Color.Transparent;
                _btnInk.Cursor = Cursors.Hand;
                _btnInk.Click += (s, e) => { ToggleInkDrawing(); };
                _btnInk.Paint += (s, e) => {
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    Color iconColor = _isInkDrawing ? Color.FromArgb(255, 59, 48) : Color.FromArgb(74, 74, 78);
                    using (var pen = new Pen(iconColor, 1.5f))
                    {
                        e.Graphics.DrawLine(pen, 7, 17, 17, 7);
                        e.Graphics.DrawLine(pen, 9, 19, 19, 9);
                        e.Graphics.DrawLine(pen, 7, 17, 5, 19);
                        e.Graphics.DrawLine(pen, 9, 19, 5, 19);
                        e.Graphics.DrawLine(pen, 17, 7, 19, 9);
                    }
                };

                toolTip.SetToolTip(btnLiquidGlass, "关闭 (Close)");
                toolTip.SetToolTip(_btnRefresh, "刷新网页 (Reload)");
                toolTip.SetToolTip(_btnFullScreen, "全屏 (Fullscreen)");
                toolTip.SetToolTip(_btnInk, "标注 (Ink Draw)");

                // Back Button (Modern Apple Arrow Icon)
                _btnBack = new Button();
                _btnBack.Size = new Size(24, 24);
                _btnBack.Location = new Point(82, 10);
                _btnBack.FlatStyle = FlatStyle.Flat;
                _btnBack.FlatAppearance.BorderSize = 0;
                _btnBack.BackColor = Color.Transparent;
                _btnBack.Enabled = false;
                _btnBack.Cursor = Cursors.Hand;
                _btnBack.Click += (s, e) => { if (_webView.CanGoBack) _webView.GoBack(); };
                _btnBack.Paint += (s, e) => {
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    Color arrowColor = _btnBack.Enabled ? Color.FromArgb(74, 74, 78) : Color.FromArgb(198, 198, 204);
                    using (var pen = new Pen(arrowColor, 2))
                    {
                        e.Graphics.DrawLine(pen, 14, 7, 8, 12);
                        e.Graphics.DrawLine(pen, 8, 12, 14, 17);
                    }
                };

                // Forward Button (Modern Apple Arrow Icon)
                _btnForward = new Button();
                _btnForward.Size = new Size(24, 24);
                _btnForward.Location = new Point(110, 10);
                _btnForward.FlatStyle = FlatStyle.Flat;
                _btnForward.FlatAppearance.BorderSize = 0;
                _btnForward.BackColor = Color.Transparent;
                _btnForward.Enabled = false;
                _btnForward.Cursor = Cursors.Hand;
                _btnForward.Click += (s, e) => { if (_webView.CanGoForward) _webView.GoForward(); };
                _btnForward.Paint += (s, e) => {
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    Color arrowColor = _btnForward.Enabled ? Color.FromArgb(74, 74, 78) : Color.FromArgb(198, 198, 204);
                    using (var pen = new Pen(arrowColor, 2))
                    {
                        e.Graphics.DrawLine(pen, 10, 7, 16, 12);
                        e.Graphics.DrawLine(pen, 16, 12, 10, 17);
                    }
                };

                // Safari-like Capsule Address Panel (No Anchor to avoid WinForms scaling bugs)
                addressPanel = new Panel();
                addressPanel.BackColor = Color.FromArgb(233, 233, 235);
                addressPanel.Location = new Point(196, 8);
                addressPanel.Size = new Size(this.ClientSize.Width - 196 - 80, 28);

                // Center-aligned Address TextBox
                _txtAddress = new TextBox();
                _txtAddress.Text = _targetUrl;
                _txtAddress.BorderStyle = BorderStyle.None;
                _txtAddress.BackColor = Color.FromArgb(233, 233, 235);
                _txtAddress.ForeColor = Color.FromArgb(29, 29, 31);
                _txtAddress.Font = new Font("Segoe UI", 9.5f);
                _txtAddress.Location = new Point(10, 6);
                _txtAddress.Width = addressPanel.Width - 20;
                _txtAddress.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top | AnchorStyles.Bottom;
                _txtAddress.KeyDown += (s, e) =>
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        e.SuppressKeyPress = true;
                        string url = _txtAddress.Text.Trim();
                        if (File.Exists(url))
                        {
                            string path = Path.GetFullPath(url).Replace("\\", "/");
                            if (!path.StartsWith("/")) path = "/" + path;
                            url = "file://" + path;
                        }
                        else if (!url.StartsWith("http://") && !url.StartsWith("https://") && !url.StartsWith("file://"))
                        {
                            url = "https://" + url;
                        }
                        _webView.CoreWebView2.Navigate(url);
                    }
                };
                addressPanel.Controls.Add(_txtAddress);

                // Zoom Out Button (Modern Apple UI, positioned dynamically via LayoutControls)
                _btnZoomOut = new Button();
                _btnZoomOut.Size = new Size(24, 24);
                _btnZoomOut.FlatStyle = FlatStyle.Flat;
                _btnZoomOut.FlatAppearance.BorderSize = 0;
                _btnZoomOut.BackColor = Color.Transparent;
                _btnZoomOut.Cursor = Cursors.Hand;
                _btnZoomOut.Click += (s, e) =>
                {
                    if (_zoomLevelIndex > 0)
                    {
                        _zoomLevelIndex--;
                        UpdateZoomFactor();
                    }
                };
                _btnZoomOut.Paint += (s, e) => {
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    using (var pen = new Pen(Color.FromArgb(74, 74, 78), 1.5f))
                    {
                        e.Graphics.DrawLine(pen, 7, 12, 17, 12);
                    }
                };

                // Zoom In Button (Modern Apple UI, positioned dynamically via LayoutControls)
                _btnZoomIn = new Button();
                _btnZoomIn.Size = new Size(24, 24);
                _btnZoomIn.FlatStyle = FlatStyle.Flat;
                _btnZoomIn.FlatAppearance.BorderSize = 0;
                _btnZoomIn.BackColor = Color.Transparent;
                _btnZoomIn.Cursor = Cursors.Hand;
                _btnZoomIn.Click += (s, e) =>
                {
                    if (_zoomLevelIndex < ZoomLevels.Length - 1)
                    {
                        _zoomLevelIndex++;
                        UpdateZoomFactor();
                    }
                };
                _btnZoomIn.Paint += (s, e) => {
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    using (var pen = new Pen(Color.FromArgb(74, 74, 78), 1.5f))
                    {
                        e.Graphics.DrawLine(pen, 7, 12, 17, 12);
                        e.Graphics.DrawLine(pen, 12, 7, 12, 17);
                    }
                };

                _navPanel.Controls.AddRange(new Control[] { 
                    btnLiquidGlass, _btnRefresh, _btnFullScreen, _btnBack, _btnForward, _btnInk, addressPanel, _btnZoomOut, _btnZoomIn 
                });
                this.Controls.Add(_navPanel);

                // Create WebView2
                _webView = new WebView2();
                _webView.Dock = DockStyle.Fill;
                this.Controls.Add(_webView);

                // Create Interactive Edge & Corner Resize Controls
                _resizeGripBR = new Panel();
                _resizeGripBR.Size = new Size(22, 22);
                _resizeGripBR.BackColor = Color.Transparent;
                _resizeGripBR.Cursor = Cursors.SizeNWSE;
                _resizeGripBR.Paint += (s, e) => {
                    var g = e.Graphics;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    using (var p = new Pen(Color.FromArgb(140, 140, 145), 1.5f))
                    {
                        g.DrawLine(p, 16, 6, 6, 16);
                        g.DrawLine(p, 16, 10, 10, 16);
                        g.DrawLine(p, 16, 14, 14, 16);
                    }
                };
                
                Action<int, MouseEventArgs> startResize = (mode, e) => {
                    if (e.Button == MouseButtons.Left && !this.IsMaximized) {
                        _isBorderResizing = true;
                        _resizeMode = mode;
                        _resizingStartMouse = Cursor.Position;
                        _resizingStartSize = this.Size;
                        _resizingStartLoc = this.Location;
                    }
                };

                Action handleResizeMove = () => {
                    if (_isBorderResizing && !this.IsMaximized && (Control.MouseButtons & MouseButtons.Left) != 0) {
                        Point cur = Cursor.Position;
                        int dx = cur.X - _resizingStartMouse.X;
                        int dy = cur.Y - _resizingStartMouse.Y;

                        if (_resizeMode == 3) { // Bottom-Right Corner
                            int nw = Math.Max(420, _resizingStartSize.Width + dx);
                            int nh = Math.Max(260, _resizingStartSize.Height + dy);
                            this.Size = new Size(nw, nh);
                        } else if (_resizeMode == 1) { // Right Edge
                            int nw = Math.Max(420, _resizingStartSize.Width + dx);
                            this.Width = nw;
                        } else if (_resizeMode == 2) { // Bottom Edge
                            int nh = Math.Max(260, _resizingStartSize.Height + dy);
                            this.Height = nh;
                        } else if (_resizeMode == 4) { // Left Edge
                            int nw = Math.Max(420, _resizingStartSize.Width - dx);
                            int nx = _resizingStartLoc.X + (_resizingStartSize.Width - nw);
                            this.Location = new Point(nx, this.Location.Y);
                            this.Width = nw;
                        }
                    } else {
                        _isBorderResizing = false;
                    }
                };

                _resizeGripBR.MouseDown += (s, e) => startResize(3, e);
                _resizeGripBR.MouseMove += (s, e) => handleResizeMove();
                _resizeGripBR.MouseUp += (s, e) => { _isBorderResizing = false; };

                _borderRight = new Panel();
                _borderRight.BackColor = Color.Transparent;
                _borderRight.Cursor = Cursors.SizeWE;
                _borderRight.MouseDown += (s, e) => startResize(1, e);
                _borderRight.MouseMove += (s, e) => handleResizeMove();
                _borderRight.MouseUp += (s, e) => { _isBorderResizing = false; };

                _borderBottom = new Panel();
                _borderBottom.BackColor = Color.Transparent;
                _borderBottom.Cursor = Cursors.SizeNS;
                _borderBottom.MouseDown += (s, e) => startResize(2, e);
                _borderBottom.MouseMove += (s, e) => handleResizeMove();
                _borderBottom.MouseUp += (s, e) => { _isBorderResizing = false; };

                _borderLeft = new Panel();
                _borderLeft.BackColor = Color.Transparent;
                _borderLeft.Cursor = Cursors.SizeWE;
                _borderLeft.MouseDown += (s, e) => startResize(4, e);
                _borderLeft.MouseMove += (s, e) => handleResizeMove();
                _borderLeft.MouseUp += (s, e) => { _isBorderResizing = false; };

                this.Controls.Add(_resizeGripBR);
                this.Controls.Add(_borderRight);
                this.Controls.Add(_borderBottom);
                this.Controls.Add(_borderLeft);
                _resizeGripBR.BringToFront();
                _borderRight.BringToFront();
                _borderBottom.BringToFront();
                _borderLeft.BringToFront();
                _webView.BringToFront(); // Ensure it docks correctly below top panel

                // Call layouts initialization once
                ApplyRoundedCorners();
                LayoutControls();

                                CoreWebView2EnvironmentOptions options = null;
                if (SettingsManager.Current.DisableWebSecurity)
                {
                    options = new CoreWebView2EnvironmentOptions();
                    options.AdditionalBrowserArguments = "--disable-web-security --allow-insecure-localhost";
                }

                _tempUserDataFolder = Path.Combine(
                    Path.GetTempPath(), 
                    "PPTWebBrowserAddIn_" + Guid.NewGuid().ToString().Substring(0, 8)
                );

                CoreWebView2Environment environment = null;
                try
                {
                    environment = await CoreWebView2Environment.CreateAsync(null, _tempUserDataFolder, options);
                    await _webView.EnsureCoreWebView2Async(environment);
                    _webView.CoreWebView2.WebMessageReceived += WebView_WebMessageReceived;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("初始化 WebView2 环境失败，请确保系统已安装 Microsoft Edge WebView2 Runtime。\n详情: " + ex.Message, "环境初始化错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Hook event handlers
                _webView.SourceChanged += (s, e) =>
                {
                    _txtAddress.Text = _webView.Source.ToString();
                };

                _webView.NavigationCompleted += async (s, e) =>
                {
                    _btnBack.Enabled = _webView.CanGoBack;
                    _btnForward.Enabled = _webView.CanGoForward;
                    UpdateZoomFactor();

                    // Inject CSS to hide scrollbars and presenter remote listener
                    try
                    {
                        await _webView.CoreWebView2.ExecuteScriptAsync(
                            "var style = document.createElement('style');" +
                            "style.innerHTML = '::-webkit-scrollbar { display: none; }';" +
                            "document.head.appendChild(style);"
                        );

                        // Global presenter clicker & keyboard listener for PPT slide navigation
                        string presenterScript = @"
                            (function() {
                                if (window.__yulink_presenter_hooked) return;
                                window.__yulink_presenter_hooked = true;
                                window.addEventListener('keydown', function(e) {
                                    var tag = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
                                    if (tag === 'input' || tag === 'textarea' || (document.activeElement && document.activeElement.isContentEditable)) return;
                                    
                                    if (e.key === 'PageDown' || e.key === 'ArrowRight' || e.key === ' ') {
                                        if (window.chrome && window.chrome.webview) {
                                            window.chrome.webview.postMessage({ action: 'slide_next' });
                                        }
                                    } else if (e.key === 'PageUp' || e.key === 'ArrowLeft') {
                                        if (window.chrome && window.chrome.webview) {
                                            window.chrome.webview.postMessage({ action: 'slide_prev' });
                                        }
                                    }
                                }, true);
                            })();
                        ";
                        await _webView.CoreWebView2.ExecuteScriptAsync(presenterScript);
                    }
                    catch { }

                    // Auto-inject drawing tool if it is currently enabled by user
                    if (_isInkDrawing)
                    {
                        try { await InjectInkDrawingUI(); } catch { }
                    }
                };

                _webView.CoreWebView2.NewWindowRequested += (s, e) =>
                {
                    e.Handled = true; // Block opening in external window
                    _webView.CoreWebView2.Navigate(e.Uri); // Force load in same window
                };

                UpdateZoomFactor();

                // Normalize target URL to handle missing schemes (e.g. "google.com")
                string normalizedUrl = _targetUrl;
                if (string.IsNullOrEmpty(normalizedUrl))
                {
                    normalizedUrl = "about:blank";
                }
                else
                {
                    normalizedUrl = normalizedUrl.Trim();
                    if (!normalizedUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                        !normalizedUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
                        !normalizedUrl.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                    {
                        normalizedUrl = "https://" + normalizedUrl;
                    }
                }

                try
                {
                    if (normalizedUrl.StartsWith("file://"))
                    {
                        Uri uri = new Uri(normalizedUrl);
                        string localPath = uri.LocalPath;
                        string folder = Path.GetDirectoryName(localPath);
                        string fileName = Path.GetFileName(localPath);
                        
                        _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            "local.html", 
                            folder, 
                            CoreWebView2HostResourceAccessKind.Allow
                        );
                        
                        _webView.CoreWebView2.Navigate("https://local.html/" + fileName);
                    }
                    else
                    {
                        _webView.Source = new Uri(normalizedUrl);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("加载网页失败，网址格式不正确或无法解析。\n网址: " + normalizedUrl + "\n详情: " + ex.Message, "导航错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                Globals.ThisAddIn.LogToFile("WebBrowserForm InitializeWebView outer exception: " + ex.Message);
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            
            if (_webView != null)
            {
                try { _webView.Dispose(); } catch { }
            }

            try
            {
                if (!string.IsNullOrEmpty(_tempUserDataFolder) && Directory.Exists(_tempUserDataFolder))
                {
                    Directory.Delete(_tempUserDataFolder, true);
                }
            }
            catch { }
        }

        private void ConfigureRequestInterceptor()
        {
            _webView.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.Document);
            _webView.CoreWebView2.WebResourceRequested += async (sender, args) =>
            {
                if (args.ResourceContext != CoreWebView2WebResourceContext.Document) return;

                var deferral = args.GetDeferral();
                try
                {
                    var handler = new HttpClientHandler();
                    handler.UseCookies = false;
                    
                    using (var client = new HttpClient(handler))
                    {
                        var requestMessage = new HttpRequestMessage(new HttpMethod(args.Request.Method), args.Request.Uri);

                        foreach (var header in args.Request.Headers)
                        {
                            requestMessage.Headers.TryAddWithoutValidation(header.Key, header.Value);
                        }

                        if (args.Request.Content != null)
                        {
                            requestMessage.Content = new StreamContent(args.Request.Content);
                            if (args.Request.Headers.Contains("Content-Type"))
                            {
                                requestMessage.Content.Headers.TryAddWithoutValidation("Content-Type", args.Request.Headers.GetHeader("Content-Type"));
                            }
                        }

                        var response = await client.SendAsync(requestMessage);
                        var responseStream = await response.Content.ReadAsStreamAsync();

                        var newResponse = _webView.CoreWebView2.Environment.CreateWebResourceResponse(
                            responseStream,
                            (int)response.StatusCode,
                            response.ReasonPhrase,
                            ""
                        );

                        // Helper to copy headers
                        CopyHeadersHelper(response.Headers, newResponse);
                        CopyHeadersHelper(response.Content.Headers, newResponse);

                        args.Response = newResponse;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(string.Format("Interceptor Exception: {0}", ex.Message));
                }
                finally
                {
                    deferral.Complete();
                }
            };
        }

        private void CopyHeadersHelper(System.Net.Http.Headers.HttpHeaders headers, CoreWebView2WebResourceResponse response)
        {
            foreach (var h in headers)
            {
                string name = h.Key;
                if (name.Equals("Content-Security-Policy", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("X-Frame-Options", StringComparison.OrdinalIgnoreCase) ||
                    name.Equals("Frame-Options", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                response.Headers.AppendHeader(name, string.Join(", ", h.Value));
            }
        }

        public async void ExecuteJavaScript(string script)
        {
            if (_webView != null && _webView.CoreWebView2 != null)
            {
                try
                {
                    await _webView.CoreWebView2.ExecuteScriptAsync(script);
                }
                catch (Exception ex)
                {
                    Globals.ThisAddIn.LogToFile("WebView ExecuteScriptAsync failed: " + ex.Message);
                }
            }
        }
    
        private void WebView_WebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string json = e.WebMessageAsJson;
                if (json != null)
                {
                    if (json.Contains("share"))
                    {
                        ShareViaQrCode();
                    }
                    else if (json.Contains("slide_next"))
                    {
                        if (Globals.ThisAddIn != null) Globals.ThisAddIn.NavigateSlide(true);
                    }
                    else if (json.Contains("slide_prev"))
                    {
                        if (Globals.ThisAddIn != null) Globals.ThisAddIn.NavigateSlide(false);
                    }
                }
            }
            catch { }
        }
    }
}
