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

                // Adjust right-aligned buttons location dynamically (Zoom In, Zoom Out, Share, Ink)
                if (_btnZoomIn != null)
                {
                    _btnZoomIn.Location = new Point(formWidth - 36, 10);
                }
                if (_btnZoomOut != null)
                {
                    _btnZoomOut.Location = new Point(formWidth - 66, 10);
                }
                if (_btnShare != null)
                {
                    _btnShare.Location = new Point(formWidth - 96, 10);
                }
                if (_btnInk != null)
                {
                    _btnInk.Location = new Point(formWidth - 126, 10);
                }

                // Adjust address bar panel location & width (leaving space for 4 right-aligned buttons)
                if (addressPanel != null)
                {
                    addressPanel.Location = new Point(146, 8);
                    addressPanel.Width = formWidth - 146 - 140;

                    // Reapply rounded region
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

                            const css = `
                                #ppt-web-floating-menu-container {
                                    position: fixed;
                                    bottom: 80px;
                                    right: 30px;
                                    z-index: 1000000;
                                    user-select: none;
                                    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
                                    width: 48px;
                                    height: 48px;
                                }
                                .ppt-skeuo-circle {
                                    width: 48px;
                                    height: 48px;
                                    border-radius: 50%;
                                    background: linear-gradient(135deg, #ffffff, #e1e1e6);
                                    border: 1px solid #c8c8cd;
                                    box-shadow: 0 4px 12px rgba(0,0,0,0.15), inset 0 2px 4px rgba(255,255,255,0.8), inset 0 -2px 4px rgba(0,0,0,0.05);
                                    display: flex;
                                    align-items: center;
                                    justify-content: center;
                                    cursor: grab;
                                    transition: transform 0.2s cubic-bezier(0.16, 1, 0.3, 1), box-shadow 0.2s ease;
                                }
                                .ppt-skeuo-circle:hover {
                                    transform: scale(1.05);
                                    box-shadow: 0 6px 16px rgba(0,0,0,0.2), inset 0 2px 4px rgba(255,255,255,0.9);
                                }
                                .ppt-skeuo-circle:active {
                                    cursor: grabbing;
                                    transform: scale(0.95);
                                }
                                .ppt-skeuo-circle svg {
                                    width: 22px;
                                    height: 22px;
                                    fill: #4a4a4e;
                                }
                                .ppt-floating-tray {
                                    position: absolute;
                                    top: 50%;
                                    display: flex;
                                    align-items: center;
                                    height: 46px;
                                    white-space: nowrap;
                                    background: rgba(255, 255, 255, 0.85);
                                    backdrop-filter: blur(20px);
                                    -webkit-backdrop-filter: blur(20px);
                                    border-radius: 23px;
                                    border: 1px solid rgba(210, 210, 215, 0.7);
                                    box-shadow: 0 8px 32px rgba(0,0,0,0.12);
                                    padding: 0 12px;
                                    opacity: 0;
                                    pointer-events: none;
                                    /* default: expand to the LEFT of the button */
                                    right: calc(100% + 12px);
                                    left: auto;
                                    transform: translateY(-50%) scaleX(0.85);
                                    transform-origin: right center;
                                    transition: opacity 0.28s cubic-bezier(0.16, 1, 0.3, 1),
                                                transform 0.28s cubic-bezier(0.16, 1, 0.3, 1);
                                }
                                .ppt-floating-tray.expand-right {
                                    right: auto;
                                    left: calc(100% + 12px);
                                    transform: translateY(-50%) scaleX(0.85);
                                    transform-origin: left center;
                                }
                                .ppt-floating-tray.expanded {
                                    opacity: 1;
                                    transform: translateY(-50%) scaleX(1);
                                    pointer-events: auto;
                                }
                                .ppt-tray-divider {
                                    width: 1px;
                                    height: 20px;
                                    background: rgba(210, 210, 215, 0.8);
                                    margin: 0 10px;
                                }
                                .ppt-color-dot {
                                    width: 16px;
                                    height: 16px;
                                    border-radius: 50%;
                                    margin: 0 5px;
                                    cursor: pointer;
                                    box-shadow: inset 0 1px 3px rgba(0,0,0,0.2), 0 1px 2px rgba(0,0,0,0.1);
                                    transition: transform 0.15s ease;
                                }
                                .ppt-color-dot:hover {
                                    transform: scale(1.2);
                                }
                                .ppt-color-dot.active {
                                    box-shadow: 0 0 0 2px #007aff, inset 0 1px 3px rgba(0,0,0,0.2);
                                }
                                .ppt-size-dot {
                                    display: flex;
                                    align-items: center;
                                    justify-content: center;
                                    width: 22px;
                                    height: 22px;
                                    margin: 0 3px;
                                    cursor: pointer;
                                    border-radius: 50%;
                                    transition: background 0.15s ease;
                                }
                                .ppt-size-dot:hover {
                                    background: rgba(0,0,0,0.05);
                                }
                                .ppt-size-dot.active {
                                    background: rgba(0,122,255,0.1);
                                }
                                .ppt-size-dot::after {
                                    content: '';
                                    display: block;
                                    background: #4a4a4e;
                                    border-radius: 50%;
                                }
                                .ppt-size-dot[data-size='2']::after { width: 3px; height: 3px; }
                                .ppt-size-dot[data-size='5']::after { width: 6px; height: 6px; }
                                .ppt-size-dot[data-size='10']::after { width: 10px; height: 10px; }
                                
                                .ppt-tray-btn {
                                    width: 28px;
                                    height: 28px;
                                    border-radius: 6px;
                                    display: flex;
                                    align-items: center;
                                    justify-content: center;
                                    cursor: pointer;
                                    margin: 0 3px;
                                    transition: background 0.15s ease;
                                }
                                .ppt-tray-btn:hover {
                                    background: rgba(0,0,0,0.05);
                                }
                                .ppt-tray-btn.active {
                                    background: rgba(255, 59, 48, 0.15);
                                }
                                .ppt-tray-btn svg {
                                    width: 16px;
                                    height: 16px;
                                    fill: #4a4a4e;
                                }
                                .ppt-tray-btn.active svg {
                                    fill: #ff3b30;
                                }
                            `;

                            const style = document.createElement('style');
                            style.id = 'ppt-web-floating-menu-style';
                            style.innerHTML = css;
                            document.head.appendChild(style);

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
                                canvas.style.cursor = 'crosshair';
                                canvas.style.pointerEvents = 'auto';
                                canvas.width = window.innerWidth;
                                canvas.height = window.innerHeight;
                                document.body.appendChild(canvas);
                            }
                            canvas.style.display = 'block';

                            const ctx = canvas.getContext('2d');
                            let inkColor = 'rgba(255, 59, 48, 0.85)';
                            let inkWidth = 5;
                            let isEraser = false;

                            ctx.strokeStyle = inkColor;
                            ctx.lineWidth = inkWidth;
                            ctx.lineCap = 'round';
                            ctx.lineJoin = 'round';

                            let drawing = false;
                            let lastX = 0;
                            let lastY = 0;
                            let history = [];

                            const pushState = () => {
                                if (history.length >= 30) history.shift();
                                history.push(ctx.getImageData(0, 0, canvas.width, canvas.height));
                            };

                            canvas.addEventListener('mousedown', (e) => {
                                pushState();
                                drawing = true;
                                lastX = e.clientX;
                                lastY = e.clientY;
                            });

                            canvas.addEventListener('mousemove', (e) => {
                                if (!drawing) return;
                                ctx.beginPath();
                                ctx.moveTo(lastX, lastY);
                                ctx.lineTo(e.clientX, e.clientY);
                                if (isEraser) {
                                    ctx.globalCompositeOperation = 'destination-out';
                                    ctx.lineWidth = inkWidth * 3;
                                } else {
                                    ctx.globalCompositeOperation = 'source-over';
                                    ctx.lineWidth = inkWidth;
                                    ctx.strokeStyle = inkColor;
                                }
                                ctx.stroke();
                                lastX = e.clientX;
                                lastY = e.clientY;
                            });

                            canvas.addEventListener('mouseup', () => drawing = false);
                            canvas.addEventListener('mouseleave', () => drawing = false);

                            window.addEventListener('resize', () => {
                                const tempCanvas = document.createElement('canvas');
                                tempCanvas.width = canvas.width;
                                tempCanvas.height = canvas.height;
                                const tempCtx = tempCanvas.getContext('2d');
                                tempCtx.drawImage(canvas, 0, 0);

                                canvas.width = window.innerWidth;
                                canvas.height = window.innerHeight;

                                ctx.drawImage(tempCanvas, 0, 0);
                                ctx.strokeStyle = inkColor;
                                ctx.lineWidth = inkWidth;
                                ctx.lineCap = 'round';
                                ctx.lineJoin = 'round';
                            });

                            const container = document.createElement('div');
                            container.id = 'ppt-web-floating-menu-container';

                            const tray = document.createElement('div');
                            tray.className = 'ppt-floating-tray';
                            tray.innerHTML = `
                                <div class=""ppt-color-dot active"" style=""background:#ff3b30"" data-color=""rgba(255,59,48,0.85)""></div>
                                <div class=""ppt-color-dot"" style=""background:#007aff"" data-color=""rgba(0,122,255,0.85)""></div>
                                <div class=""ppt-color-dot"" style=""background:#34c759"" data-color=""rgba(52,199,89,0.85)""></div>
                                <div class=""ppt-color-dot"" style=""background:#1c1c1e"" data-color=""rgba(28,28,30,0.85)""></div>
                                <div class=""ppt-desktop-divider ppt-tray-divider""></div>
                                <div class=""ppt-size-dot"" data-size=""2""></div>
                                <div class=""ppt-size-dot active"" data-size=""5""></div>
                                <div class=""ppt-size-dot"" data-size=""10""></div>
                                <div class=""ppt-desktop-divider ppt-tray-divider""></div>
                                <div class=""ppt-tray-btn"" id=""ppt-btn-undo"" title=""撤销"">
                                    <svg viewBox=""0 0 24 24""><path d=""M12.5 8c-2.65 0-5.05.99-6.9 2.6L2 7v9h9l-3.62-3.62c1.39-1.16 3.16-1.88 5.12-1.88 3.54 0 6.55 2.31 7.6 5.5l2.37-.78C21.08 11.03 17.15 8 12.5 8z""/></svg>
                                </div>
                                <div class=""ppt-tray-btn"" id=""ppt-btn-eraser"" title=""橡皮擦"">
                                    <svg viewBox=""0 0 24 24""><path d=""M16.24 7.56l4.2 4.2c.78.78.78 2.05 0 2.83L14.7 20.3c-.78.78-2.05.78-2.83 0l-8.48-8.48c-.78-.78-.78-2.05 0-2.83L9.12 3.25c.78-.78 2.05-.78 2.83 0l4.29 4.31zm-7.07 7.07l2.83 2.83 4.24-4.24-2.83-2.83-4.24 4.24z""/></svg>
                                </div>
                                <div class=""ppt-tray-btn"" id=""ppt-btn-clear"" title=""清空"">
                                    <svg viewBox=""0 0 24 24""><path d=""M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z""/></svg>
                                </div>
                                <div class=""ppt-tray-btn"" id=""ppt-btn-share"" title=""扫码共享网页"">
                                    <svg viewBox=""0 0 24 24""><path d=""M4 4h6v6H4V4zm2 2v2h2V6H6zm0 7h2v2H6v-2zm2 2H6v2h2v-2zm8-11h6v6h-6V4zm2 2v2h2V6h-2zM4 14h6v6H4v-6zm2 2v2h2v-2H6zm10-12h2v2h-2V4zm0 8h2v2h-2v-2zm2 2h2v2h-2v-2zm-2 2h2v2h-2v-2zm2 2h2v2h-2v-2zm-4-4h2v2h-2v-2zm0-4h2v2h-2v-2zm2-2h2v2h-2V8zm-2-2h2v2h-2V6zm8 8h2v2h-2v-2zm0 4h2v2h-2v-2z""/></svg>
                                </div>
                            `;

                            const mainButton = document.createElement('div');
                            mainButton.className = 'ppt-skeuo-circle';
                            mainButton.title = '展开标注工具';
                            mainButton.innerHTML = `<svg viewBox=""0 0 24 24""><path d=""M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z""/></svg>`;

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
                                    // Ball on right → expand tray to the LEFT
                                    tray.classList.remove('expand-right');
                                } else {
                                    // Ball on left → expand tray to the RIGHT
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
                            };

                            let drawingActive = true;

                            const setDrawingActive = (active) => {
                                drawingActive = active;
                                canvas.style.pointerEvents = active ? 'auto' : 'none';
                                canvas.style.cursor = active ? 'crosshair' : 'default';
                                // Dim the circle button when paused
                                mainButton.style.opacity = active ? '1' : '0.5';
                                mainButton.title = active ? '标注中（点击颜色可暂停）' : '已暂停 – 可操作网页（点击颜色恢复）';
                            };

                            tray.querySelectorAll('.ppt-color-dot').forEach(dot => {
                                dot.addEventListener('click', () => {
                                    const wasActiveColor = dot.classList.contains('active');
                                    if (wasActiveColor && drawingActive) {
                                        // Second click on same color: suspend drawing, hand control back to webpage
                                        setDrawingActive(false);
                                        tray.querySelectorAll('.ppt-color-dot').forEach(d => d.classList.remove('active'));
                                        dot.classList.add('active'); // keep colour selected but paused
                                        dot.style.boxShadow = 'inset 0 1px 3px rgba(0,0,0,0.2), 0 0 0 2px rgba(0,0,0,0.15)';
                                    } else {
                                        // First click or different colour: activate drawing
                                        setDrawingActive(true);
                                        isEraser = false;
                                        tray.querySelector('#ppt-btn-eraser').classList.remove('active');
                                        tray.querySelectorAll('.ppt-color-dot').forEach(d => {
                                            d.classList.remove('active');
                                            d.style.boxShadow = '';
                                        });
                                        dot.classList.add('active');
                                        dot.style.boxShadow = '';
                                        inkColor = dot.getAttribute('data-color');
                                        ctx.globalCompositeOperation = 'source-over';
                                        ctx.strokeStyle = inkColor;
                                    }
                                });
                            });

                            tray.querySelectorAll('.ppt-size-dot').forEach(dot => {
                                dot.addEventListener('click', () => {
                                    tray.querySelectorAll('.ppt-size-dot').forEach(d => d.classList.remove('active'));
                                    dot.classList.add('active');
                                    inkWidth = parseInt(dot.getAttribute('data-size'));
                                    ctx.lineWidth = inkWidth;
                                });
                            });

                            tray.querySelector('#ppt-btn-undo').addEventListener('click', () => {
                                if (history.length > 0) {
                                    let previousState = history.pop();
                                    ctx.putImageData(previousState, 0, 0);
                                }
                            });

                            tray.querySelector('#ppt-btn-eraser').addEventListener('click', () => {
                                isEraser = !isEraser;
                                if (isEraser) {
                                    tray.querySelector('#ppt-btn-eraser').classList.add('active');
                                    tray.querySelectorAll('.ppt-color-dot').forEach(d => d.classList.remove('active'));
                                } else {
                                    tray.querySelector('#ppt-btn-eraser').classList.remove('active');
                                    const activeColor = tray.querySelector(`.ppt-color-dot[data-color=""${inkColor}""]`);
                                    if (activeColor) activeColor.classList.add('active');
                                }
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
            if (_webView == null) return;
            string currentUrl = (_webView.Source != null) ? _webView.Source.ToString() : null;
            if (string.IsNullOrEmpty(currentUrl)) return;

            // Replace localhost / 127.0.0.1 with the machine's LAN IP so mobile
            // devices on the same network can access the page directly.
            string shareUrl = ReplaceLANUrl(currentUrl);

            try
            {
                Form qrForm = new Form();
                qrForm.Size = new Size(280, 335);
                qrForm.Text = "扫码共享网页";
                qrForm.FormBorderStyle = FormBorderStyle.None;
                qrForm.StartPosition = FormStartPosition.CenterParent;
                qrForm.TopMost = true;
                qrForm.BackColor = Color.FromArgb(246, 246, 246);

                // Add Apple-style Title Bar
                Panel titlePanel = new Panel();
                titlePanel.Height = 36;
                titlePanel.Dock = DockStyle.Top;
                titlePanel.BackColor = Color.FromArgb(246, 246, 246);
                titlePanel.Paint += (s, e) => {
                    using (var pen = new Pen(Color.FromArgb(210, 210, 215), 1)) {
                        e.Graphics.DrawLine(pen, 0, titlePanel.Height - 1, titlePanel.Width, titlePanel.Height - 1);
                    }
                };

                // Close Button (Red Circle)
                Button btnClose = new Button();
                btnClose.Size = new Size(11, 11);
                btnClose.Location = new Point(12, 12);
                btnClose.FlatStyle = FlatStyle.Flat;
                btnClose.FlatAppearance.BorderSize = 0;
                btnClose.BackColor = Color.FromArgb(255, 95, 86);
                MakeCircular(btnClose);
                btnClose.Click += (s, e) => { qrForm.Close(); };
                titlePanel.Controls.Add(btnClose);

                // Title label
                Label lblTitle = new Label();
                lblTitle.Text = "扫码共享网页";
                lblTitle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
                lblTitle.ForeColor = Color.FromArgb(74, 74, 78);
                lblTitle.Size = new Size(120, 20);
                lblTitle.Location = new Point((qrForm.Width - lblTitle.Width) / 2, 8);
                lblTitle.TextAlign = ContentAlignment.MiddleCenter;
                titlePanel.Controls.Add(lblTitle);

                qrForm.Controls.Add(titlePanel);

                // PictureBox to display QR code
                PictureBox pb = new PictureBox();
                pb.Size = new Size(250, 250);
                pb.Location = new Point(15, 50);
                pb.SizeMode = PictureBoxSizeMode.StretchImage;
                pb.BackColor = Color.White;
                qrForm.Controls.Add(pb);

                // Hint label
                Label lblHint = new Label();
                // If URL was rewritten, show the LAN URL so user knows what was encoded
                lblHint.Text = shareUrl.Equals(currentUrl) ? "使用手机扫码同步浏览此网页" : "已将 localhost 替换为局域网 IP";
                lblHint.Font = new Font("Segoe UI", 8.5f);
                lblHint.ForeColor = Color.FromArgb(120, 120, 125);
                lblHint.Size = new Size(250, 20);
                lblHint.Location = new Point(15, 305);
                lblHint.TextAlign = ContentAlignment.MiddleCenter;
                qrForm.Controls.Add(lblHint);

                // Generate QR code locally offline using QRCoder
                using (var qrGenerator = new QRCoder.QRCodeGenerator())
                using (var qrCodeData = qrGenerator.CreateQrCode(shareUrl, QRCoder.QRCodeGenerator.ECCLevel.Q))
                using (var qrCode = new QRCoder.QRCode(qrCodeData))
                using (var qrCodeImage = qrCode.GetGraphic(20))
                {
                    pb.Image = new Bitmap(qrCodeImage);
                }

                // Apply rounded corners to QR form
                qrForm.Load += (s, e) => {
                    IntPtr hRgn = CreateRoundRectRgn(0, 0, qrForm.Width, qrForm.Height, 14, 14);
                    qrForm.Region = Region.FromHrgn(hRgn);
                    DeleteObject(hRgn);
                };

                qrForm.ShowDialog(this);
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成二维码失败: " + ex.Message, "分享错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        public bool IsMaximized { get; set; }

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
                // Create navigation panel (Safari-like top bar)
                _navPanel = new Panel();
                _navPanel.Height = 44; // Touch-friendly Apple height
                _navPanel.Dock = DockStyle.Top;
                _navPanel.BackColor = Color.FromArgb(246, 246, 246);
                _navPanel.Padding = new Padding(5);
                _navPanel.Paint += (s, e) => {
                    using (var pen = new Pen(Color.FromArgb(210, 210, 215), 1))
                    {
                        e.Graphics.DrawLine(pen, 0, _navPanel.Height - 1, _navPanel.Width, _navPanel.Height - 1);
                    }
                };

                _toolTip = new ToolTip();
                var toolTip = _toolTip;

                // btnLiquidGlass (Red macOS Traffic Light close button)
                btnLiquidGlass = new Button();
                btnLiquidGlass.Size = new Size(13, 13);
                btnLiquidGlass.Location = new Point(16, 15);
                btnLiquidGlass.FlatStyle = FlatStyle.Flat;
                btnLiquidGlass.FlatAppearance.BorderSize = 0;
                btnLiquidGlass.BackColor = Color.FromArgb(255, 95, 86);
                btnLiquidGlass.Cursor = Cursors.Hand;
                MakeCircular(btnLiquidGlass);
                btnLiquidGlass.Click += (s, e) =>
                {
                    this.Close();
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

                // Create QR code scan button
                _btnShare = new Button();
                _btnShare.Size = new Size(24, 24);
                _btnShare.FlatStyle = FlatStyle.Flat;
                _btnShare.FlatAppearance.BorderSize = 0;
                _btnShare.BackColor = Color.Transparent;
                _btnShare.Cursor = Cursors.Hand;
                _btnShare.Click += (s, e) => { ShareViaQrCode(); };
                _btnShare.Paint += (s, e) => {
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    using (var pen = new Pen(Color.FromArgb(74, 74, 78), 1.5f))
                    {
                        e.Graphics.DrawRectangle(pen, 5, 5, 14, 14);
                        using (var brush = new SolidBrush(Color.FromArgb(74, 74, 78)))
                        {
                            e.Graphics.FillRectangle(brush, 7, 7, 3, 3);
                            e.Graphics.FillRectangle(brush, 14, 7, 3, 3);
                            e.Graphics.FillRectangle(brush, 7, 14, 3, 3);
                            e.Graphics.FillRectangle(brush, 13, 13, 2, 2);
                        }
                    }
                };

                // Yellow Refresh Button (macOS Traffic Light)
                _btnRefresh = new Button();
                _btnRefresh.Size = new Size(13, 13);
                _btnRefresh.Location = new Point(36, 15);
                _btnRefresh.FlatStyle = FlatStyle.Flat;
                _btnRefresh.FlatAppearance.BorderSize = 0;
                _btnRefresh.BackColor = Color.FromArgb(254, 188, 44);
                _btnRefresh.Cursor = Cursors.Hand;
                MakeCircular(_btnRefresh);
                _btnRefresh.Click += (s, e) => { _webView.Reload(); };

                // Green Fullscreen Button (macOS Traffic Light)
                _btnFullScreen = new Button();
                _btnFullScreen.Size = new Size(13, 13);
                _btnFullScreen.Location = new Point(56, 15);
                _btnFullScreen.FlatStyle = FlatStyle.Flat;
                _btnFullScreen.FlatAppearance.BorderSize = 0;
                _btnFullScreen.BackColor = Color.FromArgb(39, 201, 63);
                _btnFullScreen.Cursor = Cursors.Hand;
                MakeCircular(_btnFullScreen);
                _btnFullScreen.Click += (s, e) =>
                {
                    this.IsMaximized = !this.IsMaximized;
                    Globals.ThisAddIn.TriggerReposition();
                    UpdateZoomFactor();
                };

                toolTip.SetToolTip(btnLiquidGlass, "关闭 (Close)");
                toolTip.SetToolTip(_btnRefresh, "刷新 (Reload)");
                toolTip.SetToolTip(_btnFullScreen, "全屏 (Fullscreen)");
                toolTip.SetToolTip(_btnInk, "标注 (Ink Draw)");
                toolTip.SetToolTip(_btnShare, "扫码共享 (QR Share)");

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
                    btnLiquidGlass, _btnRefresh, _btnFullScreen, _btnBack, _btnForward, _btnInk, _btnShare, addressPanel, _btnZoomOut, _btnZoomIn 
                });
                this.Controls.Add(_navPanel);

                // Create WebView2
                _webView = new WebView2();
                _webView.Dock = DockStyle.Fill;
                this.Controls.Add(_webView);
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
