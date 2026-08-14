using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPTWebBrowserAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PPTWebBrowserAddIn.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            try
            {
                Globals.ThisAddIn.EnsureInitialized();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("网页嵌入插件功能区加载初始化失败！\n详情: " + ex.ToString(), "加载错误", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void OnInsertWebView(Office.IRibbonControl control)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeWindow = app.ActiveWindow;
                var slide = (PowerPoint.Slide)activeWindow.View.Slide;

                // Check if the currently selected shape is an existing container to edit it
                PowerPoint.Shape selectedShape = null;
                try
                {
                    if (activeWindow.Selection != null && activeWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    {
                        selectedShape = activeWindow.Selection.ShapeRange[1];
                    }
                }
                catch { }

                bool isUpdating = (selectedShape != null && selectedShape.Name.StartsWith("WebViewContainer_"));
                string defaultUrl = isUpdating ? selectedShape.Tags["WebViewUrl"] : "https://bing.com";
                if (string.IsNullOrEmpty(defaultUrl)) defaultUrl = "https://bing.com";

                string targetUrl = defaultUrl;
                using (var dialog = new UrlInputDialog(defaultUrl))
                {
                    if (dialog.ShowDialog() != DialogResult.OK)
                    {
                        return; // User cancelled
                    }
                    targetUrl = dialog.SelectedUrl;
                }

                if (isUpdating)
                {
                    // Update existing container shape
                    selectedShape.Tags.Delete("WebViewUrl");
                    selectedShape.Tags.Add("WebViewUrl", targetUrl);

                    selectedShape.TextFrame.TextRange.Text = string.Format("🌐 网页嵌入容器 (PowerPoint Web Addin)\n\nURL: {0}\n\n[提示: 在放映 PPT 时，此区域将自动渲染为可交互的网页窗口]", targetUrl);
                    selectedShape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    selectedShape.TextFrame.TextRange.Font.Size = 14;
                    selectedShape.TextFrame.TextRange.Font.Name = "微软雅黑";
                    selectedShape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    MessageBox.Show("网页容器 URL 更新成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    // Add new container shape
                    var shape = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeRectangle,
                        100, 100, 480, 270 // Default to 16:9 ratio
                    );

                    string id = Guid.NewGuid().ToString().Substring(0, 8);
                    shape.Name = "WebViewContainer_" + id;
                    shape.Tags.Add("WebViewUrl", targetUrl);

                    shape.TextFrame.TextRange.Text = string.Format("🌐 网页嵌入容器 (PowerPoint Web Addin)\n\nURL: {0}\n\n[提示: 在放映 PPT 时，此区域将自动渲染为可交互的网页窗口]", targetUrl);
                    shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    shape.TextFrame.TextRange.Font.Size = 14;
                    shape.TextFrame.TextRange.Font.Name = "微软雅黑";
                    shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新或插入网页容器失败，请确保当前处于编辑视图下并且选中了正确的容器。\n错误详情: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnToggleConsole(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ToggleRemoteConsole();
        }

        public void OnShowQR(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ShowQRCodeWindow();
        }

        public void OnOpenSettings(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.OpenSettingsDialog();
        }

        public void OnClearCache(Office.IRibbonControl control)
        {
            try
            {
                string cachePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PPTWebBrowserAddIn");
                if (Directory.Exists(cachePath))
                {
                    Directory.Delete(cachePath, true);
                    MessageBox.Show("网页缓存清除成功，重启 PPT 并进入放映模式后生效！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("暂无网页缓存目录。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("清除缓存失败，可能由于 WebView2 进程仍在后台运行。请关闭所有 PowerPoint 窗口后重试:\n" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] names = asm.GetManifestResourceNames();
            for (int i = 0; i < names.Length; i++)
            {
                if (string.Compare(resourceName, names[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(names[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

    public class UrlInputDialog : Form
    {
        private TextBox txtUrl;
        private Button btnOk;
        private Button btnCancel;
        private Button btnBrowse;
        public string SelectedUrl { get; private set; }

        public UrlInputDialog(string defaultUrl = "https://bing.com")
        {
            this.Text = "设置网页嵌入容器";
            this.Size = new Size(420, 160);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;

            var lbl = new Label() { Text = "输入网址 (URL) 或选择本地 HTML 文件:", Location = new Point(15, 15), Size = new Size(300, 20) };
            txtUrl = new TextBox() { Text = defaultUrl, Location = new Point(15, 40), Size = new Size(280, 25) };
            
            btnBrowse = new Button() { Text = "浏览...", Location = new Point(305, 38), Size = new Size(80, 27) };
            btnBrowse.Click += (s, e) =>
            {
                using (var ofd = new OpenFileDialog())
                {
                    ofd.Filter = "HTML 文件 (*.html;*.htm)|*.html;*.htm|所有文件 (*.*)|*.*";
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        string path = ofd.FileName.Replace("\\", "/");
                        if (!path.StartsWith("/")) path = "/" + path;
                        txtUrl.Text = "file://" + path;
                    }
                }
            };

            btnOk = new Button() { Text = "确定", DialogResult = DialogResult.OK, Location = new Point(210, 85), Size = new Size(80, 25) };
            btnCancel = new Button() { Text = "取消", DialogResult = DialogResult.Cancel, Location = new Point(305, 85), Size = new Size(80, 25) };

            this.Controls.AddRange(new Control[] { lbl, txtUrl, btnBrowse, btnOk, btnCancel });
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                SelectedUrl = txtUrl.Text.Trim();
                if (string.IsNullOrEmpty(SelectedUrl))
                {
                    MessageBox.Show("地址不能为空！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                }
                else if (!SelectedUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) &&
                         !SelectedUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase) &&
                         !SelectedUrl.StartsWith("file://", StringComparison.OrdinalIgnoreCase))
                {
                    SelectedUrl = "https://" + SelectedUrl;
                }
            }
            base.OnFormClosing(e);
        }
    }
}
