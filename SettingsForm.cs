using System;
using System.Drawing;
using System.Windows.Forms;

namespace PPTWebBrowserAddIn
{
    public class SettingsForm : Form
    {
        private TextBox txtDefaultUrl;
        private NumericUpDown numPort;
        private CheckBox chkDisableSecurity;
        private Button btnSave;
        private Button btnCancel;

        public SettingsForm()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.Text = "安全与缓存设置";
            this.Size = new Size(400, 280);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Segoe UI", 9.5f);

            var lblUrl = new Label { Text = "默认加载网址 (Default URL):", Location = new Point(20, 20), Size = new Size(360, 20) };
            txtDefaultUrl = new TextBox { Location = new Point(20, 45), Size = new Size(340, 25) };

            var lblPort = new Label { Text = "局域网控制台端口 (Console Port):", Location = new Point(20, 85), Size = new Size(200, 20) };
            numPort = new NumericUpDown { Location = new Point(20, 110), Size = new Size(100, 25), Minimum = 1024, Maximum = 65535 };

            chkDisableSecurity = new CheckBox 
            { 
                Text = "绕过跨域限制与 CSP (跨域/网页防嵌套安全过滤)", 
                Location = new Point(20, 150), 
                Size = new Size(360, 25),
                Checked = true
            };

            btnSave = new Button { Text = "保存", Location = new Point(200, 200), Size = new Size(80, 30), DialogResult = DialogResult.OK };
            btnCancel = new Button { Text = "取消", Location = new Point(290, 200), Size = new Size(80, 30), DialogResult = DialogResult.Cancel };

            btnSave.Click += BtnSave_Click;

            this.Controls.AddRange(new Control[] { lblUrl, txtDefaultUrl, lblPort, numPort, chkDisableSecurity, btnSave, btnCancel });
            
            // Apply a cleaner button style
            btnSave.BackColor = Color.FromArgb(0, 113, 227);
            btnSave.ForeColor = Color.White;
            btnSave.FlatStyle = FlatStyle.Flat;
            btnSave.FlatAppearance.BorderSize = 0;
            
            btnCancel.FlatStyle = FlatStyle.Flat;
            btnCancel.FlatAppearance.BorderColor = Color.LightGray;
        }

        private void LoadSettings()
        {
            var config = SettingsManager.Current;
            txtDefaultUrl.Text = config.DefaultUrl;
            numPort.Value = config.ConsolePort;
            chkDisableSecurity.Checked = config.DisableWebSecurity;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            var config = SettingsManager.Current;
            config.DefaultUrl = txtDefaultUrl.Text;
            config.ConsolePort = (int)numPort.Value;
            config.DisableWebSecurity = chkDisableSecurity.Checked;
            SettingsManager.Save();
            this.Close();
        }
    }
}
