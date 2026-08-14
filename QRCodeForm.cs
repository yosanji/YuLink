using QRCoder;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace PPTWebBrowserAddIn
{
    public class QRCodeForm : Form
    {
        private string _baseUrl;
        private int _port = 8765;
        private PictureBox _picQR;
        private Label _lblUrl;
        private ComboBox _cboIp;

        public QRCodeForm(string url)
        {
            _baseUrl = url;
            try
            {
                var uri = new Uri(url);
                _port = uri.Port;
            }
            catch { }

            InitializeComponent();
            LoadAvailableIps();
            GenerateQR(GetCurrentUrl());
        }

        private void InitializeComponent()
        {
            this.Text = "YuLink 手机连接二维码";
            this.Size = new Size(340, 440);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Segoe UI", 9.5f);
            this.BackColor = Color.FromArgb(248, 249, 251);

            _picQR = new PictureBox
            {
                Location = new Point(45, 20),
                Size = new Size(240, 240),
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            var lblSelect = new Label
            {
                Text = "局域网 IP (手机与电脑需在同一 Wi-Fi)：",
                Location = new Point(25, 270),
                Size = new Size(280, 20),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                ForeColor = Color.FromArgb(100, 100, 105)
            };

            _cboIp = new ComboBox
            {
                Location = new Point(25, 292),
                Size = new Size(280, 26),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.0f)
            };
            _cboIp.SelectedIndexChanged += (s, e) => {
                GenerateQR(GetCurrentUrl());
            };

            _lblUrl = new Label
            {
                Text = _baseUrl,
                Location = new Point(20, 328),
                Size = new Size(290, 40),
                TextAlign = ContentAlignment.TopCenter,
                Font = new Font("Segoe UI", 9.0f, FontStyle.Bold),
                ForeColor = Color.FromArgb(48, 140, 255)
            };

            this.Controls.AddRange(new Control[] { _picQR, lblSelect, _cboIp, _lblUrl });
        }

        private void LoadAvailableIps()
        {
            try
            {
                var ips = RemoteServer.GetAllAvailableIps();
                int selectIdx = 0;
                string bestIp = RemoteServer.GetBestLocalIpAddress();

                for (int i = 0; i < ips.Count; i++)
                {
                    _cboIp.Items.Add(ips[i].Value);
                    if (ips[i].Key == bestIp)
                    {
                        selectIdx = i;
                    }
                }

                if (_cboIp.Items.Count > 0)
                {
                    _cboIp.SelectedIndex = selectIdx;
                }
            }
            catch
            {
                _cboIp.Items.Add("127.0.0.1 (Localhost)");
                _cboIp.SelectedIndex = 0;
            }
        }

        private string GetCurrentUrl()
        {
            try
            {
                if (_cboIp.SelectedItem != null)
                {
                    string itemText = _cboIp.SelectedItem.ToString();
                    int start = itemText.IndexOf('(');
                    int end = itemText.IndexOf(')');
                    if (start >= 0 && end > start)
                    {
                        string ip = itemText.Substring(start + 1, end - start - 1);
                        return string.Format("http://{0}:{1}/console", ip, _port);
                    }
                }
            }
            catch { }
            return _baseUrl;
        }

        private void GenerateQR(string targetUrl)
        {
            try
            {
                _lblUrl.Text = targetUrl;
                using (var qrGenerator = new QRCodeGenerator())
                using (var qrCodeData = qrGenerator.CreateQrCode(targetUrl, QRCodeGenerator.ECCLevel.Q))
                using (var qrCode = new QRCode(qrCodeData))
                using (var qrCodeImage = qrCode.GetGraphic(20))
                {
                    _picQR.Image = new Bitmap(qrCodeImage);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成二维码失败: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
