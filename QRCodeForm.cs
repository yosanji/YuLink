using QRCoder;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace PPTWebBrowserAddIn
{
    public class QRCodeForm : Form
    {
        private string _url;
        private PictureBox _picQR;
        private Label _lblUrl;

        public QRCodeForm(string url)
        {
            _url = url;
            InitializeComponent();
            GenerateQR();
        }

        private void InitializeComponent()
        {
            this.Text = "局域网遥控二维码";
            this.Size = new Size(320, 380);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.Font = new Font("Segoe UI", 9.5f);
            this.BackColor = Color.White;

            _picQR = new PictureBox
            {
                Location = new Point(35, 30),
                Size = new Size(230, 230),
                SizeMode = PictureBoxSizeMode.Zoom
            };

            _lblUrl = new Label
            {
                Text = "请使用手机扫描二维码，连接同一局域网以遥控 PPT 播放：\n" + _url,
                Location = new Point(20, 280),
                Size = new Size(260, 50),
                TextAlign = ContentAlignment.TopCenter,
                ForeColor = Color.DimGray
            };

            this.Controls.AddRange(new Control[] { _picQR, _lblUrl });
        }

        private void GenerateQR()
        {
            try
            {
                using (var qrGenerator = new QRCodeGenerator())
                using (var qrCodeData = qrGenerator.CreateQrCode(_url, QRCodeGenerator.ECCLevel.Q))
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
