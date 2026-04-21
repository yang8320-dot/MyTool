using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Threading;

public class App_Screenshot : UserControl {
    private MainForm parentForm;
    private PictureBox previewBox;
    private Label statusLabel;
    private float scale;

    public App_Screenshot(MainForm mainForm) {
        this.parentForm = mainForm;
        this.scale = this.DeviceDpi / 96f;
        this.BackColor = UITheme.BgGray;
        this.Padding = new Padding((int)(10 * scale));

        // --- 頂部控制列 ---
        Panel topPanel = new Panel() { Dock = DockStyle.Top, Height = (int)(80 * scale) };
        
        // 截圖按鈕 (Apple Blue)
        Button btnNew = new Button() { 
            Text = "截圖", 
            Left = 0, Top = (int)(5 * scale), Width = (int)(85 * scale), Height = (int)(38 * scale), 
            BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, 
            FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(10.5f, FontStyle.Bold), Cursor = Cursors.Hand 
        };
        btnNew.FlatAppearance.BorderSize = 0;
        btnNew.Click += BtnNew_Click;

        // 複製按鈕 (Apple Green)
        Button btnCopy = new Button() { 
            Text = "複製", 
            Left = (int)(95 * scale), Top = (int)(5 * scale), Width = (int)(85 * scale), Height = (int)(38 * scale), 
            BackColor = UITheme.AppleGreen, ForeColor = UITheme.CardWhite, 
            FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(10.5f, FontStyle.Bold), Cursor = Cursors.Hand 
        };
        btnCopy.FlatAppearance.BorderSize = 0;
        btnCopy.Click += BtnCopy_Click;

        // 儲存按鈕 (Gray)
        Button btnSave = new Button() { 
            Text = "儲存", 
            Left = (int)(190 * scale), Top = (int)(5 * scale), Width = (int)(85 * scale), Height = (int)(38 * scale), 
            BackColor = Color.Gainsboro, ForeColor = UITheme.TextMain, 
            FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(10.5f, FontStyle.Bold), Cursor = Cursors.Hand 
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += BtnSave_Click;

        // 清除按鈕 (Apple Red)
        Button btnClear = new Button() { 
            Text = "清除", 
            Left = (int)(285 * scale), Top = (int)(5 * scale), Width = (int)(85 * scale), Height = (int)(38 * scale), 
            BackColor = UITheme.AppleRed, ForeColor = UITheme.CardWhite, 
            FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(10.5f, FontStyle.Bold), Cursor = Cursors.Hand 
        };
        btnClear.FlatAppearance.BorderSize = 0;
        btnClear.Click += BtnClear_Click;

        // 狀態文字
        statusLabel = new Label() { 
            Text = "點擊「截圖」開始...", 
            Left = (int)(2 * scale), Top = (int)(50 * scale), AutoSize = true, 
            Font = UITheme.GetFont(10f), ForeColor = UITheme.TextSub 
        };

        topPanel.Controls.AddRange(new Control[] { btnNew, btnCopy, btnSave, btnClear, statusLabel });
        this.Controls.Add(topPanel);

        // --- 預覽區塊 (iOS 圓角白底卡片) ---
        Panel previewContainer = new Panel() { 
            Dock = DockStyle.Fill, BackColor = UITheme.CardWhite, 
            Padding = new Padding((int)(8 * scale)), Margin = new Padding(0, (int)(10 * scale), 0, 0)
        };
        
        previewContainer.Paint += (s, e) => {
            UITheme.DrawRoundedBackground(e.Graphics, new Rectangle(0, 0, previewContainer.Width - 1, previewContainer.Height - 1), (int)(10 * scale), UITheme.CardWhite);
            using (var pen = new Pen(Color.FromArgb(220, 220, 220), 1)) {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, previewContainer.Width - 1, previewContainer.Height - 1), (int)(10 * scale)));
            }
        };

        previewBox = new PictureBox() { 
            Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, 
            BackColor = Color.FromArgb(245, 245, 245) 
        };
        previewContainer.Controls.Add(previewBox);
        
        this.Controls.Add(previewContainer);
        previewContainer.BringToFront();
    }

    private void BtnNew_Click(object sender, EventArgs e) {
        parentForm.Hide();
        Application.DoEvents();
        Thread.Sleep(250); 

        using (SnippingOverlayForm snipForm = new SnippingOverlayForm()) {
            if (snipForm.ShowDialog() == DialogResult.OK && snipForm.ResultImage != null) {
                if (previewBox.Image != null) previewBox.Image.Dispose();
                previewBox.Image = (Image)snipForm.ResultImage.Clone();
                
                Clipboard.SetImage(previewBox.Image);
                statusLabel.Text = "截圖成功！已自動複製。";
                statusLabel.ForeColor = UITheme.AppleGreen;
            } else {
                statusLabel.Text = "已取消截圖。";
                statusLabel.ForeColor = UITheme.TextSub;
            }
        }
        
        // 確保喚醒後切換回截圖分頁 (Index 5)
        parentForm.ShowAppWindow(5); 
    }

    private void BtnCopy_Click(object sender, EventArgs e) {
        if (previewBox.Image == null) {
            MessageBox.Show("目前沒有可複製的截圖！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }
        Clipboard.SetImage(previewBox.Image);
        statusLabel.Text = "已重新複製到剪貼簿！";
        statusLabel.ForeColor = UITheme.AppleBlue;
    }

    private void BtnSave_Click(object sender, EventArgs e) {
        if (previewBox.Image == null) return;

        using (SaveFileDialog sfd = new SaveFileDialog()) {
            sfd.Filter = "PNG 圖片|*.png|JPEG 圖片|*.jpg";
            sfd.FileName = "截圖_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            if (sfd.ShowDialog() == DialogResult.OK) {
                ImageFormat format = sfd.FileName.EndsWith(".jpg") ? ImageFormat.Jpeg : ImageFormat.Png;
                previewBox.Image.Save(sfd.FileName, format);
                statusLabel.Text = "檔案已成功儲存！";
                statusLabel.ForeColor = UITheme.TextSub;
            }
        }
    }

    private void BtnClear_Click(object sender, EventArgs e) {
        if (previewBox.Image != null) {
            previewBox.Image.Dispose();
            previewBox.Image = null;
            statusLabel.Text = "內容已清除。";
            statusLabel.ForeColor = UITheme.TextSub;
        }
    }
}

// ==========================================
// 核心：全螢幕截圖覆蓋視窗 (支援高 DPI 與多螢幕)
// ==========================================
public class SnippingOverlayForm : Form {
    private Bitmap screenBmp;
    private Point startPt;
    private Rectangle selectRect;
    private bool isDragging = false;
    public Image ResultImage { get; private set; }

    public SnippingOverlayForm() {
        this.FormBorderStyle = FormBorderStyle.None;
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true; 
        this.ShowInTaskbar = false;
        this.Cursor = Cursors.Cross; 
        this.DoubleBuffered = true; 

        // 取得跨越多螢幕的虛擬螢幕尺寸
        this.Location = SystemInformation.VirtualScreen.Location;
        this.Size = SystemInformation.VirtualScreen.Size;

        screenBmp = new Bitmap(this.Width, this.Height);
        using (Graphics g = Graphics.FromImage(screenBmp)) {
            g.CopyFromScreen(this.Location, Point.Empty, this.Size);
        }
    }

    protected override void OnPaint(PaintEventArgs e) {
        base.OnPaint(e);
        e.Graphics.DrawImageUnscaled(screenBmp, 0, 0);
        
        // 畫一層半透明黑色遮罩
        using (SolidBrush dimBrush = new SolidBrush(Color.FromArgb(120, 0, 0, 0))) {
            e.Graphics.FillRectangle(dimBrush, this.ClientRectangle);
        }
        
        // 畫出選取的區域 (亮起)
        if (selectRect.Width > 0 && selectRect.Height > 0) {
            e.Graphics.DrawImage(screenBmp, selectRect, selectRect, GraphicsUnit.Pixel);
            using (Pen borderPen = new Pen(Color.Red, 2)) {
                e.Graphics.DrawRectangle(borderPen, selectRect);
            }
        }
    }

    protected override void OnMouseDown(MouseEventArgs e) {
        if (e.Button == MouseButtons.Left) { 
            isDragging = true; 
            startPt = e.Location; 
            selectRect = new Rectangle(e.X, e.Y, 0, 0); 
        } 
        else if (e.Button == MouseButtons.Right) { 
            this.DialogResult = DialogResult.Cancel; 
            this.Close(); 
        }
    }

    protected override void OnMouseMove(MouseEventArgs e) {
        if (isDragging) {
            int x = Math.Min(startPt.X, e.X); 
            int y = Math.Min(startPt.Y, e.Y);
            int w = Math.Abs(startPt.X - e.X); 
            int h = Math.Abs(startPt.Y - e.Y);
            selectRect = new Rectangle(x, y, w, h); 
            this.Invalidate(); 
        }
    }

    protected override void OnMouseUp(MouseEventArgs e) {
        if (isDragging) {
            isDragging = false;
            if (selectRect.Width > 5 && selectRect.Height > 5) {
                ResultImage = new Bitmap(selectRect.Width, selectRect.Height);
                using (Graphics g = Graphics.FromImage(ResultImage)) {
                    g.DrawImage(screenBmp, new Rectangle(0, 0, selectRect.Width, selectRect.Height), selectRect, GraphicsUnit.Pixel);
                }
                this.DialogResult = DialogResult.OK;
            } else { 
                this.DialogResult = DialogResult.Cancel; 
            }
            this.Close();
        }
    }

    protected override void Dispose(bool disposing) {
        if (disposing && screenBmp != null) screenBmp.Dispose();
        base.Dispose(disposing);
    }
}
