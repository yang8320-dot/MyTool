using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.Linq;
using Microsoft.Data.Sqlite;

public class App_Shortcuts : UserControl {
    private MainForm parentForm;
    private FlowLayoutPanel taskPanel;

    // 拖曳排序相關變數
    private int dragInsertIndex = -1; 
    private float scale;

    public class ShortcutItem {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public int OrderIndex { get; set; }
    }
    
    // UI 綁定的資料集合
    public List<ShortcutItem> shortcuts = new List<ShortcutItem>();

    public App_Shortcuts(MainForm mainForm) {
        this.parentForm = mainForm;
        this.scale = this.DeviceDpi / 96f;
        
        this.BackColor = UITheme.BgGray;
        this.Padding = new Padding((int)(10 * scale)); 

        // --- 頂部標題列 ---
        TableLayoutPanel header = new TableLayoutPanel() { 
            Dock = DockStyle.Top, 
            Height = (int)(45 * scale), 
            ColumnCount = 2 
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(100 * scale)));

        Label lblTitle = new Label() { 
            Text = "常用捷徑", 
            Font = UITheme.GetFont(12f, FontStyle.Bold), 
            Dock = DockStyle.Fill, 
            TextAlign = ContentAlignment.MiddleLeft, 
            Padding = new Padding((int)(5 * scale), 0, 0, 0),
            ForeColor = UITheme.TextMain
        };
        
        Button btnAdd = new Button() { 
            Text = "新增", 
            Dock = DockStyle.Fill, 
            FlatStyle = FlatStyle.Flat, 
            Margin = new Padding((int)(2 * scale), (int)(6 * scale), (int)(2 * scale), (int)(8 * scale)), 
            Cursor = Cursors.Hand, 
            BackColor = UITheme.AppleBlue, 
            ForeColor = UITheme.CardWhite,
            Font = UITheme.GetFont(10f, FontStyle.Bold)
        };
        btnAdd.FlatAppearance.BorderSize = 0; 
        btnAdd.Click += (s, e) => { new EditShortcutWindow(this, null).ShowDialog(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnAdd, 1, 0);
        this.Controls.Add(header);

        // --- 捷徑容器 ---
        taskPanel = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, 
            AutoScroll = true, 
            FlowDirection = FlowDirection.TopDown, 
            WrapContents = false, 
            BackColor = UITheme.BgGray,
            AllowDrop = true 
        };

        // 綁定拖曳相關事件
        taskPanel.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
        taskPanel.DragOver += OnTaskDragOver;
        taskPanel.DragLeave += (s, e) => { dragInsertIndex = -1; taskPanel.Invalidate(); };
        taskPanel.DragDrop += OnTaskDragDrop;
        taskPanel.Paint += OnTaskContainerPaint;

        taskPanel.Resize += (s, e) => {
            int safeWidth = taskPanel.ClientSize.Width - (int)(15 * scale);
            if (safeWidth > 0) {
                foreach (Control c in taskPanel.Controls) if (c is Panel) c.Width = safeWidth;
            }
        };

        this.Controls.Add(taskPanel);
        taskPanel.BringToFront();

        LoadShortcutsFromDb();
    }

    public void RefreshUI() {
        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > (int)(50 * scale) ? taskPanel.ClientSize.Width - (int)(15 * scale) : (int)(450 * scale);

        foreach (var s in shortcuts) {
            Panel card = new Panel() { 
                Width = startWidth, 
                AutoSize = true, 
                Margin = new Padding(0, 0, 0, (int)(10 * scale)), 
                BackColor = UITheme.CardWhite,
                Tag = s // 綁定資料庫 ID 與物件
            };

            // 自訂 iOS 圓角繪製
            card.Paint += (sender, e) => {
                UITheme.DrawRoundedBackground(e.Graphics, new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale), UITheme.CardWhite);
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var pen = new Pen(Color.FromArgb(230, 230, 230), 1)) {
                    e.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale)));
                }
            };

            TableLayoutPanel tlp = new TableLayoutPanel() { 
                Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, AutoSize = true, 
                Padding = new Padding((int)(8 * scale)),
                BackColor = Color.Transparent
            };
            
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(60 * scale))); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(38 * scale))); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(38 * scale))); 

            Button btnOpen = new Button() { Text = "開啟", Dock = DockStyle.Top, Height = (int)(32 * scale), BackColor = UITheme.AppleGreen, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0, 0, (int)(5 * scale), 0), Font = UITheme.GetFont(9f, FontStyle.Bold) };
            btnOpen.FlatAppearance.BorderSize = 0; 
            btnOpen.Click += (sender, e) => {
                try { 
                    ProcessStartInfo psi = new ProcessStartInfo() { FileName = s.Path, UseShellExecute = true };
                    Process.Start(psi); 
                } 
                catch { MessageBox.Show("無法開啟此捷徑，請檢查路徑或檔案是否存在！", "開啟失敗", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            };

            Button btnDel = new Button() { Text = "✕", Dock = DockStyle.Top, Height = (int)(32 * scale), BackColor = UITheme.AppleRed, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding(0, 0, (int)(5 * scale), 0), Font = UITheme.GetFont(9f, FontStyle.Bold) };
            btnDel.FlatAppearance.BorderSize = 0; 
            btnDel.Click += (sender, e) => { 
                if (MessageBox.Show("確定移除捷徑？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK) {
                    DeleteShortcutDb(s.Id);
                    shortcuts.Remove(s);
                    RefreshUI();
                }
            };

            Label lbl = new Label() { 
                Text = s.Name, 
                Dock = DockStyle.Fill, 
                TextAlign = ContentAlignment.MiddleLeft, 
                AutoSize = true, 
                Font = UITheme.GetFont(10.5f), 
                ForeColor = UITheme.TextMain,
                Padding = new Padding((int)(5 * scale), (int)(5 * scale), 0, (int)(5 * scale)),
                Cursor = Cursors.SizeAll 
            };
            lbl.MouseDown += (sender, e) => {
                if (e.Button == MouseButtons.Left) card.DoDragDrop(card, DragDropEffects.Move);
            };

            Button btnEdit = new Button() { Text = "修", Dock = DockStyle.Top, Height = (int)(32 * scale), BackColor = UITheme.BgGray, ForeColor = UITheme.TextMain, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, Margin = new Padding((int)(5 * scale), 0, 0, 0), Font = UITheme.GetFont(9f, FontStyle.Bold) };
            btnEdit.FlatAppearance.BorderSize = 0; 
            btnEdit.Click += (sender, e) => {
                new EditShortcutWindow(this, s).ShowDialog();
            };

            tlp.Controls.Add(btnOpen, 0, 0);
            tlp.Controls.Add(btnDel, 1, 0);
            tlp.Controls.Add(lbl, 2, 0);
            tlp.Controls.Add(btnEdit, 3, 0);

            card.Controls.Add(tlp);
            taskPanel.Controls.Add(card);
        }
    }

    // --- 拖曳排序邏輯 ---
    private void OnTaskDragOver(object sender, DragEventArgs e) {
        e.Effect = DragDropEffects.Move;
        Point clientPoint = taskPanel.PointToClient(new Point(e.X, e.Y));
        Control target = taskPanel.GetChildAtPoint(clientPoint);
        if (target != null) {
            int idx = taskPanel.Controls.GetChildIndex(target);
            if (clientPoint.Y > target.Top + (target.Height / 2)) idx++;
            if (dragInsertIndex != idx) { dragInsertIndex = idx; taskPanel.Invalidate(); }
        }
    }

    private void OnTaskContainerPaint(object sender, PaintEventArgs e) {
        if (dragInsertIndex != -1 && taskPanel.Controls.Count > 0) {
            int y = (dragInsertIndex < taskPanel.Controls.Count) ? taskPanel.Controls[dragInsertIndex].Top - 2 : taskPanel.Controls[taskPanel.Controls.Count - 1].Bottom + 2;
            e.Graphics.FillRectangle(new SolidBrush(UITheme.AppleBlue), 5, y, taskPanel.Width - 30, 3);
        }
    }

    private void OnTaskDragDrop(object sender, DragEventArgs e) {
        Panel draggedCard = (Panel)e.Data.GetData(typeof(Panel));
        if (draggedCard != null && dragInsertIndex != -1) {
            int targetIdx = dragInsertIndex;
            int currentIdx = taskPanel.Controls.GetChildIndex(draggedCard);
            if (currentIdx < targetIdx) targetIdx--; 
            
            taskPanel.Controls.SetChildIndex(draggedCard, targetIdx);
            
            // 寫入資料庫新順序
            using (var conn = DbHelper.GetConnection()) {
                conn.Open();
                using (var transaction = conn.BeginTransaction()) {
                    for (int i = 0; i < taskPanel.Controls.Count; i++) {
                        Panel card = (Panel)taskPanel.Controls[i];
                        ShortcutItem item = (ShortcutItem)card.Tag;
                        item.OrderIndex = i;
                        
                        using (var cmd = new SqliteCommand("UPDATE Shortcuts SET OrderIndex=@Order WHERE Id=@Id", conn, transaction)) {
                            cmd.Parameters.AddWithValue("@Order", item.OrderIndex);
                            cmd.Parameters.AddWithValue("@Id", item.Id);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }
            }
            shortcuts = shortcuts.OrderBy(s => s.OrderIndex).ToList();
            
            dragInsertIndex = -1; 
            taskPanel.Invalidate(); 
        }
    }

    // --- 資料庫操作 ---
    public void LoadShortcutsFromDb() {
        shortcuts.Clear();
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "SELECT Id, Name, Path, OrderIndex FROM Shortcuts ORDER BY OrderIndex ASC";
            using (var cmd = new SqliteCommand(sql, conn)) {
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) {
                        shortcuts.Add(new ShortcutItem {
                            Id = reader.GetInt32(0),
                            Name = reader.GetString(1),
                            Path = reader.GetString(2),
                            OrderIndex = reader.GetInt32(3)
                        });
                    }
                }
            }
        }
        RefreshUI();
    }

    public void AddShortcutDb(string name, string path) {
        int orderIdx = shortcuts.Count > 0 ? shortcuts.Max(s => s.OrderIndex) + 1 : 0;
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "INSERT INTO Shortcuts (Name, Path, OrderIndex) VALUES (@Name, @Path, @Order)";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@Name", name);
                cmd.Parameters.AddWithValue("@Path", path);
                cmd.Parameters.AddWithValue("@Order", orderIdx);
                cmd.ExecuteNonQuery();
            }
        }
        LoadShortcutsFromDb();
    }

    public void UpdateShortcutDb(int id, string name, string path) {
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "UPDATE Shortcuts SET Name=@Name, Path=@Path WHERE Id=@Id";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@Name", name);
                cmd.Parameters.AddWithValue("@Path", path);
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.ExecuteNonQuery();
            }
        }
        LoadShortcutsFromDb();
    }

    public void DeleteShortcutDb(int id) {
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            using (var cmd = new SqliteCommand("DELETE FROM Shortcuts WHERE Id=@Id", conn)) {
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.ExecuteNonQuery();
            }
        }
    }
}

// ==========================================
// 視窗：新增/編輯捷徑 (UI/DPI 升級)
// ==========================================
public class EditShortcutWindow : Form {
    private App_Shortcuts parent;
    private App_Shortcuts.ShortcutItem currentItem;
    private TextBox txtName, txtPath;

    public EditShortcutWindow(App_Shortcuts p, App_Shortcuts.ShortcutItem item) {
        this.parent = p; 
        this.currentItem = item;
        float scale = this.DeviceDpi / 96f;

        this.Text = item == null ? "新增捷徑" : "編輯捷徑";
        this.Width = (int)(420 * scale); 
        this.Height = (int)(280 * scale); 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false; 
        this.MinimizeBox = false;
        this.BackColor = UITheme.BgGray;

        FlowLayoutPanel f = new FlowLayoutPanel() { 
            Dock = DockStyle.Fill, 
            FlowDirection = FlowDirection.TopDown, 
            Padding = new Padding((int)(20 * scale)) 
        };
        
        f.Controls.Add(new Label() { Text = "捷徑名稱：", AutoSize = true, Margin = new Padding(0, 0, 0, (int)(5 * scale)), Font = UITheme.GetFont(10.5f, FontStyle.Bold) });
        txtName = new TextBox() { Width = (int)(360 * scale), Text = item?.Name ?? "", Margin = new Padding(0, 0, 0, (int)(15 * scale)), Font = UITheme.GetFont(10.5f) };
        f.Controls.Add(txtName);

        f.Controls.Add(new Label() { Text = "目標路徑 (檔案 / 資料夾 / 網址)：", AutoSize = true, Margin = new Padding(0, 0, 0, (int)(5 * scale)), Font = UITheme.GetFont(10.5f, FontStyle.Bold) });
        
        TableLayoutPanel pathRow = new TableLayoutPanel() { 
            Width = (int)(360 * scale), Height = (int)(40 * scale), 
            ColumnCount = 2, Margin = new Padding(0, 0, 0, (int)(20 * scale)) 
        };
        pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        pathRow.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(80 * scale)));
        
        txtPath = new TextBox() { Dock = DockStyle.Fill, Text = item?.Path ?? "", Font = UITheme.GetFont(10.5f), Margin = new Padding(0, (int)(5 * scale), (int)(5 * scale), 0) };
        
        Button btnBrowse = new Button() { 
            Text = "瀏覽", Dock = DockStyle.Fill, 
            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, 
            BackColor = UITheme.CardWhite, Font = UITheme.GetFont(10f) 
        };
        btnBrowse.FlatAppearance.BorderColor = Color.LightGray;
        btnBrowse.Click += (s, e) => {
            OpenFileDialog ofd = new OpenFileDialog() { Title = "選擇捷徑目標檔案" };
            if (ofd.ShowDialog() == DialogResult.OK) { txtPath.Text = ofd.FileName; }
        };
        
        pathRow.Controls.Add(txtPath, 0, 0);
        pathRow.Controls.Add(btnBrowse, 1, 0);
        f.Controls.Add(pathRow);

        Button btnSave = new Button() { 
            Text = "儲存設定", Width = (int)(360 * scale), Height = (int)(40 * scale), 
            BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, 
            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, 
            Font = UITheme.GetFont(11f, FontStyle.Bold) 
        };
        btnSave.FlatAppearance.BorderSize = 0;
        btnSave.Click += (s, e) => {
            if (string.IsNullOrWhiteSpace(txtName.Text) || string.IsNullOrWhiteSpace(txtPath.Text)) {
                MessageBox.Show("名稱與路徑不可為空！"); return;
            }
            if (currentItem == null) {
                parent.AddShortcutDb(txtName.Text, txtPath.Text);
            } else {
                parent.UpdateShortcutDb(currentItem.Id, txtName.Text, txtPath.Text);
            }
            this.Close();
        };
        f.Controls.Add(btnSave);

        this.Controls.Add(f);
    }
}
