using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Linq;
using System.Threading;
using Microsoft.Data.Sqlite;

public class App_TodoList : UserControl {
    // 【修改】：支援轉移到多個清單
    public Dictionary<string, App_TodoList> TargetLists = new Dictionary<string, App_TodoList>();
    
    private string listType; // "todo" 或 "plan" 等
    private string titleName; // 用於列印顯示的清單名稱

    private TextBox inputField;
    private FlowLayoutPanel taskContainer;

    public class TaskInfo {
        public int Id;
        public string Text;
        public string Color;
        public string Note;
        public DateTime Time;
        public int OrderIndex;
    }
    
    // UI 中的卡片集合
    private List<TaskInfo> taskDataList = new List<TaskInfo>();
    private int dragInsertIndex = -1; 
    private MainForm mainForm;
    private float scale;

    private readonly string[] colorCycle = { "Black", "Red", "DodgerBlue", "MediumOrchid", "DarkGreen", "DarkOrange" };

    public App_TodoList(MainForm parent, string type, string title) {
        this.mainForm = parent; 
        this.listType = type;
        this.titleName = title;
        this.Dock = DockStyle.Fill; 
        this.scale = this.DeviceDpi / 96f;

        this.BackColor = UITheme.BgGray;
        this.Padding = new Padding((int)(10 * scale));

        // --- 頂部標題與列印按鈕區 (新增需求) ---
        TableLayoutPanel header = new TableLayoutPanel() { 
            Dock = DockStyle.Top, Height = (int)(40 * scale), ColumnCount = 2 
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(110 * scale)));

        Label lblTitle = new Label() { 
            Text = titleName, Font = UITheme.GetFont(12f, FontStyle.Bold), 
            Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, 
            ForeColor = UITheme.TextMain 
        };
        
        Button btnPrint = new Button() { 
            Text = "列印/PDF", Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat, 
            Margin = new Padding((int)(2 * scale), (int)(5 * scale), (int)(2 * scale), (int)(5 * scale)), 
            Cursor = Cursors.Hand, BackColor = UITheme.AppleGreen, ForeColor = UITheme.CardWhite,
            Font = UITheme.GetFont(9f, FontStyle.Bold)
        };
        btnPrint.FlatAppearance.BorderSize = 0;
        btnPrint.Click += (s, e) => ExecutePrint();

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnPrint, 1, 0);

        // --- 頂部輸入區 ---
        TableLayoutPanel top = new TableLayoutPanel();
        top.Dock = DockStyle.Top;
        top.Height = (int)(45 * scale);
        top.ColumnCount = 2;
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        top.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(85 * scale))); 

        inputField = new TextBox();
        inputField.Dock = DockStyle.Fill;
        inputField.Font = UITheme.GetFont(11f);
        inputField.Margin = new Padding(0, (int)(8 * scale), (int)(8 * scale), 0);
        inputField.KeyDown += new KeyEventHandler(InputField_KeyDown);
        
        Button btnAdd = new Button();
        btnAdd.Text = "新增";
        btnAdd.Dock = DockStyle.Fill;
        btnAdd.FlatStyle = FlatStyle.Flat;
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.BackColor = UITheme.AppleBlue;
        btnAdd.ForeColor = UITheme.CardWhite;
        btnAdd.Font = UITheme.GetFont(11f, FontStyle.Bold);
        btnAdd.Margin = new Padding(0, (int)(5 * scale), 0, (int)(5 * scale));
        btnAdd.Cursor = Cursors.Hand;
        btnAdd.Click += new EventHandler(BtnAdd_Click);

        top.Controls.Add(inputField, 0, 0);
        top.Controls.Add(btnAdd, 1, 0);

        // --- 任務清單容器 ---
        taskContainer = new FlowLayoutPanel();
        taskContainer.Dock = DockStyle.Fill;
        taskContainer.AutoScroll = true;
        taskContainer.FlowDirection = FlowDirection.TopDown;
        taskContainer.WrapContents = false;
        taskContainer.BackColor = UITheme.BgGray;
        taskContainer.AllowDrop = true;
        
        taskContainer.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
        taskContainer.DragOver += OnTaskDragOver;
        taskContainer.DragLeave += (s, e) => { dragInsertIndex = -1; taskContainer.Invalidate(); };
        taskContainer.DragDrop += OnTaskDragDrop;
        taskContainer.Paint += OnTaskContainerPaint;

        taskContainer.Resize += (s, e) => {
            int safeWidth = taskContainer.ClientSize.Width - (int)(10 * scale); 
            if (safeWidth > 0) {
                foreach (Control c in taskContainer.Controls) {
                    if (c is Panel) c.Width = safeWidth;
                }
            }
        };

        this.Controls.Add(taskContainer);
        this.Controls.Add(top);
        this.Controls.Add(header); // 加入列印標題區塊
        taskContainer.BringToFront(); 
        
        LoadTasksFromDb();
    }

    private void InputField_KeyDown(object sender, KeyEventArgs e) {
        if (e.KeyCode == Keys.Enter) { 
            e.SuppressKeyPress = true; 
            AddTask(inputField.Text); 
            inputField.Text = ""; 
        }
    }

    private void BtnAdd_Click(object sender, EventArgs e) {
        AddTask(inputField.Text); 
        inputField.Text = "";
    }

    // --- 資料庫操作與任務新增 ---
    public void AddTask(string text, string colorName = "Black", string source = "手動", string note = "") {
        text = text.Trim(); 
        if (string.IsNullOrEmpty(text)) return;
        
        // 【新增需求】：手動新增時，自動在文字最前方加上年月日 yy/M/d
        if (source == "手動") {
            string datePrefix = DateTime.Now.ToString("yy/M/d") + " ";
            if (!text.StartsWith(datePrefix)) {
                text = datePrefix + text;
            }
        }

        DateTime now = DateTime.Now;
        int orderIdx = taskDataList.Count > 0 ? taskDataList.Min(t => t.OrderIndex) - 1 : 0;

        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = @"
                INSERT INTO Tasks (ListType, Text, Color, Note, CreatedTime, OrderIndex) 
                VALUES (@Type, @Text, @Color, @Note, @Time, @Order);
                SELECT last_insert_rowid();";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@Type", listType);
                cmd.Parameters.AddWithValue("@Text", text);
                cmd.Parameters.AddWithValue("@Color", colorName);
                cmd.Parameters.AddWithValue("@Note", note);
                cmd.Parameters.AddWithValue("@Time", now);
                cmd.Parameters.AddWithValue("@Order", orderIdx);
                
                int newId = Convert.ToInt32(cmd.ExecuteScalar());
                
                var newTask = new TaskInfo { Id = newId, Text = text, Color = colorName, Note = note, Time = now, OrderIndex = orderIdx };
                taskDataList.Insert(0, newTask);
                
                CreateTaskUICard(newTask, true); // Insert at top
            }
        }
    }

    private void LoadTasksFromDb() {
        taskContainer.Controls.Clear();
        taskDataList.Clear();

        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "SELECT Id, Text, Color, Note, CreatedTime, OrderIndex FROM Tasks WHERE ListType = @Type ORDER BY OrderIndex ASC";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@Type", listType);
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) {
                        var t = new TaskInfo {
                            Id = reader.GetInt32(0),
                            Text = reader.GetString(1),
                            Color = reader.GetString(2),
                            Note = reader.IsDBNull(3) ? "" : reader.GetString(3),
                            Time = reader.GetDateTime(4),
                            OrderIndex = reader.GetInt32(5)
                        };
                        taskDataList.Add(t);
                    }
                }
            }
        }

        foreach (var task in taskDataList) {
            CreateTaskUICard(task, false);
        }
    }

    // --- iOS 風格卡片 UI 生成 ---
    private void CreateTaskUICard(TaskInfo task, bool insertAtTop) {
        Color textColor = Color.FromName(task.Color);
        int startWidth = taskContainer.ClientSize.Width > (int)(20 * scale) ? taskContainer.ClientSize.Width - (int)(10 * scale) : (int)(450 * scale);

        // 卡片容器 【修改需求】：卡片下邊距改為 3 * scale，讓排列變緊密
        Panel card = new Panel() {
            Width = startWidth,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, (int)(3 * scale)),
            BackColor = UITheme.CardWhite,
            Tag = task // 將資料綁定到 UI 上
        };

        // 自訂繪製卡片圓角與邊框
        card.Paint += (s, e) => {
            UITheme.DrawRoundedBackground(e.Graphics, new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale), UITheme.CardWhite);
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (var pen = new Pen(Color.FromArgb(230, 230, 230), 1)) {
                e.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale)));
            }
        };

        TableLayoutPanel item = new TableLayoutPanel() {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 6,
            RowCount = 1,
            Padding = new Padding((int)(5 * scale)),
            BackColor = Color.Transparent
        };

        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(35 * scale))); 
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(38 * scale))); // 註
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(38 * scale))); // 轉
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(38 * scale))); // 色
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(38 * scale))); // 修
        
        // 完成 Checkbox
        CheckBox chk = new CheckBox() {
            Dock = DockStyle.Fill,
            Cursor = Cursors.Hand,
            BackColor = Color.Transparent,
            CheckAlign = ContentAlignment.MiddleCenter,
            Padding = new Padding((int)(5 * scale), (int)(5 * scale), 0, 0)
        };
        
        chk.CheckedChanged += (s, e) => {
            if (chk.Checked) {
                // 從資料庫刪除
                using (var conn = DbHelper.GetConnection()) {
                    conn.Open();
                    using (var cmd = new SqliteCommand("DELETE FROM Tasks WHERE Id = @Id", conn)) {
                        cmd.Parameters.AddWithValue("@Id", task.Id);
                        cmd.ExecuteNonQuery();
                    }
                }
                taskDataList.Remove(task);
                taskContainer.Controls.Remove(card);
            }
        };

        // 任務文字
        Label lbl = new Label() {
            Text = task.Text,
            Dock = DockStyle.Fill,
            Font = UITheme.GetFont(10.5f),
            ForeColor = textColor,
            AutoSize = true,
            Padding = new Padding(0, (int)(6 * scale), 0, (int)(6 * scale)),
            Cursor = Cursors.SizeAll, // 提示可拖曳
            BackColor = Color.Transparent,
            TextAlign = ContentAlignment.MiddleLeft
        };

        // 註解按鈕
        Button btnNote = CreateCardButton("註");
        Action updateNoteStyle = () => {
            if (!string.IsNullOrEmpty(task.Note)) {
                btnNote.BackColor = UITheme.AppleYellow; 
                btnNote.ForeColor = UITheme.TextMain;
            } else {
                btnNote.BackColor = UITheme.BgGray; 
                btnNote.ForeColor = textColor; 
            }
        };
        updateNoteStyle();

        btnNote.Click += (s, e) => {
            string newNote = ShowNoteEditBox(task.Text, task.Note);
            if (newNote != null) {
                task.Note = newNote;
                UpdateTaskInDb(task);
                updateNoteStyle();
            }
        };

        // 【修改需求】：將原本固定的轉移按鈕，改為動態選單「轉」，可選目標清單
        Button btnMove = CreateCardButton("轉");
        btnMove.BackColor = Color.FromArgb(235, 240, 255);
        btnMove.ForeColor = UITheme.AppleBlue;
        btnMove.Click += (s, e) => {
            ContextMenuStrip menu = new ContextMenuStrip();
            foreach (var kvp in TargetLists) {
                string targetName = kvp.Key;
                App_TodoList targetApp = kvp.Value;
                menu.Items.Add($"轉至 {targetName}", null, (sender, ev) => {
                    targetApp.AddTask(task.Text, task.Color, "轉移寫入", task.Note);
                    chk.Checked = true; // 本清單觸發刪除邏輯
                });
            }
            if (menu.Items.Count > 0) {
                menu.Show(btnMove, new Point(0, btnMove.Height));
            }
        };

        // 顏色切換按鈕
        Button btnColor = CreateCardButton("色");
        btnColor.Click += (s, e) => {
            int nextIdx = (Array.IndexOf(colorCycle, task.Color) + 1) % colorCycle.Length;
            task.Color = colorCycle[nextIdx];
            Color newColor = Color.FromName(task.Color);
            
            lbl.ForeColor = newColor;
            btnColor.ForeColor = newColor;
            
            UpdateTaskInDb(task);
            updateNoteStyle();
        };

        // 編輯按鈕
        Button btnEdit = CreateCardButton("修");
        Action triggerEdit = () => {
            string newText = ShowLargeEditBox(task.Text); 
            if (!string.IsNullOrEmpty(newText) && newText != task.Text) {
                task.Text = newText;
                lbl.Text = newText;
                UpdateTaskInDb(task);
            }
        };

        lbl.MouseDoubleClick += (s, e) => triggerEdit();
        btnEdit.Click += (s, e) => triggerEdit();
        
        // 啟動拖曳
        lbl.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) card.DoDragDrop(card, DragDropEffects.Move); };

        item.Controls.Add(chk, 0, 0);
        item.Controls.Add(lbl, 1, 0);
        item.Controls.Add(btnNote, 2, 0);
        item.Controls.Add(btnMove, 3, 0);
        item.Controls.Add(btnColor, 4, 0);
        item.Controls.Add(btnEdit, 5, 0);
        
        card.Controls.Add(item); 
        
        if (insertAtTop) {
            taskContainer.Controls.Add(card); 
            taskContainer.Controls.SetChildIndex(card, 0);
        } else {
            taskContainer.Controls.Add(card);
        }
    }

    private Button CreateCardButton(string text) {
        Button btn = new Button() {
            Text = text,
            Dock = DockStyle.Fill,
            Height = (int)(32 * scale),
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            BackColor = UITheme.BgGray,
            Font = UITheme.GetFont(9f, FontStyle.Bold),
            Margin = new Padding((int)(3 * scale))
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    private void UpdateTaskInDb(TaskInfo task) {
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "UPDATE Tasks SET Text=@Text, Color=@Color, Note=@Note, OrderIndex=@Order WHERE Id=@Id";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@Text", task.Text);
                cmd.Parameters.AddWithValue("@Color", task.Color);
                cmd.Parameters.AddWithValue("@Note", task.Note);
                cmd.Parameters.AddWithValue("@Order", task.OrderIndex);
                cmd.Parameters.AddWithValue("@Id", task.Id);
                cmd.ExecuteNonQuery();
            }
        }
    }

    // --- 拖曳排序機制 ---
    private void OnTaskDragOver(object sender, DragEventArgs e) {
        e.Effect = DragDropEffects.Move;
        Point clientPoint = taskContainer.PointToClient(new Point(e.X, e.Y));
        Control target = taskContainer.GetChildAtPoint(clientPoint);
        if (target != null) {
            int idx = taskContainer.Controls.GetChildIndex(target);
            if (clientPoint.Y > target.Top + (target.Height / 2)) idx++;
            if (dragInsertIndex != idx) { dragInsertIndex = idx; taskContainer.Invalidate(); }
        }
    }

    private void OnTaskContainerPaint(object sender, PaintEventArgs e) {
        if (dragInsertIndex != -1 && taskContainer.Controls.Count > 0) {
            int y = (dragInsertIndex < taskContainer.Controls.Count) ? taskContainer.Controls[dragInsertIndex].Top - 2 : taskContainer.Controls[taskContainer.Controls.Count - 1].Bottom + 2;
            e.Graphics.FillRectangle(new SolidBrush(UITheme.AppleBlue), 5, y, taskContainer.Width - 30, 3);
        }
    }

    private void OnTaskDragDrop(object sender, DragEventArgs e) {
        Panel draggedCard = (Panel)e.Data.GetData(typeof(Panel));
        if (draggedCard != null && dragInsertIndex != -1) {
            int targetIdx = dragInsertIndex;
            int currentIdx = taskContainer.Controls.GetChildIndex(draggedCard);
            if (currentIdx < targetIdx) targetIdx--; 
            
            taskContainer.Controls.SetChildIndex(draggedCard, targetIdx);
            
            // 重新計算所有的 OrderIndex 並寫入資料庫
            using (var conn = DbHelper.GetConnection()) {
                conn.Open();
                using (var transaction = conn.BeginTransaction()) {
                    for (int i = 0; i < taskContainer.Controls.Count; i++) {
                        Panel card = (Panel)taskContainer.Controls[i];
                        TaskInfo task = (TaskInfo)card.Tag;
                        task.OrderIndex = i; // 畫面由上到下，數字由小到大
                        
                        using (var cmd = new SqliteCommand("UPDATE Tasks SET OrderIndex=@Order WHERE Id=@Id", conn, transaction)) {
                            cmd.Parameters.AddWithValue("@Order", task.OrderIndex);
                            cmd.Parameters.AddWithValue("@Id", task.Id);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }
            }
            // 同步記憶體內的 List 排序
            taskDataList = taskDataList.OrderBy(t => t.OrderIndex).ToList();
            
            dragInsertIndex = -1; 
            taskContainer.Invalidate(); 
        }
    }

    // --- 【新增需求】列印/轉存PDF功能 ---
    private void ExecutePrint() {
        PrintDocument pd = new PrintDocument();
        int currentLine = 0;
        Font titleFont = UITheme.GetFont(18f, FontStyle.Bold);
        Font txtFont = UITheme.GetFont(12f);
        Font noteFont = UITheme.GetFont(10f);

        pd.PrintPage += (sender, args) => {
            float yPos = args.MarginBounds.Top;
            float leftMargin = args.MarginBounds.Left;

            // 印標題
            if (currentLine == 0) {
                args.Graphics.DrawString($"【 {titleName} 】", titleFont, Brushes.Black, leftMargin, yPos);
                yPos += 40;
                args.Graphics.DrawString("產生時間: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), noteFont, Brushes.Gray, leftMargin, yPos);
                yPos += 30;
                args.Graphics.DrawLine(Pens.Black, leftMargin, yPos, args.MarginBounds.Right, yPos);
                yPos += 15;
            }

            while (currentLine < taskDataList.Count) {
                var t = taskDataList[currentLine];
                string mainTxt = "□ " + t.Text;
                SizeF size = args.Graphics.MeasureString(mainTxt, txtFont, args.MarginBounds.Width);
                
                if (yPos + size.Height > args.MarginBounds.Bottom) {
                    args.HasMorePages = true; 
                    return; 
                }

                args.Graphics.DrawString(mainTxt, txtFont, Brushes.Black, new RectangleF(leftMargin, yPos, args.MarginBounds.Width, size.Height));
                yPos += size.Height + 5; 

                if (!string.IsNullOrWhiteSpace(t.Note)) {
                    string noteTxt = "   備註: " + t.Note.Replace("\r\n", " ").Replace("\n", " ");
                    SizeF noteSize = args.Graphics.MeasureString(noteTxt, noteFont, args.MarginBounds.Width);
                    if (yPos + noteSize.Height > args.MarginBounds.Bottom) { 
                        args.HasMorePages = true; 
                        return; 
                    }
                    args.Graphics.DrawString(noteTxt, noteFont, Brushes.DimGray, new RectangleF(leftMargin, yPos, args.MarginBounds.Width, noteSize.Height));
                    yPos += noteSize.Height + 5;
                }
                
                yPos += 10;
                currentLine++;
            }
            args.HasMorePages = false;
        };

        PrintDialog pdlg = new PrintDialog();
        pdlg.Document = pd;
        pdlg.UseEXDialog = true; 

        if (pdlg.ShowDialog() == DialogResult.OK) {
            try {
                pd.Print();
                MessageBox.Show("列印 / 存檔指令已送出！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            } catch (Exception ex) {
                MessageBox.Show("列印失敗：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    // --- 對話框 UI (DPI 適應) ---
    private string ShowLargeEditBox(string defaultValue) {
        Form form = new Form() { 
            Width = (int)(450 * scale), Height = (int)(280 * scale), 
            Text = "修正任務內容", StartPosition = FormStartPosition.CenterScreen, 
            FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false,
            BackColor = UITheme.BgGray 
        };
        Label lbl = new Label() { Text = "請輸入修正後的內容：", Left = (int)(15 * scale), Top = (int)(15 * scale), AutoSize = true, Font = UITheme.GetFont(10.5f, FontStyle.Bold) };
        TextBox txt = new TextBox() { Left = (int)(15 * scale), Top = (int)(45 * scale), Width = (int)(405 * scale), Height = (int)(120 * scale), Multiline = true, ScrollBars = ScrollBars.Vertical, Font = UITheme.GetFont(11f), Text = defaultValue };
        Button btnOk = new Button() { Text = "確認修改", Left = (int)(300 * scale), Top = (int)(180 * scale), Width = (int)(120 * scale), Height = (int)(40 * scale), DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, Font = UITheme.GetFont(10f, FontStyle.Bold), Cursor = Cursors.Hand };
        btnOk.FlatAppearance.BorderSize = 0;

        form.Controls.AddRange(new Control[] { lbl, txt, btnOk });
        form.AcceptButton = btnOk;
        txt.SelectionStart = txt.Text.Length; 
        
        if (form.ShowDialog() == DialogResult.OK) return txt.Text.Trim();
        return "";
    }

    private string ShowNoteEditBox(string taskName, string currentNote) {
        Form form = new Form() { 
            Width = (int)(420 * scale), Height = (int)(380 * scale), 
            Text = "任務詳細說明 (註)", StartPosition = FormStartPosition.CenterScreen, 
            FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false,
            BackColor = UITheme.BgGray 
        };
        
        Label lbl = new Label() { Text = "【 " + taskName + " 】", Left = (int)(15 * scale), Top = (int)(15 * scale), Width = (int)(370 * scale), Height = (int)(45 * scale), Font = UITheme.GetFont(11f, FontStyle.Bold), ForeColor = UITheme.AppleBlue };
        TextBox txt = new TextBox() { Left = (int)(15 * scale), Top = (int)(65 * scale), Width = (int)(370 * scale), Height = (int)(200 * scale), Multiline = true, ScrollBars = ScrollBars.Vertical, Font = UITheme.GetFont(10.5f), Text = currentNote };
        Button btnOk = new Button() { Text = "儲存說明", Left = (int)(265 * scale), Top = (int)(280 * scale), Width = (int)(120 * scale), Height = (int)(40 * scale), DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, Font = UITheme.GetFont(10f, FontStyle.Bold), Cursor = Cursors.Hand };
        btnOk.FlatAppearance.BorderSize = 0;

        form.Controls.AddRange(new Control[] { lbl, txt, btnOk });
        form.AcceptButton = btnOk;
        txt.SelectionStart = txt.Text.Length; 
        
        if (form.ShowDialog() == DialogResult.OK) return txt.Text.Trim();
        return null;
    }
}
