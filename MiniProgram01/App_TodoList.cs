// ============================================================
// FILE: MiniProgram01/App_TodoList.cs 
// ============================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;
using System.Linq;
using System.Threading;
using Microsoft.Data.Sqlite;

public class App_TodoList : UserControl {
    public Dictionary<string, App_TodoList> TargetLists = new Dictionary<string, App_TodoList>();
    
    private string listType; 
    private string titleName; 

    private TextBox inputField;
    private CheckBox chkDate;
    private FlowLayoutPanel taskContainer;

    public class TaskInfo {
        public int Id;
        public string Text;
        public string Color;
        public string Note;
        public DateTime Time;
        public int OrderIndex;
        public string DueDate; 
    }
    
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

        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            try { 
                using (var cmd = new SqliteCommand("ALTER TABLE Tasks ADD COLUMN DueDate TEXT", conn)) {
                    cmd.ExecuteNonQuery();
                }
            } catch { }
        }

        TableLayoutPanel topBar = new TableLayoutPanel();
        topBar.Dock = DockStyle.Top;
        topBar.Height = (int)(90 * scale); 
        topBar.RowCount = 2;
        topBar.ColumnCount = 4;
        topBar.Padding = new Padding(0);
        
        topBar.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
        topBar.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));

        topBar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(60 * scale))); 
        topBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));               
        topBar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(90 * scale))); 
        topBar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(80 * scale))); 

        Label lblTitle = new Label();
        lblTitle.Text = "項目：";
        lblTitle.Dock = DockStyle.Fill;
        lblTitle.TextAlign = ContentAlignment.MiddleRight;
        lblTitle.Font = UITheme.GetFont(11f, FontStyle.Bold);
        lblTitle.ForeColor = UITheme.TextMain;
        
        inputField = new TextBox();
        inputField.Dock = DockStyle.Fill;
        inputField.Font = UITheme.GetFont(11f);
        inputField.Margin = new Padding(0, (int)(8 * scale), (int)(5 * scale), 0);
        inputField.KeyDown += new KeyEventHandler(InputField_KeyDown);

        chkDate = new CheckBox();
        chkDate.Text = "設日期";
        chkDate.Dock = DockStyle.Fill;
        chkDate.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        chkDate.ForeColor = UITheme.TextMain;
        chkDate.Margin = new Padding(0, (int)(8 * scale), 0, 0);
        chkDate.Cursor = Cursors.Hand;

        Button btnAdd = new Button();
        btnAdd.Text = "新增";
        btnAdd.Dock = DockStyle.Fill;
        btnAdd.FlatStyle = FlatStyle.Flat;
        btnAdd.BackColor = UITheme.AppleBlue;
        btnAdd.ForeColor = UITheme.CardWhite;
        btnAdd.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        btnAdd.Margin = new Padding(0, (int)(5 * scale), (int)(5 * scale), (int)(5 * scale));
        btnAdd.Cursor = Cursors.Hand;
        btnAdd.FlatAppearance.BorderSize = 0;
        btnAdd.Click += new EventHandler(BtnAdd_Click);

        FlowLayoutPanel row2Panel = new FlowLayoutPanel();
        row2Panel.Dock = DockStyle.Fill;
        row2Panel.FlowDirection = FlowDirection.LeftToRight;
        row2Panel.Margin = new Padding(0);
        row2Panel.Padding = new Padding(0, (int)(5 * scale), 0, 0);

        Button btnCalendar = new Button();
        btnCalendar.Text = "日曆總覽";
        btnCalendar.Width = (int)(100 * scale);
        btnCalendar.Height = (int)(32 * scale);
        btnCalendar.FlatStyle = FlatStyle.Flat;
        btnCalendar.BackColor = UITheme.AppleYellow;
        btnCalendar.ForeColor = Color.Black;
        btnCalendar.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnCalendar.Margin = new Padding(0, 0, (int)(10 * scale), 0);
        btnCalendar.Cursor = Cursors.Hand;
        btnCalendar.FlatAppearance.BorderSize = 0;
        btnCalendar.Click += (s, e) => {
            TaskCalendarWindow calWin = new TaskCalendarWindow(this.listType);
            calWin.Show();
        };

        Button btnPrint = new Button();
        btnPrint.Text = "導出PDF";
        btnPrint.Width = (int)(90 * scale);
        btnPrint.Height = (int)(32 * scale);
        btnPrint.FlatStyle = FlatStyle.Flat;
        btnPrint.BackColor = UITheme.AppleGreen;
        btnPrint.ForeColor = UITheme.CardWhite;
        btnPrint.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnPrint.Margin = new Padding(0);
        btnPrint.Cursor = Cursors.Hand;
        btnPrint.FlatAppearance.BorderSize = 0;
        btnPrint.Click += (s, e) => ExecuteExportPDF();

        row2Panel.Controls.Add(btnCalendar);
        row2Panel.Controls.Add(btnPrint);

        topBar.Controls.Add(lblTitle, 0, 0);
        topBar.Controls.Add(inputField, 1, 0);
        topBar.Controls.Add(chkDate, 2, 0);
        topBar.Controls.Add(btnAdd, 3, 0);
        
        topBar.SetColumnSpan(row2Panel, 3);
        topBar.Controls.Add(row2Panel, 1, 1);

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
                taskContainer.SuspendLayout();
                foreach (Control c in taskContainer.Controls) {
                    if (c is Panel) c.Width = safeWidth;
                }
                taskContainer.ResumeLayout(true);
            }
        };

        this.Controls.Add(taskContainer);
        this.Controls.Add(topBar); 
        taskContainer.BringToFront(); 
        
        LoadTasksFromDb();
    }

    private void InputField_KeyDown(object sender, KeyEventArgs e) {
        if (e.KeyCode == Keys.Enter) { 
            e.SuppressKeyPress = true; 
            ExecuteAddLogic();
        }
    }

    private void BtnAdd_Click(object sender, EventArgs e) {
        ExecuteAddLogic();
    }

    private void ExecuteAddLogic() {
        string text = inputField.Text.Trim();
        if (string.IsNullOrEmpty(text)) return;

        string dueDateStr = "";
        string noteAddon = "";

        if (chkDate.Checked) {
            DateTimePickerDialog dialog = new DateTimePickerDialog(scale);
            if (dialog.ShowDialog() == DialogResult.OK) {
                DateTime selectedDt = dialog.SelectedDateTime;
                dueDateStr = selectedDt.ToString("yyyy-MM-dd HH:mm");
                noteAddon = $"期程：{selectedDt.ToString("yyyy年MM月dd日-時間HH:mm")}";
            } else {
                return; 
            }
        }

        AddTask(text, "Black", "手動", "", dueDateStr, noteAddon);
        inputField.Text = "";
        chkDate.Checked = false;
    }

    public void AddTask(string text, string colorName = "Black", string source = "手動", string note = "", string dueDateStr = "", string noteAddon = "") {
        text = text.Trim(); 
        if (string.IsNullOrEmpty(text)) return;
        
        if (source == "手動") {
            string dateNote = $"本項目於：{DateTime.Now:yyyy年MM月dd日} 新增";
            if (!string.IsNullOrEmpty(noteAddon)) {
                dateNote += "\r\n" + noteAddon;
            }

            if (string.IsNullOrEmpty(note)) {
                note = dateNote;
            } else {
                note = dateNote + "\r\n" + note;
            }
        }

        DateTime now = DateTime.Now;
        int orderIdx = 0;
        if (taskDataList.Count > 0) {
            orderIdx = taskDataList.Min(t => t.OrderIndex) - 1;
        }

        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = @"
                INSERT INTO Tasks (ListType, Text, Color, Note, CreatedTime, OrderIndex, DueDate) 
                VALUES (@Type, @Text, @Color, @Note, @Time, @Order, @Due);
                SELECT last_insert_rowid();";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@Type", listType);
                cmd.Parameters.AddWithValue("@Text", text);
                cmd.Parameters.AddWithValue("@Color", colorName);
                cmd.Parameters.AddWithValue("@Note", note);
                cmd.Parameters.AddWithValue("@Time", now);
                cmd.Parameters.AddWithValue("@Order", orderIdx);
                cmd.Parameters.AddWithValue("@Due", dueDateStr ?? "");
                
                int newId = Convert.ToInt32(cmd.ExecuteScalar());
                
                var newTask = new TaskInfo();
                newTask.Id = newId;
                newTask.Text = text;
                newTask.Color = colorName;
                newTask.Note = note;
                newTask.Time = now;
                newTask.OrderIndex = orderIdx;
                newTask.DueDate = dueDateStr ?? "";

                taskDataList.Insert(0, newTask);
                CreateTaskUICard(newTask, true); 
            }
        }
    }

    public void LoadTasksFromDb() {
        taskContainer.Controls.Clear();
        taskDataList.Clear();

        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "SELECT Id, Text, Color, Note, CreatedTime, OrderIndex, DueDate FROM Tasks WHERE ListType = @Type ORDER BY OrderIndex ASC";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@Type", listType);
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) {
                        var t = new TaskInfo();
                        t.Id = reader.GetInt32(0);
                        t.Text = reader.GetString(1);
                        t.Color = reader.GetString(2);
                        t.Note = reader.IsDBNull(3) ? "" : reader.GetString(3);
                        t.Time = reader.GetDateTime(4);
                        t.OrderIndex = reader.GetInt32(5);
                        t.DueDate = reader.IsDBNull(6) ? "" : reader.GetString(6);
                        
                        taskDataList.Add(t);
                    }
                }
            }
        }

        foreach (var task in taskDataList) {
            CreateTaskUICard(task, false);
        }
    }

    private void CreateTaskUICard(TaskInfo task, bool insertAtTop) {
        Color textColor = Color.FromName(task.Color);
        int startWidth = taskContainer.ClientSize.Width > (int)(20 * scale) ? taskContainer.ClientSize.Width - (int)(10 * scale) : (int)(450 * scale);

        Panel card = new Panel();
        card.Width = startWidth;
        card.AutoSize = true;
        card.Margin = new Padding(0, 0, 0, (int)(3 * scale));
        card.BackColor = UITheme.CardWhite;
        card.Tag = task;

        card.Paint += (s, e) => {
            UITheme.DrawRoundedBackground(e.Graphics, new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale), UITheme.CardWhite);
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (var pen = new Pen(Color.FromArgb(230, 230, 230), 1)) {
                e.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale)));
            }
        };

        TableLayoutPanel item = new TableLayoutPanel();
        item.Dock = DockStyle.Fill;
        item.AutoSize = true;
        item.ColumnCount = 6;
        item.RowCount = 1;
        item.Padding = new Padding((int)(5 * scale));
        item.BackColor = Color.Transparent;

        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(35 * scale))); 
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(32 * scale))); // 註
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(32 * scale))); // 轉
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(32 * scale))); // 色
        item.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(32 * scale))); // 修
        
        CheckBox chk = new CheckBox();
        chk.Dock = DockStyle.Fill;
        chk.Cursor = Cursors.Hand;
        chk.BackColor = Color.Transparent;
        chk.CheckAlign = ContentAlignment.MiddleCenter;
        chk.Padding = new Padding((int)(5 * scale), (int)(5 * scale), 0, 0);
        
        chk.CheckedChanged += (s, e) => {
            if (chk.Checked) {
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

        string displayTxt = task.Text;
        if (!string.IsNullOrEmpty(task.DueDate)) {
            displayTxt = $"[期] {task.Text}";
        }

        Label lbl = new Label();
        lbl.Text = displayTxt;
        lbl.Dock = DockStyle.Fill;
        lbl.Font = UITheme.GetFont(10.5f);
        lbl.ForeColor = textColor;
        lbl.AutoSize = true;
        lbl.Padding = new Padding(0, (int)(6 * scale), 0, (int)(6 * scale));
        lbl.Cursor = Cursors.SizeAll;
        lbl.BackColor = Color.Transparent;
        lbl.TextAlign = ContentAlignment.MiddleLeft;

        Button btnNote = CreateCardButton("註");
        Action updateNoteStyle = () => {
            bool isOnlySystemNote = true;
            if (!string.IsNullOrEmpty(task.Note)) {
                string[] lines = task.Note.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string line in lines) {
                    string tStr = line.Trim();
                    if (tStr.StartsWith("本項目於：") && tStr.EndsWith("新增")) continue;
                    if (tStr.StartsWith("期程：")) continue;
                    isOnlySystemNote = false;
                    break;
                }
            }

            if (!string.IsNullOrEmpty(task.Note) && !isOnlySystemNote) {
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

        Button btnMove = CreateCardButton("轉");
        btnMove.BackColor = Color.FromArgb(235, 240, 255);
        btnMove.ForeColor = UITheme.AppleBlue;
        btnMove.Click += (s, e) => {
            ContextMenuStrip menu = new ContextMenuStrip();
            foreach (var kvp in TargetLists) {
                string targetName = kvp.Key;
                App_TodoList targetApp = kvp.Value;
                menu.Items.Add($"轉至 {targetName}", null, (sender, ev) => {
                    targetApp.AddTask(task.Text, task.Color, "轉移寫入", task.Note, task.DueDate, "");
                    chk.Checked = true; 
                });
            }
            if (menu.Items.Count > 0) {
                menu.Show(btnMove, new Point(0, btnMove.Height));
            }
        };

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

        Button btnEdit = CreateCardButton("修");
        Action triggerEdit = () => {
            string newText = ShowLargeEditBox(task.Text); 
            if (!string.IsNullOrEmpty(newText) && newText != task.Text) {
                task.Text = newText; 
                lbl.Text = string.IsNullOrEmpty(task.DueDate) ? newText : $"[期] {newText}"; 
                UpdateTaskInDb(task);
            }
        };
        lbl.MouseDoubleClick += (s, e) => triggerEdit();
        btnEdit.Click += (s, e) => triggerEdit();
        
        lbl.MouseDown += (s, e) => { 
            if (e.Button == MouseButtons.Left) {
                card.DoDragDrop(card, DragDropEffects.Move); 
            }
        };

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
        Button btn = new Button();
        btn.Text = text;
        btn.Dock = DockStyle.Fill;
        btn.Height = (int)(32 * scale);
        btn.FlatStyle = FlatStyle.Flat;
        btn.Cursor = Cursors.Hand;
        btn.BackColor = UITheme.BgGray;
        btn.Font = UITheme.GetFont(9f, FontStyle.Bold);
        btn.Margin = new Padding((int)(2 * scale));
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    private void UpdateTaskInDb(TaskInfo task) {
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "UPDATE Tasks SET Text=@Text, Color=@Color, Note=@Note, OrderIndex=@Order, DueDate=@Due WHERE Id=@Id";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@Text", task.Text);
                cmd.Parameters.AddWithValue("@Color", task.Color);
                cmd.Parameters.AddWithValue("@Note", task.Note);
                cmd.Parameters.AddWithValue("@Order", task.OrderIndex);
                cmd.Parameters.AddWithValue("@Due", task.DueDate ?? "");
                cmd.Parameters.AddWithValue("@Id", task.Id);
                cmd.ExecuteNonQuery();
            }
        }
    }

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
            
            using (var conn = DbHelper.GetConnection()) {
                conn.Open();
                using (var transaction = conn.BeginTransaction()) {
                    for (int i = 0; i < taskContainer.Controls.Count; i++) {
                        Panel card = (Panel)taskContainer.Controls[i];
                        TaskInfo task = (TaskInfo)card.Tag;
                        task.OrderIndex = i; 
                        
                        using (var cmd = new SqliteCommand("UPDATE Tasks SET OrderIndex=@Order WHERE Id=@Id", conn, transaction)) {
                            cmd.Parameters.AddWithValue("@Order", task.OrderIndex);
                            cmd.Parameters.AddWithValue("@Id", task.Id);
                            cmd.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }
            }
            taskDataList = taskDataList.OrderBy(t => t.OrderIndex).ToList();
            dragInsertIndex = -1; 
            taskContainer.Invalidate(); 
        }
    }

    private void ExecuteExportPDF() {
        using (SaveFileDialog sfd = new SaveFileDialog()) {
            sfd.Filter = "PDF 檔案|*.pdf";
            sfd.FileName = $"{titleName}_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            
            if (sfd.ShowDialog() == DialogResult.OK) {
                using (PrintDocument pd = new PrintDocument()) {
                    pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                    pd.PrinterSettings.PrintToFile = true;
                    pd.PrinterSettings.PrintFileName = sfd.FileName;

                    int currentLine = 0;
                    Font titleFont = UITheme.GetFont(18f, FontStyle.Bold);
                    Font txtFont = UITheme.GetFont(12f);
                    Font noteFont = UITheme.GetFont(10f);

                    pd.PrintPage += (sender, args) => {
                        float yPos = args.MarginBounds.Top;
                        float leftMargin = args.MarginBounds.Left;

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
                            if(!string.IsNullOrEmpty(t.DueDate)) mainTxt = "□ [期] " + t.Text;

                            SizeF size = args.Graphics.MeasureString(mainTxt, txtFont, args.MarginBounds.Width);
                            
                            if (yPos + size.Height > args.MarginBounds.Bottom) {
                                args.HasMorePages = true; 
                                return; 
                            }

                            args.Graphics.DrawString(mainTxt, txtFont, Brushes.Black, new RectangleF(leftMargin, yPos, args.MarginBounds.Width, size.Height));
                            yPos += size.Height + 5; 

                            if (!string.IsNullOrWhiteSpace(t.Note)) {
                                string notePrefix = "   備註:\n      ";
                                string formattedNote = notePrefix + t.Note.Replace("\n", "\n      ");
                                
                                SizeF noteSize = args.Graphics.MeasureString(formattedNote, noteFont, args.MarginBounds.Width);
                                if (yPos + noteSize.Height > args.MarginBounds.Bottom) { 
                                    args.HasMorePages = true; 
                                    return; 
                                }
                                args.Graphics.DrawString(formattedNote, noteFont, Brushes.DimGray, new RectangleF(leftMargin, yPos, args.MarginBounds.Width, noteSize.Height));
                                yPos += noteSize.Height + 5;
                            }
                            
                            yPos += 10;
                            currentLine++;
                        }
                        args.HasMorePages = false;
                    };

                    try {
                        pd.Print();
                        MessageBox.Show("PDF 檔案已成功導出！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    } catch (Exception ex) {
                        MessageBox.Show("導出失敗！請確認安裝了Microsoft Print to PDF。\n" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                } 
            }
        }
    }

    private string ShowLargeEditBox(string defaultValue) {
        Form form = new Form();
        form.Width = (int)(450 * scale);
        form.Height = (int)(280 * scale);
        form.Text = "修正任務內容";
        form.StartPosition = FormStartPosition.CenterScreen;
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.MaximizeBox = false;
        form.MinimizeBox = false;
        form.TopMost = true; // 【修復】強制在最上層
        form.BackColor = UITheme.BgGray;

        Label lbl = new Label();
        lbl.Text = "請輸入修正後的內容：";
        lbl.Left = (int)(15 * scale);
        lbl.Top = (int)(15 * scale);
        lbl.AutoSize = true;
        lbl.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        
        TextBox txt = new TextBox();
        txt.Left = (int)(15 * scale);
        txt.Top = (int)(45 * scale);
        txt.Width = (int)(405 * scale);
        txt.Height = (int)(120 * scale);
        txt.Multiline = true;
        txt.AcceptsReturn = true;
        txt.ScrollBars = ScrollBars.Vertical;
        txt.Font = UITheme.GetFont(11f);
        txt.Text = defaultValue;
        
        Button btnOk = new Button();
        btnOk.Text = "確認修改";
        btnOk.Left = (int)(300 * scale);
        btnOk.Top = (int)(180 * scale);
        btnOk.Width = (int)(120 * scale);
        btnOk.Height = (int)(40 * scale);
        btnOk.DialogResult = DialogResult.OK;
        btnOk.FlatStyle = FlatStyle.Flat;
        btnOk.BackColor = UITheme.AppleBlue;
        btnOk.ForeColor = UITheme.CardWhite;
        btnOk.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnOk.Cursor = Cursors.Hand;
        btnOk.FlatAppearance.BorderSize = 0;

        form.Controls.Add(lbl);
        form.Controls.Add(txt);
        form.Controls.Add(btnOk);
        
        form.AcceptButton = btnOk;
        txt.SelectionStart = txt.Text.Length; 
        
        if (form.ShowDialog() == DialogResult.OK) {
            return txt.Text.Trim();
        }
        return "";
    }

    private string ShowNoteEditBox(string taskName, string currentNote) {
        Form form = new Form();
        form.Width = (int)(420 * scale);
        form.Height = (int)(380 * scale);
        form.Text = "任務詳細說明 (註)";
        form.StartPosition = FormStartPosition.CenterScreen;
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.MaximizeBox = false;
        form.MinimizeBox = false;
        form.TopMost = true; // 【修復】強制在最上層
        form.BackColor = UITheme.BgGray;
        
        Label lbl = new Label();
        lbl.Text = "【 " + taskName + " 】";
        lbl.Left = (int)(15 * scale);
        lbl.Top = (int)(15 * scale);
        lbl.Width = (int)(370 * scale);
        lbl.Height = (int)(45 * scale);
        lbl.Font = UITheme.GetFont(11f, FontStyle.Bold);
        lbl.ForeColor = UITheme.AppleBlue;
        
        TextBox txt = new TextBox();
        txt.Left = (int)(15 * scale);
        txt.Top = (int)(65 * scale);
        txt.Width = (int)(370 * scale);
        txt.Height = (int)(200 * scale);
        txt.Multiline = true;
        txt.AcceptsReturn = true;
        txt.ScrollBars = ScrollBars.Vertical;
        txt.Font = UITheme.GetFont(10.5f);
        txt.Text = currentNote;
        
        Button btnOk = new Button();
        btnOk.Text = "儲存說明";
        btnOk.Left = (int)(265 * scale);
        btnOk.Top = (int)(280 * scale);
        btnOk.Width = (int)(120 * scale);
        btnOk.Height = (int)(40 * scale);
        btnOk.DialogResult = DialogResult.OK;
        btnOk.FlatStyle = FlatStyle.Flat;
        btnOk.BackColor = UITheme.AppleBlue;
        btnOk.ForeColor = UITheme.CardWhite;
        btnOk.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnOk.Cursor = Cursors.Hand;
        btnOk.FlatAppearance.BorderSize = 0;

        form.Controls.Add(lbl);
        form.Controls.Add(txt);
        form.Controls.Add(btnOk);
        
        form.AcceptButton = btnOk;
        txt.SelectionStart = txt.Text.Length; 
        
        if (form.ShowDialog() == DialogResult.OK) {
            return txt.Text.Trim();
        }
        return null;
    }
}

// ============================================================
// 【新增】日期時間選擇視窗
// ============================================================
public class DateTimePickerDialog : Form {
    private DateTimePicker dpDate;
    private DateTimePicker dpTime;
    public DateTime SelectedDateTime { get; private set; }

    public DateTimePickerDialog(float scale) {
        this.Text = "設定期程";
        this.ClientSize = new Size((int)(300 * scale), (int)(220 * scale)); // 【修復】明確定義 ClientSize
        this.StartPosition = FormStartPosition.CenterScreen; // 【修復】強制螢幕置中
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.TopMost = true; // 【修復】永遠在最上層，避免被主視窗蓋住
        this.ShowInTaskbar = false; 
        this.BackColor = UITheme.BgGray;

        Label l1 = new Label();
        l1.Text = "請選擇日期：";
        l1.AutoSize = true;
        l1.Location = new Point((int)(20 * scale), (int)(20 * scale));
        l1.Font = UITheme.GetFont(10f, FontStyle.Bold);
        
        dpDate = new DateTimePicker();
        dpDate.Format = DateTimePickerFormat.Custom;
        dpDate.CustomFormat = "yyyy-MM-dd";
        dpDate.Location = new Point((int)(20 * scale), (int)(45 * scale));
        dpDate.Width = (int)(240 * scale);
        dpDate.Font = UITheme.GetFont(11f);

        Label l2 = new Label();
        l2.Text = "請選擇時間：";
        l2.AutoSize = true;
        l2.Location = new Point((int)(20 * scale), (int)(85 * scale));
        l2.Font = UITheme.GetFont(10f, FontStyle.Bold);

        dpTime = new DateTimePicker();
        dpTime.Format = DateTimePickerFormat.Custom;
        dpTime.CustomFormat = "HH:mm";
        dpTime.ShowUpDown = true;
        dpTime.Location = new Point((int)(20 * scale), (int)(110 * scale));
        dpTime.Width = (int)(240 * scale);
        dpTime.Font = UITheme.GetFont(11f);

        Button btnOk = new Button();
        btnOk.Text = "確認";
        btnOk.Location = new Point((int)(160 * scale), (int)(160 * scale));
        btnOk.Width = (int)(100 * scale);
        btnOk.Height = (int)(35 * scale);
        btnOk.DialogResult = DialogResult.OK;
        btnOk.FlatStyle = FlatStyle.Flat;
        btnOk.BackColor = UITheme.AppleBlue;
        btnOk.ForeColor = UITheme.CardWhite;
        btnOk.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnOk.FlatAppearance.BorderSize = 0;

        btnOk.Click += (s, e) => {
            SelectedDateTime = new DateTime(dpDate.Value.Year, dpDate.Value.Month, dpDate.Value.Day, dpTime.Value.Hour, dpTime.Value.Minute, 0);
        };

        this.Controls.Add(l1);
        this.Controls.Add(dpDate);
        this.Controls.Add(l2);
        this.Controls.Add(dpTime);
        this.Controls.Add(btnOk);
    }
}

// ============================================================
// 【新增】全螢幕日曆看板視窗
// ============================================================
public class TaskCalendarWindow : Form {
    private ComboBox cmbMode, cmbYear, cmbMonth;
    private TableLayoutPanel calendarGrid;
    private FlowLayoutPanel unassignedPanel;
    private float scale;

    public TaskCalendarWindow(string defaultType) {
        this.scale = this.DeviceDpi / 96f;
        this.Text = "日曆任務總覽";
        this.WindowState = FormWindowState.Maximized;
        this.TopMost = true; // 【修復】永遠在最上層，蓋過主視窗
        this.BackColor = UITheme.BgGray;

        // 頂部控制列
        Panel topPanel = new Panel();
        topPanel.Dock = DockStyle.Top;
        topPanel.Height = (int)(60 * scale);
        topPanel.BackColor = UITheme.CardWhite;
        
        Label l1 = new Label();
        l1.Text = "檢視模式：";
        l1.AutoSize = true;
        l1.Location = new Point((int)(20 * scale), (int)(20 * scale));
        l1.Font = UITheme.GetFont(11f, FontStyle.Bold);
        topPanel.Controls.Add(l1);

        cmbMode = new ComboBox();
        cmbMode.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbMode.Items.AddRange(new string[] { "總覽", "待辦", "待規", "行程" });
        cmbMode.Width = (int)(100 * scale);
        cmbMode.Location = new Point((int)(110 * scale), (int)(16 * scale));
        cmbMode.Font = UITheme.GetFont(11f);
        
        string mapType = "總覽";
        if (defaultType == "todo") mapType = "待辦";
        if (defaultType == "plan") mapType = "待規";
        if (defaultType == "schedule") mapType = "行程";
        cmbMode.Text = mapType;
        cmbMode.SelectedIndexChanged += (s, e) => RefreshData();
        topPanel.Controls.Add(cmbMode);

        Label l2 = new Label();
        l2.Text = "年份：";
        l2.AutoSize = true;
        l2.Location = new Point((int)(230 * scale), (int)(20 * scale));
        l2.Font = UITheme.GetFont(11f, FontStyle.Bold);
        topPanel.Controls.Add(l2);

        cmbYear = new ComboBox();
        cmbYear.DropDownStyle = ComboBoxStyle.DropDownList;
        int curYear = DateTime.Now.Year;
        for (int y = curYear - 2; y <= curYear + 5; y++) cmbYear.Items.Add(y.ToString());
        cmbYear.Text = curYear.ToString();
        cmbYear.Width = (int)(80 * scale);
        cmbYear.Location = new Point((int)(280 * scale), (int)(16 * scale));
        cmbYear.Font = UITheme.GetFont(11f);
        cmbYear.SelectedIndexChanged += (s, e) => RefreshData();
        topPanel.Controls.Add(cmbYear);

        Label l3 = new Label();
        l3.Text = "月份：";
        l3.AutoSize = true;
        l3.Location = new Point((int)(380 * scale), (int)(20 * scale));
        l3.Font = UITheme.GetFont(11f, FontStyle.Bold);
        topPanel.Controls.Add(l3);

        cmbMonth = new ComboBox();
        cmbMonth.DropDownStyle = ComboBoxStyle.DropDownList;
        for (int m = 1; m <= 12; m++) cmbMonth.Items.Add(m.ToString("D2"));
        cmbMonth.Text = DateTime.Now.Month.ToString("D2");
        cmbMonth.Width = (int)(60 * scale);
        cmbMonth.Location = new Point((int)(430 * scale), (int)(16 * scale));
        cmbMonth.Font = UITheme.GetFont(11f);
        cmbMonth.SelectedIndexChanged += (s, e) => RefreshData();
        topPanel.Controls.Add(cmbMonth);

        Button btnToday = new Button();
        btnToday.Text = "回到本月";
        btnToday.Width = (int)(100 * scale);
        btnToday.Height = (int)(32 * scale);
        btnToday.Location = new Point((int)(520 * scale), (int)(15 * scale));
        btnToday.BackColor = UITheme.AppleBlue;
        btnToday.ForeColor = UITheme.CardWhite;
        btnToday.FlatStyle = FlatStyle.Flat;
        btnToday.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnToday.FlatAppearance.BorderSize = 0;
        btnToday.Cursor = Cursors.Hand;
        btnToday.Click += (s, e) => {
            cmbYear.Text = DateTime.Now.Year.ToString();
            cmbMonth.Text = DateTime.Now.Month.ToString("D2");
        };
        topPanel.Controls.Add(btnToday);

        this.Controls.Add(topPanel);

        // 主畫面切割 (70% 日曆, 30% 未排定期程)
        TableLayoutPanel mainSplit = new TableLayoutPanel();
        mainSplit.Dock = DockStyle.Fill;
        mainSplit.ColumnCount = 2;
        mainSplit.RowCount = 1;
        mainSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70f));
        mainSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30f));
        mainSplit.Padding = new Padding((int)(10 * scale));

        // 左側：日曆區域
        TableLayoutPanel calWrapper = new TableLayoutPanel();
        calWrapper.Dock = DockStyle.Fill;
        calWrapper.RowCount = 2;
        calWrapper.ColumnCount = 1;
        calWrapper.RowStyles.Add(new RowStyle(SizeType.Absolute, (int)(40 * scale)));
        calWrapper.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        TableLayoutPanel daysHeader = new TableLayoutPanel();
        daysHeader.Dock = DockStyle.Fill;
        daysHeader.ColumnCount = 7;
        daysHeader.RowCount = 1;
        string[] wDays = { "日", "一", "二", "三", "四", "五", "六" };
        for (int i = 0; i < 7; i++) {
            daysHeader.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.28f));
            Label dLbl = new Label();
            dLbl.Text = wDays[i];
            dLbl.Dock = DockStyle.Fill;
            dLbl.TextAlign = ContentAlignment.MiddleCenter;
            dLbl.Font = UITheme.GetFont(11f, FontStyle.Bold);
            dLbl.ForeColor = (i == 0 || i == 6) ? UITheme.AppleRed : UITheme.TextMain;
            daysHeader.Controls.Add(dLbl, i, 0);
        }
        calWrapper.Controls.Add(daysHeader, 0, 0);

        calendarGrid = new TableLayoutPanel();
        calendarGrid.Dock = DockStyle.Fill;
        calendarGrid.ColumnCount = 7;
        calendarGrid.RowCount = 6;
        for (int i = 0; i < 7; i++) calendarGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 14.28f));
        for (int i = 0; i < 6; i++) calendarGrid.RowStyles.Add(new RowStyle(SizeType.Percent, 16.66f));
        calendarGrid.CellBorderStyle = TableLayoutPanelCellBorderStyle.Single;
        calendarGrid.BackColor = UITheme.CardWhite;
        calWrapper.Controls.Add(calendarGrid, 0, 1);

        mainSplit.Controls.Add(calWrapper, 0, 0);

        // 右側：未設定期程任務
        Panel rightPanel = new Panel();
        rightPanel.Dock = DockStyle.Fill;
        rightPanel.Padding = new Padding((int)(10 * scale), 0, 0, 0);
        
        Label rTitle = new Label();
        rTitle.Text = "未排定期程清單";
        rTitle.Dock = DockStyle.Top;
        rTitle.Height = (int)(40 * scale);
        rTitle.TextAlign = ContentAlignment.MiddleLeft;
        rTitle.Font = UITheme.GetFont(12f, FontStyle.Bold);
        rTitle.ForeColor = UITheme.AppleBlue;
        rightPanel.Controls.Add(rTitle);

        unassignedPanel = new FlowLayoutPanel();
        unassignedPanel.Dock = DockStyle.Fill;
        unassignedPanel.AutoScroll = true;
        unassignedPanel.FlowDirection = FlowDirection.TopDown;
        unassignedPanel.WrapContents = false;
        unassignedPanel.BackColor = UITheme.BgGray;
        rightPanel.Controls.Add(unassignedPanel);
        unassignedPanel.BringToFront();

        mainSplit.Controls.Add(rightPanel, 1, 0);
        this.Controls.Add(mainSplit);

        this.Load += (s, e) => RefreshData();
    }

    private string GetDbListType() {
        if (cmbMode.Text == "待辦") return "todo";
        if (cmbMode.Text == "待規") return "plan";
        if (cmbMode.Text == "行程") return "schedule";
        return "all";
    }

    public void RefreshData() {
        calendarGrid.SuspendLayout();
        unassignedPanel.SuspendLayout();
        calendarGrid.Controls.Clear();
        unassignedPanel.Controls.Clear();

        int targetYear = int.Parse(cmbYear.Text);
        int targetMonth = int.Parse(cmbMonth.Text);
        string typeFilter = GetDbListType();

        List<App_TodoList.TaskInfo> allTasks = new List<App_TodoList.TaskInfo>();
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "SELECT Id, Text, Color, Note, CreatedTime, OrderIndex, DueDate, ListType FROM Tasks";
            if (typeFilter != "all") sql += " WHERE ListType = @Type";
            
            using (var cmd = new SqliteCommand(sql, conn)) {
                if (typeFilter != "all") cmd.Parameters.AddWithValue("@Type", typeFilter);
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) {
                        var t = new App_TodoList.TaskInfo();
                        t.Id = reader.GetInt32(0);
                        t.Text = reader.GetString(1);
                        t.Color = reader.GetString(2);
                        t.Note = reader.IsDBNull(3) ? "" : reader.GetString(3);
                        t.Time = reader.GetDateTime(4);
                        t.OrderIndex = reader.GetInt32(5);
                        t.DueDate = reader.IsDBNull(6) ? "" : reader.GetString(6);
                        
                        string lType = reader.GetString(7);
                        string pfx = "";
                        if (typeFilter == "all") {
                            if (lType == "todo") pfx = "[待辦] ";
                            if (lType == "plan") pfx = "[待規] ";
                            if (lType == "schedule") pfx = "[行程] ";
                        }
                        t.Text = pfx + t.Text;
                        allTasks.Add(t);
                    }
                }
            }
        }

        var unassignedTasks = allTasks.Where(t => string.IsNullOrEmpty(t.DueDate)).ToList();
        var assignedTasks = allTasks.Where(t => !string.IsNullOrEmpty(t.DueDate)).ToList();

        foreach(var t in unassignedTasks) {
            Label lbl = new Label();
            lbl.Text = t.Text;
            lbl.Width = unassignedPanel.ClientSize.Width - (int)(20 * scale) > 0 ? unassignedPanel.ClientSize.Width - (int)(20 * scale) : (int)(200 * scale);
            lbl.AutoSize = true;
            lbl.Margin = new Padding((int)(5 * scale));
            lbl.Padding = new Padding((int)(10 * scale));
            lbl.BackColor = UITheme.CardWhite;
            lbl.ForeColor = Color.FromName(t.Color);
            lbl.Font = UITheme.GetFont(10.5f);
            lbl.Cursor = Cursors.Hand;
            lbl.BorderStyle = BorderStyle.FixedSingle;
            
            lbl.DoubleClick += (s, e) => {
                CalendarTaskEditForm editF = new CalendarTaskEditForm(t, this);
                editF.ShowDialog();
            };
            unassignedPanel.Controls.Add(lbl);
        }

        DateTime firstDay = new DateTime(targetYear, targetMonth, 1);
        int daysInMonth = DateTime.DaysInMonth(targetYear, targetMonth);
        int startDayOfWeek = (int)firstDay.DayOfWeek; 

        int currentDay = 1;
        for (int row = 0; row < 6; row++) {
            for (int col = 0; col < 7; col++) {
                FlowLayoutPanel cell = new FlowLayoutPanel();
                cell.Dock = DockStyle.Fill;
                cell.FlowDirection = FlowDirection.TopDown;
                cell.WrapContents = false;
                cell.AutoScroll = true;
                cell.Margin = new Padding(0);

                if (row == 0 && col < startDayOfWeek) {
                    cell.BackColor = Color.FromArgb(245, 245, 245);
                } else if (currentDay > daysInMonth) {
                    cell.BackColor = Color.FromArgb(245, 245, 245);
                } else {
                    Label dayNum = new Label();
                    dayNum.Text = currentDay.ToString();
                    dayNum.AutoSize = true;
                    dayNum.Font = UITheme.GetFont(10f, FontStyle.Bold);
                    dayNum.ForeColor = Color.Gray;
                    dayNum.Margin = new Padding((int)(2 * scale));
                    
                    if (targetYear == DateTime.Now.Year && targetMonth == DateTime.Now.Month && currentDay == DateTime.Now.Day) {
                        dayNum.ForeColor = UITheme.CardWhite;
                        dayNum.BackColor = UITheme.AppleBlue;
                    }
                    cell.Controls.Add(dayNum);

                    string dateMatchStr = new DateTime(targetYear, targetMonth, currentDay).ToString("yyyy-MM-dd");
                    var dayTasks = assignedTasks.Where(t => t.DueDate.StartsWith(dateMatchStr)).OrderBy(t => t.DueDate).ToList();

                    foreach (var t in dayTasks) {
                        string timeOnly = "";
                        if (t.DueDate.Length >= 16) {
                            timeOnly = t.DueDate.Substring(11, 5) + " ";
                        }
                        
                        Label tLbl = new Label();
                        tLbl.Text = timeOnly + t.Text;
                        tLbl.AutoSize = true;
                        tLbl.MaximumSize = new Size((int)(calendarGrid.Width / 7f) - (int)(10 * scale), 0);
                        tLbl.Margin = new Padding((int)(2 * scale), 0, 0, (int)(4 * scale));
                        tLbl.Font = UITheme.GetFont(9f);
                        tLbl.ForeColor = Color.FromName(t.Color);
                        tLbl.Cursor = Cursors.Hand;
                        
                        tLbl.DoubleClick += (s, e) => {
                            CalendarTaskEditForm editF = new CalendarTaskEditForm(t, this);
                            editF.ShowDialog();
                        };
                        cell.Controls.Add(tLbl);
                    }
                    currentDay++;
                }
                calendarGrid.Controls.Add(cell, col, row);
            }
        }
        
        calendarGrid.ResumeLayout(true);
        unassignedPanel.ResumeLayout(true);
    }
}

// ============================================================
// 【新增】日曆專用編輯視窗 (連點兩下觸發)
// ============================================================
public class CalendarTaskEditForm : Form {
    private App_TodoList.TaskInfo task;
    private TaskCalendarWindow parent;
    private float scale;

    public CalendarTaskEditForm(App_TodoList.TaskInfo t, TaskCalendarWindow p) {
        this.task = t;
        this.parent = p;
        this.scale = this.DeviceDpi / 96f;

        this.Text = "編輯任務內容";
        this.Width = (int)(450 * scale);
        this.Height = (int)(480 * scale);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.TopMost = true; // 【修復】強制在最上層
        this.BackColor = UITheme.BgGray;

        FlowLayoutPanel f = new FlowLayoutPanel();
        f.Dock = DockStyle.Fill;
        f.FlowDirection = FlowDirection.TopDown;
        f.Padding = new Padding((int)(20 * scale));

        Label l1 = new Label();
        l1.Text = "任務名稱：";
        l1.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        f.Controls.Add(l1);

        TextBox txtName = new TextBox();
        string cleanName = t.Text.Replace("[待辦] ", "").Replace("[待規] ", "").Replace("[行程] ", "");
        txtName.Text = cleanName;
        txtName.Width = (int)(390 * scale);
        txtName.Font = UITheme.GetFont(11f);
        f.Controls.Add(txtName);

        Label l2 = new Label();
        l2.Text = "備註說明：";
        l2.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        l2.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        f.Controls.Add(l2);

        TextBox txtNote = new TextBox();
        txtNote.Text = t.Note;
        txtNote.Width = (int)(390 * scale);
        txtNote.Height = (int)(120 * scale);
        txtNote.Multiline = true;
        txtNote.ScrollBars = ScrollBars.Vertical;
        txtNote.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(txtNote);

        Label l3 = new Label();
        l3.Text = "設定期程 (空白代表未排定)：";
        l3.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        l3.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        f.Controls.Add(l3);

        TextBox txtDate = new TextBox();
        txtDate.Text = t.DueDate;
        txtDate.Width = (int)(390 * scale);
        txtDate.Font = UITheme.GetFont(11f);
        f.Controls.Add(txtDate);
        
        Label hint = new Label();
        hint.Text = "格式：yyyy-MM-dd HH:mm";
        hint.ForeColor = Color.Gray;
        hint.AutoSize = true;
        f.Controls.Add(hint);

        Button btnOk = new Button();
        btnOk.Text = "儲存修改";
        btnOk.Width = (int)(390 * scale);
        btnOk.Height = (int)(45 * scale);
        btnOk.BackColor = UITheme.AppleBlue;
        btnOk.ForeColor = UITheme.CardWhite;
        btnOk.FlatStyle = FlatStyle.Flat;
        btnOk.Font = UITheme.GetFont(11f, FontStyle.Bold);
        btnOk.Margin = new Padding(0, (int)(20 * scale), 0, 0);
        btnOk.FlatAppearance.BorderSize = 0;
        
        btnOk.Click += (s, e) => {
            using (var conn = DbHelper.GetConnection()) {
                conn.Open();
                string sql = "UPDATE Tasks SET Text=@T, Note=@N, DueDate=@D WHERE Id=@Id";
                using (var cmd = new SqliteCommand(sql, conn)) {
                    cmd.Parameters.AddWithValue("@T", txtName.Text.Trim());
                    cmd.Parameters.AddWithValue("@N", txtNote.Text.Trim());
                    cmd.Parameters.AddWithValue("@D", txtDate.Text.Trim());
                    cmd.Parameters.AddWithValue("@Id", t.Id);
                    cmd.ExecuteNonQuery();
                }
            }
            parent.RefreshData();
            this.Close();
        };

        f.Controls.Add(btnOk);
        this.Controls.Add(f);
    }
}
