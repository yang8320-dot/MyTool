// ============================================================
// FILE: MiniProgram01/App_RecurringTasks.cs 
// ============================================================

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;
using System.Drawing.Printing;
using Microsoft.Data.Sqlite;
using System.Threading;

public class App_RecurringTasks : UserControl {
    private MainForm parentForm;
    private App_TodoList todoApp;
    private FlowLayoutPanel taskPanel;
    private System.Windows.Forms.Timer checkTimer;
    private float scale;

    public string digestType { get; set; } = "不提醒";
    public string digestTimeStr { get; set; } = "08:00";
    public string lastDigestDate { get; set; } = "";
    public int advanceDays { get; set; } = 0;
    public string scanFrequency { get; set; } = "10分鐘";

    public class RecurringTask { 
        public int Id { get; set; }
        public string Name { get; set; }
        public string MonthStr { get; set; }
        public string DateStr { get; set; }
        public string TimeStr { get; set; }
        public string LastTriggeredDate { get; set; }
        public string Note { get; set; }
        public string TaskType { get; set; }
        public int OrderIndex { get; set; }
    }
    
    public List<RecurringTask> tasks = new List<RecurringTask>();

    public App_RecurringTasks(MainForm mainForm, App_TodoList todoApp) {
        this.parentForm = mainForm; 
        this.todoApp = todoApp;
        this.scale = this.DeviceDpi / 96f;
        this.BackColor = UITheme.BgGray;
        this.Padding = new Padding((int)(5 * scale));

        TableLayoutPanel header = new TableLayoutPanel();
        header.Dock = DockStyle.Top;
        header.Height = (int)(50 * scale);
        header.ColumnCount = 4;
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(110 * scale)));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(110 * scale)));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(90 * scale)));

        Label lblTitle = new Label() {
            Text = "週期任務", Font = UITheme.GetFont(12f, FontStyle.Bold),
            Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft,
            Padding = new Padding((int)(5 * scale), 0, 0, 0), ForeColor = UITheme.TextMain
        };

        Button btnViewAll = CreateHeaderButton("全部檢視", UITheme.CardWhite, UITheme.AppleBlue);
        btnViewAll.Click += (s, e) => { new AllTasksViewWindow(this).Show(); }; 

        Button btnAdd = CreateHeaderButton("新增任務", UITheme.AppleBlue, UITheme.CardWhite);
        btnAdd.Click += (s, e) => { new AddRecurringTaskWindow(this).Show(); };

        Button btnSet = CreateHeaderButton("設定", Color.Gainsboro, UITheme.TextMain);
        btnSet.Margin = new Padding((int)(2 * scale), (int)(8 * scale), (int)(8 * scale), (int)(8 * scale));
        btnSet.Click += (s, e) => { new RecurringSettingsWindow(this).Show(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnViewAll, 1, 0);
        header.Controls.Add(btnAdd, 2, 0);
        header.Controls.Add(btnSet, 3, 0);
        this.Controls.Add(header);

        taskPanel = new FlowLayoutPanel();
        taskPanel.Dock = DockStyle.Fill;
        taskPanel.AutoScroll = true;
        taskPanel.FlowDirection = FlowDirection.TopDown;
        taskPanel.WrapContents = false;
        taskPanel.BackColor = UITheme.BgGray;
        
        taskPanel.Resize += (s, e) => {
            int safeWidth = taskPanel.ClientSize.Width - (int)(15 * scale);
            if (safeWidth > 0) {
                foreach (Control c in taskPanel.Controls) {
                    if (c is Panel) c.Width = safeWidth;
                }
            }
        };
        
        this.Controls.Add(taskPanel);
        taskPanel.BringToFront();

        LoadSettingsFromDb();
        LoadTasksFromDb();

        checkTimer = new System.Windows.Forms.Timer();
        checkTimer.Interval = GetTimerInterval(scanFrequency);
        checkTimer.Enabled = true;
        checkTimer.Tick += (s, e) => CheckTasks();
        CheckTasks();
    }

    private Button CreateHeaderButton(string text, Color bg, Color fg) {
        Button btn = new Button() {
            Text = text, Dock = DockStyle.Fill, FlatStyle = FlatStyle.Flat,
            Margin = new Padding((int)(2 * scale), (int)(8 * scale), (int)(2 * scale), (int)(8 * scale)),
            Cursor = Cursors.Hand, BackColor = bg, ForeColor = fg, Font = UITheme.GetFont(10f, FontStyle.Bold)
        };
        btn.FlatAppearance.BorderSize = 0;
        return btn;
    }

    private int GetTimerInterval(string freq) {
        switch (freq) {
            case "即時": return 1000; 
            case "1分鐘": return 60000; 
            case "5分鐘": return 300000;
            case "10分鐘": return 600000; 
            case "1小時": return 3600000; 
            case "12小時": return 43200000;
            case "1天": return 86400000; 
            default: return 600000;
        }
    }

    private void LoadSettingsFromDb() {
        digestType = DbHelper.GetSetting("Recur_DigestType", "不提醒");
        digestTimeStr = DbHelper.GetSetting("Recur_DigestTime", "08:00");
        lastDigestDate = DbHelper.GetSetting("Recur_LastDigest", "");
        advanceDays = int.Parse(DbHelper.GetSetting("Recur_AdvanceDays", "0"));
        scanFrequency = DbHelper.GetSetting("Recur_ScanFreq", "10分鐘");
    }

    private void SaveSettingsToDb() {
        DbHelper.SetSetting("Recur_DigestType", digestType);
        DbHelper.SetSetting("Recur_DigestTime", digestTimeStr);
        DbHelper.SetSetting("Recur_LastDigest", lastDigestDate);
        DbHelper.SetSetting("Recur_AdvanceDays", advanceDays.ToString());
        DbHelper.SetSetting("Recur_ScanFreq", scanFrequency);
    }

    private void LoadTasksFromDb() {
        tasks.Clear();
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "SELECT Id, Name, MonthStr, DateStr, TimeStr, TaskType, Note, LastTriggeredDate, OrderIndex FROM RecurringTasks ORDER BY OrderIndex ASC";
            using (var cmd = new SqliteCommand(sql, conn)) {
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) {
                        tasks.Add(new RecurringTask {
                            Id = reader.GetInt32(0), Name = reader.GetString(1),
                            MonthStr = reader.GetString(2), DateStr = reader.GetString(3),
                            TimeStr = reader.GetString(4), TaskType = reader.GetString(5),
                            Note = reader.IsDBNull(6) ? "" : reader.GetString(6),
                            LastTriggeredDate = reader.IsDBNull(7) ? "" : reader.GetString(7),
                            OrderIndex = reader.GetInt32(8)
                        });
                    }
                }
            }
        }
        RefreshUI();
    }

    public void RefreshUI() {
        taskPanel.Controls.Clear();
        int startWidth = taskPanel.ClientSize.Width > (int)(50 * scale) ? taskPanel.ClientSize.Width - (int)(15 * scale) : (int)(450 * scale);
        
        foreach (var t in tasks) {
            Panel card = new Panel() {
                Width = startWidth, AutoSize = true,
                Margin = new Padding((int)(5 * scale), 0, (int)(5 * scale), (int)(3 * scale)),
                BackColor = UITheme.CardWhite,
            };

            card.Paint += (s, e) => {
                UITheme.DrawRoundedBackground(e.Graphics, new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale), UITheme.CardWhite);
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var pen = new Pen(Color.FromArgb(230, 230, 230), 1)) {
                    e.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale)));
                }
            };

            TableLayoutPanel tlp = new TableLayoutPanel() {
                Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, AutoSize = true,
                Padding = new Padding((int)(8 * scale)), BackColor = Color.Transparent
            };
            
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(40 * scale))); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(40 * scale))); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(40 * scale))); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            Button btnEdit = new Button() { Text = "調", Dock = DockStyle.Top, Height = (int)(32 * scale), BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(9f, FontStyle.Bold) };
            btnEdit.FlatAppearance.BorderSize = 0;
            btnEdit.Click += (s, e) => { new EditRecurringTaskWindow(this, t).Show(); };

            Button btnDel = new Button() { Text = "✕", Dock = DockStyle.Top, Height = (int)(32 * scale), BackColor = UITheme.AppleRed, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(9f, FontStyle.Bold) };
            btnDel.FlatAppearance.BorderSize = 0;
            btnDel.Click += (s, e) => { 
                if (MessageBox.Show("確定移除此週期任務？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK) { DeleteTask(t); }
            };

            Button btnNote = new Button() { Text = "註", Dock = DockStyle.Top, Height = (int)(32 * scale), FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(9f, FontStyle.Bold) };
            btnNote.FlatAppearance.BorderSize = 0;
            btnNote.BackColor = string.IsNullOrEmpty(t.Note) ? UITheme.BgGray : UITheme.AppleYellow;
            btnNote.ForeColor = string.IsNullOrEmpty(t.Note) ? UITheme.TextMain : Color.Black;
            btnNote.Click += (s, e) => { 
                string n = ShowNoteEditBox(t.Name, t.Note); 
                if (n != null) { t.Note = n; UpdateTaskInDb(t); RefreshUI(); } 
            };

            string typeTag = $"[{t.TaskType}] ";
            string timeInfo = t.MonthStr == "特定日期" ? $"[{t.DateStr} {t.TimeStr}]" : $"[{t.MonthStr} {t.DateStr} {t.TimeStr}]";

            Label lbl = new Label() {
                Text = typeTag + timeInfo + " " + t.Name, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft,
                AutoSize = true, Font = UITheme.GetFont(10.5f), ForeColor = UITheme.TextMain, Padding = new Padding((int)(5 * scale), 0, 0, 0)
            };
            
            tlp.Controls.Add(btnEdit, 0, 0); tlp.Controls.Add(btnDel, 1, 0); 
            tlp.Controls.Add(btnNote, 2, 0); tlp.Controls.Add(lbl, 3, 0);
            
            card.Controls.Add(tlp); taskPanel.Controls.Add(card);
        }
    }

    public void AddNewTask(string name, string month, string date, string time, string note, string type) {
        int orderIdx = tasks.Count > 0 ? tasks.Max(t => t.OrderIndex) + 1 : 0;
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "INSERT INTO RecurringTasks (Name, MonthStr, DateStr, TimeStr, TaskType, Note, LastTriggeredDate, OrderIndex) VALUES (@N, @M, @D, @T, @Ty, @No, '', @O)";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@N", name); cmd.Parameters.AddWithValue("@M", month);
                cmd.Parameters.AddWithValue("@D", date); cmd.Parameters.AddWithValue("@T", time);
                cmd.Parameters.AddWithValue("@Ty", type); cmd.Parameters.AddWithValue("@No", note);
                cmd.Parameters.AddWithValue("@O", orderIdx); cmd.ExecuteNonQuery();
            }
        }
        LoadTasksFromDb();
    }

    public Tuple<int, int> BulkImportOrUpdate(List<RecurringTask> importedData) {
        int addedCount = 0;
        int updatedCount = 0;

        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            using (var transaction = conn.BeginTransaction()) {
                foreach (var t in importedData) {
                    string checkSql = "SELECT Id FROM RecurringTasks WHERE Name=@N AND MonthStr=@M";
                    int existingId = -1;
                    using (var cmd = new SqliteCommand(checkSql, conn, transaction)) {
                        cmd.Parameters.AddWithValue("@N", t.Name);
                        cmd.Parameters.AddWithValue("@M", t.MonthStr);
                        var result = cmd.ExecuteScalar();
                        if (result != null) existingId = Convert.ToInt32(result);
                    }

                    if (existingId != -1) {
                        string updateSql = "UPDATE RecurringTasks SET DateStr=@D, TimeStr=@T, TaskType=@Ty, Note=@No WHERE Id=@Id";
                        using (var cmd = new SqliteCommand(updateSql, conn, transaction)) {
                            cmd.Parameters.AddWithValue("@D", t.DateStr);
                            cmd.Parameters.AddWithValue("@T", t.TimeStr);
                            cmd.Parameters.AddWithValue("@Ty", t.TaskType);
                            cmd.Parameters.AddWithValue("@No", t.Note);
                            cmd.Parameters.AddWithValue("@Id", existingId);
                            cmd.ExecuteNonQuery();
                        }
                        updatedCount++;
                    } else {
                        string getOrderSql = "SELECT COALESCE(MAX(OrderIndex), 0) + 1 FROM RecurringTasks";
                        int nextOrder = 0;
                        using (var cmd = new SqliteCommand(getOrderSql, conn, transaction)) {
                            nextOrder = Convert.ToInt32(cmd.ExecuteScalar());
                        }

                        string insertSql = "INSERT INTO RecurringTasks (Name, MonthStr, DateStr, TimeStr, TaskType, Note, LastTriggeredDate, OrderIndex) VALUES (@N, @M, @D, @T, @Ty, @No, '', @O)";
                        using (var cmd = new SqliteCommand(insertSql, conn, transaction)) {
                            cmd.Parameters.AddWithValue("@N", t.Name);
                            cmd.Parameters.AddWithValue("@M", t.MonthStr);
                            cmd.Parameters.AddWithValue("@D", t.DateStr);
                            cmd.Parameters.AddWithValue("@T", t.TimeStr);
                            cmd.Parameters.AddWithValue("@Ty", t.TaskType);
                            cmd.Parameters.AddWithValue("@No", t.Note);
                            cmd.Parameters.AddWithValue("@O", nextOrder);
                            cmd.ExecuteNonQuery();
                        }
                        addedCount++;
                    }
                }
                transaction.Commit();
            }
        }
        LoadTasksFromDb(); 
        return new Tuple<int, int>(addedCount, updatedCount);
    }

    public void UpdateTaskDb(RecurringTask t) { UpdateTaskInDb(t); LoadTasksFromDb(); }

    private void UpdateTaskInDb(RecurringTask t) {
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "UPDATE RecurringTasks SET Name=@N, MonthStr=@M, DateStr=@D, TimeStr=@T, TaskType=@Ty, Note=@No, LastTriggeredDate=@L, OrderIndex=@O WHERE Id=@Id";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@N", t.Name); cmd.Parameters.AddWithValue("@M", t.MonthStr);
                cmd.Parameters.AddWithValue("@D", t.DateStr); cmd.Parameters.AddWithValue("@T", t.TimeStr);
                cmd.Parameters.AddWithValue("@Ty", t.TaskType); cmd.Parameters.AddWithValue("@No", t.Note);
                cmd.Parameters.AddWithValue("@L", t.LastTriggeredDate ?? ""); cmd.Parameters.AddWithValue("@O", t.OrderIndex);
                cmd.Parameters.AddWithValue("@Id", t.Id); cmd.ExecuteNonQuery();
            }
        }
    }

    public void DeleteTask(RecurringTask task) { 
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            using (var cmd = new SqliteCommand("DELETE FROM RecurringTasks WHERE Id=@Id", conn)) {
                cmd.Parameters.AddWithValue("@Id", task.Id); cmd.ExecuteNonQuery();
            }
        }
        LoadTasksFromDb();
    }

    public void UpdateGlobalSettings(string dType, string dTime, int aDays, string sFreq) {
        digestType = dType; digestTimeStr = dTime; advanceDays = aDays; scanFrequency = sFreq;
        checkTimer.Enabled = false; checkTimer.Interval = GetTimerInterval(sFreq); checkTimer.Enabled = true;
        SaveSettingsToDb(); MessageBox.Show("設定儲存成功！");
    }

    private void CheckTasks() {
        DateTime now = DateTime.Now; bool needsRefresh = false;
        List<RecurringTask> toRemove = new List<RecurringTask>();

        foreach (var t in tasks) {
            DateTime target;
            if (TryGetNextTriggerTime(t, now, out target)) {
                DateTime triggerThreshold = target.AddDays(-advanceDays);
                if (now >= triggerThreshold) {
                    string targetDateStr = target.ToString("yyyy-MM-dd");
                    if (t.LastTriggeredDate != targetDateStr) {
                        string prefix = advanceDays > 0 ? $"[預排-{target:MM/dd}] " : "";
                        todoApp.AddTask(prefix + t.Name, "Black", "週期觸發", t.Note); 
                        
                        t.LastTriggeredDate = targetDateStr; UpdateTaskInDb(t);
                        parentForm.AlertTab(1); 
                        
                        if (t.TaskType == "單次" || t.TaskType == "到期日") toRemove.Add(t);
                    }
                }
            }
        }

        if (toRemove.Count > 0) { foreach (var r in toRemove) DeleteTask(r); needsRefresh = true; }
        if (needsRefresh) this.Invoke(new Action(() => LoadTasksFromDb()));

        if (digestType != "不提醒") {
            DateTime dtDigest;
            if (DateTime.TryParseExact(digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtDigest)) {
                DateTime targetDigest = new DateTime(now.Year, now.Month, now.Day, dtDigest.Hour, dtDigest.Minute, 0);
                bool shouldTrigger = false;
                if (digestType == "每週一" && now.DayOfWeek == DayOfWeek.Monday && now >= targetDigest) shouldTrigger = true;
                if (digestType == "每月1號" && now.Day == 1 && now >= targetDigest) shouldTrigger = true;
                
                string todayStr = now.ToString("yyyy-MM-dd");
                if (shouldTrigger && lastDigestDate != todayStr) { 
                    lastDigestDate = todayStr; SaveSettingsToDb();
                    new AllTasksViewWindow(this).Show(); 
                }
            }
        }
    }

    private bool TryGetNextTriggerTime(RecurringTask t, DateTime now, out DateTime target) {
        target = now;
        try {
            string[] timeParts = t.TimeStr.Split(':');
            int h = int.Parse(timeParts[0]); int m = int.Parse(timeParts[1]);

            if (t.MonthStr == "特定日期") {
                if (DateTime.TryParseExact(t.DateStr, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime specificDate)) {
                    target = new DateTime(specificDate.Year, specificDate.Month, specificDate.Day, h, m, 0);
                    return true;
                }
                return false;
            }

            if (t.MonthStr == "每天") {
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0);
                if (now > target && t.LastTriggeredDate == target.ToString("yyyy-MM-dd")) {
                    target = target.AddDays(1);
                }
                return true;
            }
            
            if (t.MonthStr == "每週") {
                Dictionary<string, DayOfWeek> dow = new Dictionary<string, DayOfWeek> {
                    {"一", DayOfWeek.Monday}, {"二", DayOfWeek.Tuesday}, {"三", DayOfWeek.Wednesday},
                    {"四", DayOfWeek.Thursday}, {"五", DayOfWeek.Friday}, {"六", DayOfWeek.Saturday}, {"日", DayOfWeek.Sunday}
                };
                if (!dow.ContainsKey(t.DateStr)) return false;
                
                int diff = dow[t.DateStr] - now.DayOfWeek;
                target = new DateTime(now.Year, now.Month, now.Day, h, m, 0).AddDays(diff);
                
                if (now > target && t.LastTriggeredDate == target.ToString("yyyy-MM-dd")) {
                    target = target.AddDays(7);
                }
                return true;
            }
            
            if (t.MonthStr == "每月" || t.MonthStr.EndsWith("月")) {
                int month = t.MonthStr == "每月" ? now.Month : int.Parse(t.MonthStr.Replace("月",""));
                int day = (t.DateStr == "月底") ? DateTime.DaysInMonth(now.Year, month) : int.Parse(t.DateStr);
                int validDay = Math.Min(day, DateTime.DaysInMonth(now.Year, month));
                
                target = new DateTime(now.Year, month, validDay, h, m, 0);
                
                if (now > target && t.LastTriggeredDate == target.ToString("yyyy-MM-dd")) {
                    if (t.MonthStr == "每月") {
                        target = target.AddMonths(1);
                        validDay = Math.Min(day, DateTime.DaysInMonth(target.Year, target.Month));
                        target = new DateTime(target.Year, target.Month, validDay, h, m, 0);
                    } else {
                        target = target.AddYears(1); 
                    }
                }
                return true;
            }
        } catch { } 
        return false;
    }

    private string ShowNoteEditBox(string name, string current) {
        Form f = new Form() {
            Width = (int)(420 * scale), Height = (int)(380 * scale), 
            Text = "編輯備註", StartPosition = FormStartPosition.CenterScreen, 
            FormBorderStyle = FormBorderStyle.FixedDialog, TopMost = true, BackColor = UITheme.BgGray
        };

        Label lbl = new Label() { Text = "【" + name + "】", Left = (int)(15 * scale), Top = (int)(15 * scale), AutoSize = true, Font = UITheme.GetFont(11f, FontStyle.Bold) };
        TextBox txt = new TextBox() { Left = (int)(15 * scale), Top = (int)(50 * scale), Width = (int)(370 * scale), Height = (int)(200 * scale), Multiline = true, AcceptsReturn = true, Text = current, Font = UITheme.GetFont(10.5f) };
        Button btn = new Button() { Text = "儲存", Left = (int)(285 * scale), Top = (int)(280 * scale), Width = (int)(100 * scale), Height = (int)(40 * scale), DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, Font = UITheme.GetFont(10f, FontStyle.Bold) };
        btn.FlatAppearance.BorderSize = 0;

        f.Controls.AddRange(new Control[] { lbl, txt, btn });
        return f.ShowDialog() == DialogResult.OK ? txt.Text : null;
    }
}

public class AddRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private TextBox txtN, txtNote;
    private ComboBox cmM, cmD, cmType;
    private DateTimePicker dtpTime, dtpDate;
    private Label lblCycle, lblDate;

    public AddRecurringTaskWindow(App_RecurringTasks p) {
        this.parent = p; 
        float scale = this.DeviceDpi / 96f;
        this.Text = "新增任務"; this.Width = (int)(420 * scale); this.Height = (int)(680 * scale); 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true; this.BackColor = UITheme.BgGray;

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding((int)(25 * scale)) };

        f.Controls.Add(new Label() { Text = "任務名稱：", Font = UITheme.GetFont(10f, FontStyle.Bold) }); 
        txtN = new TextBox() { Width = (int)(340 * scale), Font = UITheme.GetFont(10.5f) }; f.Controls.Add(txtN);

        f.Controls.Add(new Label() { Text = "詳細說明 (註)：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) });
        txtNote = new TextBox() { Width = (int)(340 * scale), Height = (int)(80 * scale), Multiline = true, AcceptsReturn = true, Font = UITheme.GetFont(10.5f) }; f.Controls.Add(txtNote);
        
        f.Controls.Add(new Label() { Text = "任務類型：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) });
        cmType = new ComboBox() { Width = (int)(340 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10.5f) };
        cmType.Items.AddRange(new string[] { "循環", "單次", "到期日" });
        f.Controls.Add(cmType);

        lblCycle = new Label() { Text = "週期類型：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) };
        f.Controls.Add(lblCycle);
        
        cmM = new ComboBox() { Width = (int)(340 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10.5f) };
        cmM.Items.AddRange(new string[] { "每天", "每週", "每月" });
        for(int i = 1; i <= 12; i++) cmM.Items.Add(i.ToString() + "月");
        f.Controls.Add(cmM); 
        
        cmD = new ComboBox() { Width = (int)(340 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10.5f) }; 
        f.Controls.Add(cmD);

        lblDate = new Label() { Text = "指定日期：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) };
        f.Controls.Add(lblDate);
        dtpDate = new DateTimePicker() { Width = (int)(340 * scale), Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Font = UITheme.GetFont(10.5f) };
        f.Controls.Add(dtpDate);

        cmType.SelectedIndexChanged += (s, e) => {
            bool isLoop = cmType.Text == "循環";
            lblCycle.Visible = cmM.Visible = cmD.Visible = isLoop;
            lblDate.Visible = dtpDate.Visible = !isLoop;
        };
        cmType.SelectedIndex = 0; 

        cmM.SelectedIndexChanged += (s, e) => {
            cmD.Items.Clear();
            if(cmM.Text == "每天") { cmD.Items.Add("每日"); cmD.Enabled = false; }
            else if(cmM.Text == "每週") { 
                cmD.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" }); cmD.Enabled = true; 
            } else { 
                for(int i = 1; i <= 31; i++) cmD.Items.Add(i.ToString()); 
                cmD.Items.Add("月底"); cmD.Enabled = true; 
            }
            if (cmD.Items.Count > 0) cmD.SelectedIndex = 0;
        }; 
        cmM.SelectedIndex = 0;

        f.Controls.Add(new Label() { Text = "觸發時間：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) });
        dtpTime = new DateTimePicker() { Width = (int)(340 * scale), Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Value = DateTime.Today.AddHours(9), Font = UITheme.GetFont(10.5f) };
        f.Controls.Add(dtpTime);

        Button btn = new Button() { Text = "建立任務", Width = (int)(340 * scale), Height = (int)(45 * scale), BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, (int)(25 * scale), 0, 0), Font = UITheme.GetFont(11f, FontStyle.Bold), Cursor = Cursors.Hand };
        btn.FlatAppearance.BorderSize = 0;
        btn.Click += (s, e) => { 
            if(!string.IsNullOrWhiteSpace(txtN.Text)) { 
                string monthVal = cmType.Text == "循環" ? cmM.Text : "特定日期";
                string dateVal = cmType.Text == "循環" ? cmD.Text : dtpDate.Value.ToString("yyyy-MM-dd");
                parent.AddNewTask(txtN.Text, monthVal, dateVal, dtpTime.Value.ToString("HH:mm"), txtNote.Text, cmType.Text); 
                this.Close(); 
            } 
        };
        f.Controls.Add(btn); 
        this.Controls.Add(f);
    }
}

public class EditRecurringTaskWindow : Form {
    private App_RecurringTasks parent;
    private App_RecurringTasks.RecurringTask task;
    private TextBox txtN, txtNote;
    private ComboBox cmM, cmD, cmType;
    private DateTimePicker dtpTime, dtpDate;
    private Label lblCycle, lblDate;

    public EditRecurringTaskWindow(App_RecurringTasks p, App_RecurringTasks.RecurringTask t) {
        this.parent = p; this.task = t; 
        float scale = this.DeviceDpi / 96f;
        this.Text = "調整任務"; this.Width = (int)(420 * scale); this.Height = (int)(680 * scale); 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true; this.BackColor = UITheme.BgGray;

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding((int)(25 * scale)) };

        f.Controls.Add(new Label() { Text = "任務名稱：", Font = UITheme.GetFont(10f, FontStyle.Bold) }); 
        txtN = new TextBox() { Width = (int)(340 * scale), Text = t.Name, Font = UITheme.GetFont(10.5f) }; f.Controls.Add(txtN);

        f.Controls.Add(new Label() { Text = "詳細說明 (註)：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) });
        txtNote = new TextBox() { Width = (int)(340 * scale), Height = (int)(80 * scale), Multiline = true, AcceptsReturn = true, Text = t.Note, Font = UITheme.GetFont(10.5f) }; f.Controls.Add(txtNote);
        
        f.Controls.Add(new Label() { Text = "任務類型：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) });
        cmType = new ComboBox() { Width = (int)(340 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10.5f) };
        cmType.Items.AddRange(new string[] { "循環", "單次", "到期日" });
        cmType.Text = t.TaskType; f.Controls.Add(cmType);

        lblCycle = new Label() { Text = "週期類型：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) };
        f.Controls.Add(lblCycle);
        
        cmM = new ComboBox() { Width = (int)(340 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10.5f) };
        cmM.Items.AddRange(new string[] { "每天", "每週", "每月" });
        for(int k = 1; k <= 12; k++) cmM.Items.Add(k.ToString() + "月");
        f.Controls.Add(cmM);

        cmD = new ComboBox() { Width = (int)(340 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10.5f) }; f.Controls.Add(cmD);

        lblDate = new Label() { Text = "指定日期：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) };
        f.Controls.Add(lblDate);
        
        dtpDate = new DateTimePicker() { Width = (int)(340 * scale), Format = DateTimePickerFormat.Custom, CustomFormat = "yyyy-MM-dd", Font = UITheme.GetFont(10.5f) };
        f.Controls.Add(dtpDate);

        cmType.SelectedIndexChanged += (s, e) => {
            bool isLoop = cmType.Text == "循環";
            lblCycle.Visible = cmM.Visible = cmD.Visible = isLoop;
            lblDate.Visible = dtpDate.Visible = !isLoop;
        };

        if (t.MonthStr == "特定日期") {
            if (DateTime.TryParseExact(t.DateStr, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out DateTime d)) dtpDate.Value = d;
        } else {
            cmM.Text = t.MonthStr;
        }

        cmM.SelectedIndexChanged += (s, e) => {
            cmD.Items.Clear();
            if(cmM.Text == "每天") { cmD.Items.Add("每日"); }
            else if(cmM.Text == "每週") { cmD.Items.AddRange(new string[] { "一", "二", "三", "四", "五", "六", "日" }); }
            else { 
                for(int k = 1; k <= 31; k++) cmD.Items.Add(k.ToString()); 
                cmD.Items.Add("月底"); 
            }
            if(cmD.Items.Count > 0) cmD.SelectedIndex = 0;
        }; 
        if (t.MonthStr != "特定日期") cmD.Text = t.DateStr;

        f.Controls.Add(new Label() { Text = "觸發時間：", Margin = new Padding(0, (int)(15 * scale), 0, 0), Font = UITheme.GetFont(10f, FontStyle.Bold) });
        dtpTime = new DateTimePicker() { Width = (int)(340 * scale), Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Font = UITheme.GetFont(10.5f) };
        if(DateTime.TryParseExact(t.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime dtv)) dtpTime.Value = dtv;
        f.Controls.Add(dtpTime);

        Button btn = new Button() { Text = "儲存修改", Width = (int)(340 * scale), Height = (int)(45 * scale), BackColor = UITheme.AppleGreen, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Margin = new Padding(0, (int)(25 * scale), 0, 0), Font = UITheme.GetFont(11f, FontStyle.Bold), Cursor = Cursors.Hand };
        btn.FlatAppearance.BorderSize = 0;
        btn.Click += (s, e) => { 
            t.Name = txtN.Text;
            t.TaskType = cmType.Text;
            t.MonthStr = cmType.Text == "循環" ? cmM.Text : "特定日期";
            t.DateStr = cmType.Text == "循環" ? cmD.Text : dtpDate.Value.ToString("yyyy-MM-dd");
            t.TimeStr = dtpTime.Value.ToString("HH:mm");
            t.Note = txtNote.Text;
            parent.UpdateTaskDb(t);
            this.Close(); 
        };
        f.Controls.Add(btn); 
        this.Controls.Add(f);
    }
}

public class RecurringSettingsWindow : Form {
    private App_RecurringTasks parent;
    private ComboBox cmDig, cmAdv, cmScan;
    private DateTimePicker dtp;
    
    public RecurringSettingsWindow(App_RecurringTasks p) {
        this.parent = p; 
        float scale = this.DeviceDpi / 96f;
        this.Text = "全域排程設定"; this.Width = (int)(380 * scale); this.Height = (int)(380 * scale); 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true; this.BackColor = UITheme.BgGray;

        FlowLayoutPanel f = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, Padding = new Padding((int)(20 * scale)) };
        
        FlowLayoutPanel r1 = new FlowLayoutPanel() { AutoSize = true };
        r1.Controls.Add(new Label() { Text = "所有任務提前", AutoSize = true, Margin = new Padding(0, (int)(5 * scale), 0, 0), Font = UITheme.GetFont(10.5f) });
        cmAdv = new ComboBox() { Width = (int)(70 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10.5f) }; 
        for (int i = 0; i <= 7; i++) cmAdv.Items.Add(i.ToString());
        cmAdv.Text = p.advanceDays.ToString(); 
        r1.Controls.Add(cmAdv); 
        r1.Controls.Add(new Label() { Text = "天加入待辦", AutoSize = true, Margin = new Padding(0, (int)(5 * scale), 0, 0), Font = UITheme.GetFont(10.5f) });
        f.Controls.Add(r1);
        
        FlowLayoutPanel r2 = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, (int)(15 * scale), 0, (int)(15 * scale)) };
        r2.Controls.Add(new Label() { Text = "視窗摘要提醒：", AutoSize = true, Margin = new Padding(0, (int)(5 * scale), 0, 0), Font = UITheme.GetFont(10.5f) });
        cmDig = new ComboBox() { Width = (int)(100 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10.5f) };
        cmDig.Items.Add("不提醒"); cmDig.Items.Add("每週一"); cmDig.Items.Add("每月1號");
        cmDig.Text = p.digestType;
        dtp = new DateTimePicker() { Width = (int)(80 * scale), Format = DateTimePickerFormat.Custom, CustomFormat = "HH:mm", ShowUpDown = true, Font = UITheme.GetFont(10.5f) };
        if(DateTime.TryParseExact(p.digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out DateTime dtv)) dtp.Value = dtv;
        r2.Controls.Add(cmDig); r2.Controls.Add(dtp); f.Controls.Add(r2);
        
        f.Controls.Add(new Label() { AutoSize = false, Height = 2, Width = (int)(320 * scale), BorderStyle = BorderStyle.Fixed3D, Margin = new Padding(0, (int)(5 * scale), 0, (int)(20 * scale)) });

        FlowLayoutPanel r3 = new FlowLayoutPanel() { AutoSize = true, Margin = new Padding(0, 0, 0, (int)(25 * scale)) };
        r3.Controls.Add(new Label() { Text = "背景掃描頻率：", AutoSize = true, Margin = new Padding(0, (int)(5 * scale), 0, 0), Font = UITheme.GetFont(10.5f) });
        cmScan = new ComboBox() { Width = (int)(120 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10.5f) };
        cmScan.Items.AddRange(new string[] { "即時", "1分鐘", "5分鐘", "10分鐘", "1小時", "12小時", "1天" });
        cmScan.Text = p.scanFrequency;
        r3.Controls.Add(cmScan); f.Controls.Add(r3);

        Button btn = new Button() { Text = "儲存所有設定", Width = (int)(320 * scale), Height = (int)(45 * scale), BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(11f, FontStyle.Bold), Cursor = Cursors.Hand };
        btn.FlatAppearance.BorderSize = 0;
        btn.Click += (s, e) => { 
            int.TryParse(cmAdv.Text, out int advDays);
            p.UpdateGlobalSettings(cmDig.Text, dtp.Value.ToString("HH:mm"), advDays, cmScan.Text); 
            this.Close(); 
        };
        f.Controls.Add(btn); this.Controls.Add(f);
    }
}

public class AllTasksViewWindow : Form {
    private App_RecurringTasks parentControl;
    private FlowLayoutPanel flow;
    private float scale;
    private bool isRefreshing = false;

    public AllTasksViewWindow(App_RecurringTasks parent) {
        this.parentControl = parent; 
        this.scale = this.DeviceDpi / 96f;
        this.Text = "週期任務總覽"; 
        this.Width = (int)(900 * scale); 
        this.Height = (int)(850 * scale); 
        this.BackColor = UITheme.BgGray;
        this.TopMost = true; 

        TableLayoutPanel header = new TableLayoutPanel() { Dock = DockStyle.Top, Height = (int)(70 * scale), BackColor = UITheme.CardWhite, ColumnCount = 4 };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(110 * scale)));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(110 * scale)));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(110 * scale)));

        Label lbl = new Label() { Text = "週期任務排程總覽", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding((int)(20 * scale),0,0,0), Font = UITheme.GetFont(16f, FontStyle.Bold), ForeColor = UITheme.TextMain };
        
        Button btnImport = new Button() { Text = "匯入 Excel", Dock = DockStyle.Fill, BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Margin = new Padding((int)(5*scale), (int)(15*scale), (int)(5*scale), (int)(15*scale)), Font = UITheme.GetFont(10f, FontStyle.Bold), Cursor = Cursors.Hand };
        btnImport.FlatAppearance.BorderSize = 0;
        btnImport.Click += (s, e) => ExecuteImportExcel();

        Button btnExport = new Button() { Text = "匯出 Excel", Dock = DockStyle.Fill, BackColor = UITheme.AppleGreen, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Margin = new Padding((int)(5*scale), (int)(15*scale), (int)(5*scale), (int)(15*scale)), Font = UITheme.GetFont(10f, FontStyle.Bold), Cursor = Cursors.Hand };
        btnExport.FlatAppearance.BorderSize = 0;
        btnExport.Click += (s, e) => ExecuteExportExcel();

        Button btnPrint = new Button() { Text = "導出 PDF", Dock = DockStyle.Fill, BackColor = Color.Gray, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Margin = new Padding((int)(5*scale), (int)(15*scale), (int)(15*scale), (int)(15*scale)), Font = UITheme.GetFont(10.5f, FontStyle.Bold), Cursor = Cursors.Hand };
        btnPrint.FlatAppearance.BorderSize = 0;
        btnPrint.Click += (s, e) => ExecuteExportPDF();

        header.Controls.Add(lbl, 0, 0); 
        header.Controls.Add(btnImport, 1, 0); 
        header.Controls.Add(btnExport, 2, 0); 
        header.Controls.Add(btnPrint, 3, 0); 
        this.Controls.Add(header);

        flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding((int)(20 * scale)), FlowDirection = FlowDirection.TopDown, WrapContents = false };
        flow.Resize += (s, e) => { 
            int w = flow.ClientSize.Width - (int)(40 * scale); 
            if (w > 0) foreach (Control c in flow.Controls) if (c is Panel) c.Width = w;
        };
        this.Controls.Add(flow); flow.BringToFront(); RefreshData();
    }

    private void ExecuteExportExcel() {
        Type excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null) {
            MessageBox.Show("系統偵測不到 Microsoft Excel，請確認是否已安裝軟體。", "無法匯出", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel 活頁簿|*.xlsx", FileName = $"週期任務總覽_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx" }) {
            if (sfd.ShowDialog() == DialogResult.OK) {
                dynamic excelApp = null;
                dynamic workbook = null;
                try {
                    excelApp = Activator.CreateInstance(excelType);
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    workbook = excelApp.Workbooks.Add();
                    dynamic sheet = workbook.Sheets[1];

                    sheet.Cells[1, 1] = "任務名稱";
                    sheet.Cells[1, 2] = "任務類型";
                    sheet.Cells[1, 3] = "週期類型";
                    sheet.Cells[1, 4] = "指定日期";
                    sheet.Cells[1, 5] = "觸發時間";
                    sheet.Cells[1, 6] = "備註";

                    int row = 2;
                    foreach (var t in parentControl.tasks) {
                        sheet.Cells[row, 1] = t.Name;
                        sheet.Cells[row, 2] = t.TaskType;
                        sheet.Cells[row, 3] = t.MonthStr;
                        sheet.Cells[row, 4] = t.DateStr;
                        sheet.Cells[row, 5] = t.TimeStr;
                        sheet.Cells[row, 6] = t.Note;
                        row++;
                    }

                    sheet.Rows.RowHeight = 25;           
                    sheet.Columns[1].ColumnWidth = 25;   
                    for (int i = 2; i <= 6; i++) {
                        sheet.Columns[i].ColumnWidth = 12; 
                    }

                    workbook.SaveAs(sfd.FileName);
                    MessageBox.Show("Excel 檔案已成功導出！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } catch (Exception ex) {
                    MessageBox.Show("匯出時發生錯誤：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } finally {
                    if (workbook != null) { workbook.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook); }
                    if (excelApp != null) { excelApp.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp); }
                }
            }
        }
    }

    private void ExecuteImportExcel() {
        Type excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null) {
            MessageBox.Show("系統偵測不到 Microsoft Excel，請確認是否已安裝軟體。", "無法匯入", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel 活頁簿|*.xlsx;*.xls", Title = "選擇要匯入的排程檔案" }) {
            if (ofd.ShowDialog() == DialogResult.OK) {
                dynamic excelApp = null;
                dynamic workbook = null;
                try {
                    excelApp = Activator.CreateInstance(excelType);
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    workbook = excelApp.Workbooks.Open(ofd.FileName);
                    dynamic sheet = workbook.Sheets[1];

                    dynamic lastCell = sheet.Cells.SpecialCells(11); 
                    int lastRow = lastCell.Row;

                    List<App_RecurringTasks.RecurringTask> importList = new List<App_RecurringTasks.RecurringTask>();

                    for (int row = 2; row <= lastRow; row++) {
                        string name = Convert.ToString(sheet.Cells[row, 1].Text);
                        if (string.IsNullOrWhiteSpace(name) || name == "任務名稱") continue;

                        importList.Add(new App_RecurringTasks.RecurringTask {
                            Name = name,
                            TaskType = Convert.ToString(sheet.Cells[row, 2].Text),
                            MonthStr = Convert.ToString(sheet.Cells[row, 3].Text),
                            DateStr = Convert.ToString(sheet.Cells[row, 4].Text),
                            TimeStr = Convert.ToString(sheet.Cells[row, 5].Text),
                            Note = Convert.ToString(sheet.Cells[row, 6].Text)
                        });
                    }

                    var resultStats = parentControl.BulkImportOrUpdate(importList);

                    MessageBox.Show(
                        $"Excel 匯入完成！\n\n新增了 {resultStats.Item1} 筆任務\n更新了 {resultStats.Item2} 筆任務", 
                        "匯入成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                    RefreshData();
                } catch (Exception ex) {
                    MessageBox.Show("檔案格式有誤或被其他程式鎖定。\n詳細錯誤：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                } finally {
                    if (workbook != null) { workbook.Close(false); System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook); }
                    if (excelApp != null) { excelApp.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp); }
                }
            }
        }
    }

    private void ExecuteExportPDF() {
        using (SaveFileDialog sfd = new SaveFileDialog()) {
            sfd.Filter = "PDF 檔案|*.pdf";
            sfd.FileName = $"週期任務總覽_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            
            if (sfd.ShowDialog() == DialogResult.OK) {
                PrintDocument pd = new PrintDocument();
                pd.PrinterSettings.PrinterName = "Microsoft Print to PDF";
                pd.PrinterSettings.PrintToFile = true;
                pd.PrinterSettings.PrintFileName = sfd.FileName;

                int currentLine = 0;
                Font titleFont = UITheme.GetFont(18f, FontStyle.Bold);
                Font headerFont = UITheme.GetFont(12f, FontStyle.Bold);
                Font txtFont = UITheme.GetFont(10f);
                Font noteFont = UITheme.GetFont(9f);

                var tasks = parentControl.tasks;
                
                List<Tuple<string, Font, Brush>> lines = new List<Tuple<string, Font, Brush>>();
                lines.Add(new Tuple<string, Font, Brush>("週期任務排程總覽", titleFont, Brushes.Black));
                lines.Add(new Tuple<string, Font, Brush>("產生時間: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), txtFont, Brushes.Gray));
                lines.Add(new Tuple<string, Font, Brush>(" ", txtFont, Brushes.Black));

                Action<string, List<App_RecurringTasks.RecurringTask>> addGroup = (title, list) => {
                    if (list.Count == 0) return;
                    lines.Add(new Tuple<string, Font, Brush>("【 " + title + " 】", headerFont, Brushes.Blue));
                    foreach (var t in list) {
                        string timeInfo = t.MonthStr == "特定日期" ? $"[{t.DateStr} {t.TimeStr}]" : $"[{t.TimeStr}] {t.DateStr}";
                        string mainTxt = $"[{t.TaskType}] {timeInfo}  {t.Name}";
                        lines.Add(new Tuple<string, Font, Brush>(mainTxt, txtFont, Brushes.Black));
                        
                        if (!string.IsNullOrWhiteSpace(t.Note)) {
                            string notePrefix = "   備註:\n      ";
                            string formattedNote = notePrefix + t.Note.Replace("\n", "\n      ");
                            lines.Add(new Tuple<string, Font, Brush>(formattedNote, noteFont, Brushes.DimGray));
                        }
                    }
                    lines.Add(new Tuple<string, Font, Brush>(" ", txtFont, Brushes.Black));
                };

                addGroup("每天觸發", tasks.Where(t => t.MonthStr == "每天").ToList());
                addGroup("每週觸發", tasks.Where(t => t.MonthStr == "每週").ToList());
                addGroup("每月觸發", tasks.Where(t => t.MonthStr == "每月").ToList());
                for (int i = 1; i <= 12; i++) addGroup(i.ToString() + "月 限定", tasks.Where(t => t.MonthStr == i.ToString() + "月").ToList());
                addGroup("特定日期 (單次/到期日)", tasks.Where(t => t.MonthStr == "特定日期").ToList());

                pd.PrintPage += (sender, args) => {
                    float yPos = args.MarginBounds.Top;
                    float leftMargin = args.MarginBounds.Left;

                    while (currentLine < lines.Count) {
                        var item = lines[currentLine];
                        SizeF size = args.Graphics.MeasureString(item.Item1, item.Item2, args.MarginBounds.Width);
                        
                        if (yPos + size.Height > args.MarginBounds.Bottom) {
                            args.HasMorePages = true;
                            return; 
                        }

                        args.Graphics.DrawString(item.Item1, item.Item2, item.Item3, new RectangleF(leftMargin, yPos, args.MarginBounds.Width, size.Height));
                        yPos += size.Height + 5; 
                        currentLine++;
                    }
                    args.HasMorePages = false;
                };

                try {
                    pd.Print();
                    MessageBox.Show("PDF 檔案已成功導出！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                } catch (Exception ex) {
                    MessageBox.Show("導出失敗！請確認您的 Windows 系統是否有安裝「Microsoft Print to PDF」虛擬印表機功能。\n詳細錯誤：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    public void RefreshData() {
        if (isRefreshing) return;
        isRefreshing = true;
        flow.SuspendLayout();
        
        try {
            flow.Controls.Clear(); 
            var tasks = parentControl.tasks;
            
            AddGroup("每天觸發", tasks.Where(t => t.MonthStr == "每天").ToList());
            AddGroup("每週觸發", tasks.Where(t => t.MonthStr == "每週").ToList());
            AddGroup("每月觸發", tasks.Where(t => t.MonthStr == "每月").ToList());
            for (int i = 1; i <= 12; i++) AddGroup(i.ToString() + "月 限定", tasks.Where(t => t.MonthStr == i.ToString() + "月").ToList());
            AddGroup("特定日期 (單次/到期日)", tasks.Where(t => t.MonthStr == "特定日期").ToList());
        } finally {
            flow.ResumeLayout();
            isRefreshing = false;
        }
    }

    private void AddGroup(string header, List<App_RecurringTasks.RecurringTask> sub) {
        if (sub.Count == 0) return;

        Panel gb = new Panel() { AutoSize = true, Width = flow.ClientSize.Width - (int)(40 * scale), Margin = new Padding((int)(10 * scale), (int)(10 * scale), (int)(10 * scale), (int)(25 * scale)), Padding = new Padding((int)(15 * scale)), BackColor = UITheme.CardWhite };
        
        gb.Paint += (s, e) => {
            UITheme.DrawRoundedBackground(e.Graphics, new Rectangle(0, 0, gb.Width - 1, gb.Height - 1), (int)(12 * scale), UITheme.CardWhite);
            using (var pen = new Pen(Color.FromArgb(220, 220, 220), 1)) {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, gb.Width - 1, gb.Height - 1), (int)(12 * scale)));
            }
        };

        FlowLayoutPanel inner = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, BackColor = Color.Transparent };
        
        Label titleLbl = new Label() { Text = header, Font = UITheme.GetFont(13f, FontStyle.Bold), ForeColor = UITheme.AppleBlue, AutoSize = true, Margin = new Padding(0, 0, 0, (int)(15 * scale)) };
        inner.Controls.Add(titleLbl);

        foreach (var t in sub) {
            TableLayoutPanel row = new TableLayoutPanel() { Width = gb.Width - (int)(40 * scale), AutoSize = true, ColumnCount = 4, Margin = new Padding(0,0,0,(int)(10 * scale)) };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(45 * scale))); 
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(45 * scale)));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(45 * scale))); 
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            Button bE = new Button() { Text = "調", Height = (int)(32 * scale), Dock = DockStyle.Top, BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(9f, FontStyle.Bold), Cursor = Cursors.Hand };
            bE.FlatAppearance.BorderSize = 0;
            bE.Click += (s, e) => { 
                new EditRecurringTaskWindow(parentControl, t).ShowDialog(); 
                RefreshData(); 
            };

            Button bD = new Button() { Text = "✕", Height = (int)(32 * scale), Dock = DockStyle.Top, BackColor = UITheme.AppleRed, ForeColor = UITheme.CardWhite, FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(9f, FontStyle.Bold), Cursor = Cursors.Hand };
            bD.FlatAppearance.BorderSize = 0;
            bD.Click += (s, e) => { if (MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel) == DialogResult.OK) { parentControl.DeleteTask(t); RefreshData(); } };

            Button bN = new Button() { Text = "註", Height = (int)(32 * scale), Dock = DockStyle.Top, FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(9f, FontStyle.Bold), Cursor = Cursors.Hand };
            bN.FlatAppearance.BorderSize = 0;
            if (!string.IsNullOrEmpty(t.Note)) { bN.BackColor = UITheme.AppleYellow; bN.ForeColor = Color.Black; } 
            else { bN.BackColor = UITheme.BgGray; bN.ForeColor = UITheme.TextSub; }
            
            bN.Click += (s, e) => { 
                Form nf = new Form() { Width = (int)(420 * scale), Height = (int)(380 * scale), Text = "任務備註", StartPosition = FormStartPosition.CenterScreen, TopMost = true, BackColor = UITheme.BgGray };
                TextBox nt = new TextBox() { Left = (int)(15 * scale), Top = (int)(50 * scale), Width = (int)(370 * scale), Height = (int)(200 * scale), Multiline = true, AcceptsReturn = true, Text = t.Note, Font = UITheme.GetFont(10.5f) };
                Button nb = new Button() { Text = "儲存", Left = (int)(285 * scale), Top = (int)(280 * scale), Width = (int)(100 * scale), Height = (int)(40 * scale), DialogResult = DialogResult.OK, FlatStyle = FlatStyle.Flat, BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite, Font = UITheme.GetFont(10f, FontStyle.Bold) };
                nb.FlatAppearance.BorderSize = 0;
                nf.Controls.AddRange(new Control[] { new Label() { Text = "【" + t.Name + "】", Left = (int)(15 * scale), Top = (int)(15 * scale), AutoSize = true, Font = UITheme.GetFont(11f, FontStyle.Bold) }, nt, nb });
                if(nf.ShowDialog() == DialogResult.OK) { t.Note = nt.Text; parentControl.UpdateTaskDb(t); RefreshData(); }
            };

            string typeTag = $"[{t.TaskType}] ";
            string timeInfo = t.MonthStr == "特定日期" ? $"[{t.DateStr} {t.TimeStr}]" : $"[{t.TimeStr}] {t.DateStr}";

            row.Controls.Add(bE, 0, 0); row.Controls.Add(bD, 1, 0); row.Controls.Add(bN, 2, 0);
            row.Controls.Add(new Label() { Text = typeTag + timeInfo + "  " + t.Name, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, AutoSize = true, Padding = new Padding(0,(int)(8 * scale),0,(int)(8 * scale)), Font = UITheme.GetFont(10.5f) }, 3, 0);
            inner.Controls.Add(row);
        }
        gb.Controls.Add(inner); flow.Controls.Add(gb);
    }
}
