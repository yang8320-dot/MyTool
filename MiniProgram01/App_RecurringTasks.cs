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
using ClosedXML.Excel; 

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

        Label lblTitle = new Label();
        lblTitle.Text = "週期任務";
        lblTitle.Font = UITheme.GetFont(12f, FontStyle.Bold);
        lblTitle.Dock = DockStyle.Fill;
        lblTitle.TextAlign = ContentAlignment.MiddleLeft;
        lblTitle.Padding = new Padding((int)(5 * scale), 0, 0, 0);
        lblTitle.ForeColor = UITheme.TextMain;

        Button btnViewAll = CreateHeaderButton("全部檢視", UITheme.CardWhite, UITheme.AppleBlue);
        btnViewAll.Click += (s, e) => { 
            AllTasksViewWindow win = new AllTasksViewWindow(this);
            win.Show(); 
        }; 

        Button btnAdd = CreateHeaderButton("新增任務", UITheme.AppleBlue, UITheme.CardWhite);
        btnAdd.Click += (s, e) => { 
            AddRecurringTaskWindow win = new AddRecurringTaskWindow(this);
            win.Show(); 
        };

        Button btnSet = CreateHeaderButton("設定", Color.Gainsboro, UITheme.TextMain);
        btnSet.Margin = new Padding((int)(2 * scale), (int)(8 * scale), (int)(8 * scale), (int)(8 * scale));
        btnSet.Click += (s, e) => { 
            RecurringSettingsWindow win = new RecurringSettingsWindow(this);
            win.Show(); 
        };

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
                taskPanel.SuspendLayout();
                foreach (Control c in taskPanel.Controls) {
                    if (c is Panel) {
                        c.Width = safeWidth;
                    }
                }
                taskPanel.ResumeLayout(true);
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
        Button btn = new Button();
        btn.Text = text;
        btn.Dock = DockStyle.Fill;
        btn.FlatStyle = FlatStyle.Flat;
        btn.Margin = new Padding((int)(2 * scale), (int)(8 * scale), (int)(2 * scale), (int)(8 * scale));
        btn.Cursor = Cursors.Hand;
        btn.BackColor = bg;
        btn.ForeColor = fg;
        btn.Font = UITheme.GetFont(10f, FontStyle.Bold);
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
        
        string advStr = DbHelper.GetSetting("Recur_AdvanceDays", "0");
        int parsedAdv = 0;
        int.TryParse(advStr, out parsedAdv);
        advanceDays = parsedAdv;
        
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
                        RecurringTask rt = new RecurringTask();
                        rt.Id = reader.GetInt32(0);
                        rt.Name = reader.GetString(1);
                        rt.MonthStr = reader.GetString(2);
                        rt.DateStr = reader.GetString(3);
                        rt.TimeStr = reader.GetString(4);
                        rt.TaskType = reader.GetString(5);
                        rt.Note = reader.IsDBNull(6) ? "" : reader.GetString(6);
                        rt.LastTriggeredDate = reader.IsDBNull(7) ? "" : reader.GetString(7);
                        rt.OrderIndex = reader.GetInt32(8);
                        
                        tasks.Add(rt);
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
            Panel card = new Panel();
            card.Width = startWidth;
            card.AutoSize = true;
            card.Margin = new Padding((int)(5 * scale), 0, (int)(5 * scale), (int)(3 * scale));
            card.BackColor = UITheme.CardWhite;

            card.Paint += (s, e) => {
                UITheme.DrawRoundedBackground(e.Graphics, new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale), UITheme.CardWhite);
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using (var pen = new Pen(Color.FromArgb(230, 230, 230), 1)) {
                    e.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(8 * scale)));
                }
            };

            TableLayoutPanel tlp = new TableLayoutPanel();
            tlp.Dock = DockStyle.Fill;
            tlp.ColumnCount = 4;
            tlp.RowCount = 1;
            tlp.AutoSize = true;
            tlp.Padding = new Padding((int)(8 * scale));
            tlp.BackColor = Color.Transparent;
            
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(40 * scale))); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(40 * scale))); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(40 * scale))); 
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            Button btnEdit = new Button();
            btnEdit.Text = "調";
            btnEdit.Dock = DockStyle.Top;
            btnEdit.Height = (int)(32 * scale);
            btnEdit.BackColor = UITheme.AppleBlue;
            btnEdit.ForeColor = UITheme.CardWhite;
            btnEdit.FlatStyle = FlatStyle.Flat;
            btnEdit.Font = UITheme.GetFont(9f, FontStyle.Bold);
            btnEdit.FlatAppearance.BorderSize = 0;
            btnEdit.Click += (s, e) => { 
                EditRecurringTaskWindow win = new EditRecurringTaskWindow(this, t);
                win.Show(); 
            };

            Button btnDel = new Button();
            btnDel.Text = "✕";
            btnDel.Dock = DockStyle.Top;
            btnDel.Height = (int)(32 * scale);
            btnDel.BackColor = UITheme.AppleRed;
            btnDel.ForeColor = UITheme.CardWhite;
            btnDel.FlatStyle = FlatStyle.Flat;
            btnDel.Font = UITheme.GetFont(9f, FontStyle.Bold);
            btnDel.FlatAppearance.BorderSize = 0;
            btnDel.Click += (s, e) => { 
                DialogResult res = MessageBox.Show("確定移除此週期任務？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (res == DialogResult.OK) { 
                    DeleteTask(t); 
                }
            };

            Button btnNote = new Button();
            btnNote.Text = "註";
            btnNote.Dock = DockStyle.Top;
            btnNote.Height = (int)(32 * scale);
            btnNote.FlatStyle = FlatStyle.Flat;
            btnNote.Font = UITheme.GetFont(9f, FontStyle.Bold);
            btnNote.FlatAppearance.BorderSize = 0;
            
            if (string.IsNullOrEmpty(t.Note)) {
                btnNote.BackColor = UITheme.BgGray;
                btnNote.ForeColor = UITheme.TextMain;
            } else {
                btnNote.BackColor = UITheme.AppleYellow;
                btnNote.ForeColor = Color.Black;
            }
            
            btnNote.Click += (s, e) => { 
                string n = ShowNoteEditBox(t.Name, t.Note); 
                if (n != null) { 
                    t.Note = n; 
                    UpdateTaskInDb(t); 
                    RefreshUI(); 
                } 
            };

            string typeTag = $"[{t.TaskType}] ";
            string timeInfo = t.MonthStr == "特定日期" ? $"[{t.DateStr} {t.TimeStr}]" : $"[{t.MonthStr} {t.DateStr} {t.TimeStr}]";

            Label lbl = new Label();
            lbl.Text = typeTag + timeInfo + " " + t.Name;
            lbl.Dock = DockStyle.Fill;
            lbl.TextAlign = ContentAlignment.MiddleLeft;
            lbl.AutoSize = true;
            lbl.Font = UITheme.GetFont(10.5f);
            lbl.ForeColor = UITheme.TextMain;
            lbl.Padding = new Padding((int)(5 * scale), 0, 0, 0);
            
            tlp.Controls.Add(btnEdit, 0, 0); 
            tlp.Controls.Add(btnDel, 1, 0); 
            tlp.Controls.Add(btnNote, 2, 0); 
            tlp.Controls.Add(lbl, 3, 0);
            
            card.Controls.Add(tlp); 
            taskPanel.Controls.Add(card);
        }
    }

    public void AddNewTask(string name, string month, string date, string time, string note, string type) {
        int orderIdx = 0;
        if (tasks.Count > 0) {
            orderIdx = tasks.Max(t => t.OrderIndex) + 1;
        }
        
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "INSERT INTO RecurringTasks (Name, MonthStr, DateStr, TimeStr, TaskType, Note, LastTriggeredDate, OrderIndex) VALUES (@N, @M, @D, @T, @Ty, @No, '', @O)";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@N", name); 
                cmd.Parameters.AddWithValue("@M", month);
                cmd.Parameters.AddWithValue("@D", date); 
                cmd.Parameters.AddWithValue("@T", time);
                cmd.Parameters.AddWithValue("@Ty", type); 
                cmd.Parameters.AddWithValue("@No", note);
                cmd.Parameters.AddWithValue("@O", orderIdx); 
                cmd.ExecuteNonQuery();
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
                        if (result != null) {
                            existingId = Convert.ToInt32(result);
                        }
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

    public void UpdateTaskDb(RecurringTask t) { 
        UpdateTaskInDb(t); 
        LoadTasksFromDb(); 
    }

    private void UpdateTaskInDb(RecurringTask t) {
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "UPDATE RecurringTasks SET Name=@N, MonthStr=@M, DateStr=@D, TimeStr=@T, TaskType=@Ty, Note=@No, LastTriggeredDate=@L, OrderIndex=@O WHERE Id=@Id";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@N", t.Name); 
                cmd.Parameters.AddWithValue("@M", t.MonthStr);
                cmd.Parameters.AddWithValue("@D", t.DateStr); 
                cmd.Parameters.AddWithValue("@T", t.TimeStr);
                cmd.Parameters.AddWithValue("@Ty", t.TaskType); 
                cmd.Parameters.AddWithValue("@No", t.Note);
                cmd.Parameters.AddWithValue("@L", t.LastTriggeredDate ?? ""); 
                cmd.Parameters.AddWithValue("@O", t.OrderIndex);
                cmd.Parameters.AddWithValue("@Id", t.Id); 
                cmd.ExecuteNonQuery();
            }
        }
    }

    public void DeleteTask(RecurringTask task) { 
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            using (var cmd = new SqliteCommand("DELETE FROM RecurringTasks WHERE Id=@Id", conn)) {
                cmd.Parameters.AddWithValue("@Id", task.Id); 
                cmd.ExecuteNonQuery();
            }
        }
        LoadTasksFromDb();
    }

    public void UpdateGlobalSettings(string dType, string dTime, int aDays, string sFreq) {
        digestType = dType; 
        digestTimeStr = dTime; 
        advanceDays = aDays; 
        scanFrequency = sFreq;
        
        checkTimer.Enabled = false; 
        checkTimer.Interval = GetTimerInterval(sFreq); 
        checkTimer.Enabled = true;
        
        SaveSettingsToDb(); 
        MessageBox.Show("設定儲存成功！");
    }

    private void CheckTasks() {
        DateTime now = DateTime.Now; 
        bool needsRefresh = false;
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
                        
                        t.LastTriggeredDate = targetDateStr; 
                        UpdateTaskInDb(t);
                        parentForm.AlertTab(1); 
                        
                        if (t.TaskType == "單次" || t.TaskType == "到期日") {
                            toRemove.Add(t);
                        }
                    }
                }
            }
        }

        if (toRemove.Count > 0) { 
            foreach (var r in toRemove) {
                DeleteTask(r); 
            }
            needsRefresh = true; 
        }
        
        if (needsRefresh) {
            this.Invoke(new Action(() => LoadTasksFromDb()));
        }

        if (digestType != "不提醒") {
            DateTime dtDigest;
            if (DateTime.TryParseExact(digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtDigest)) {
                DateTime targetDigest = new DateTime(now.Year, now.Month, now.Day, dtDigest.Hour, dtDigest.Minute, 0);
                bool shouldTrigger = false;
                
                if (digestType == "每週一" && now.DayOfWeek == DayOfWeek.Monday && now >= targetDigest) {
                    shouldTrigger = true;
                }
                if (digestType == "每月1號" && now.Day == 1 && now >= targetDigest) {
                    shouldTrigger = true;
                }
                
                string todayStr = now.ToString("yyyy-MM-dd");
                if (shouldTrigger && lastDigestDate != todayStr) { 
                    lastDigestDate = todayStr; 
                    SaveSettingsToDb();
                    AllTasksViewWindow win = new AllTasksViewWindow(this);
                    win.Show(); 
                }
            }
        }
    }

    private bool TryGetNextTriggerTime(RecurringTask t, DateTime now, out DateTime target) {
        target = now;
        try {
            string[] timeParts = t.TimeStr.Split(':');
            int h = int.Parse(timeParts[0]); 
            int m = int.Parse(timeParts[1]);

            if (t.MonthStr == "特定日期") {
                DateTime specificDate;
                if (DateTime.TryParseExact(t.DateStr, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out specificDate)) {
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
                Dictionary<string, DayOfWeek> dow = new Dictionary<string, DayOfWeek>();
                dow.Add("一", DayOfWeek.Monday);
                dow.Add("二", DayOfWeek.Tuesday);
                dow.Add("三", DayOfWeek.Wednesday);
                dow.Add("四", DayOfWeek.Thursday);
                dow.Add("五", DayOfWeek.Friday);
                dow.Add("六", DayOfWeek.Saturday);
                dow.Add("日", DayOfWeek.Sunday);
                
                if (!dow.ContainsKey(t.DateStr)) {
                    return false;
                }
                
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
        Form f = new Form();
        f.Width = (int)(420 * scale); 
        f.Height = (int)(380 * scale); 
        f.Text = "編輯備註"; 
        f.StartPosition = FormStartPosition.CenterScreen; 
        f.FormBorderStyle = FormBorderStyle.FixedDialog; 
        f.TopMost = true; 
        f.BackColor = UITheme.BgGray;

        Label lbl = new Label();
        lbl.Text = "【" + name + "】";
        lbl.Left = (int)(15 * scale);
        lbl.Top = (int)(15 * scale);
        lbl.AutoSize = true;
        lbl.Font = UITheme.GetFont(11f, FontStyle.Bold);
        
        TextBox txt = new TextBox();
        txt.Left = (int)(15 * scale);
        txt.Top = (int)(50 * scale);
        txt.Width = (int)(370 * scale);
        txt.Height = (int)(200 * scale);
        txt.Multiline = true;
        txt.AcceptsReturn = true;
        txt.Text = current;
        txt.Font = UITheme.GetFont(10.5f);
        
        Button btn = new Button();
        btn.Text = "儲存";
        btn.Left = (int)(285 * scale);
        btn.Top = (int)(280 * scale);
        btn.Width = (int)(100 * scale);
        btn.Height = (int)(40 * scale);
        btn.DialogResult = DialogResult.OK;
        btn.FlatStyle = FlatStyle.Flat;
        btn.BackColor = UITheme.AppleBlue;
        btn.ForeColor = UITheme.CardWhite;
        btn.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btn.FlatAppearance.BorderSize = 0;

        f.Controls.Add(lbl);
        f.Controls.Add(txt);
        f.Controls.Add(btn);

        DialogResult dr = f.ShowDialog();
        if (dr == DialogResult.OK) {
            return txt.Text;
        } else {
            return null;
        }
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
        this.Text = "新增任務"; 
        this.Width = (int)(420 * scale); 
        this.Height = (int)(680 * scale); 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true; 
        this.BackColor = UITheme.BgGray;

        FlowLayoutPanel f = new FlowLayoutPanel();
        f.Dock = DockStyle.Fill;
        f.FlowDirection = FlowDirection.TopDown;
        f.Padding = new Padding((int)(25 * scale));

        Label l1 = new Label();
        l1.Text = "任務名稱：";
        l1.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(l1); 

        txtN = new TextBox();
        txtN.Width = (int)(340 * scale);
        txtN.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(txtN);

        Label l2 = new Label();
        l2.Text = "詳細說明 (註)：";
        l2.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        l2.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(l2);

        txtNote = new TextBox();
        txtNote.Width = (int)(340 * scale);
        txtNote.Height = (int)(80 * scale);
        txtNote.Multiline = true;
        txtNote.AcceptsReturn = true;
        txtNote.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(txtNote);
        
        Label l3 = new Label();
        l3.Text = "任務類型：";
        l3.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        l3.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(l3);

        cmType = new ComboBox();
        cmType.Width = (int)(340 * scale);
        cmType.DropDownStyle = ComboBoxStyle.DropDownList;
        cmType.Font = UITheme.GetFont(10.5f);
        cmType.Items.Add("循環");
        cmType.Items.Add("單次");
        cmType.Items.Add("到期日");
        f.Controls.Add(cmType);

        lblCycle = new Label();
        lblCycle.Text = "週期類型：";
        lblCycle.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        lblCycle.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(lblCycle);
        
        cmM = new ComboBox();
        cmM.Width = (int)(340 * scale);
        cmM.DropDownStyle = ComboBoxStyle.DropDownList;
        cmM.Font = UITheme.GetFont(10.5f);
        cmM.Items.Add("每天");
        cmM.Items.Add("每週");
        cmM.Items.Add("每月");
        for(int i = 1; i <= 12; i++) {
            cmM.Items.Add(i.ToString() + "月");
        }
        f.Controls.Add(cmM); 
        
        cmD = new ComboBox();
        cmD.Width = (int)(340 * scale);
        cmD.DropDownStyle = ComboBoxStyle.DropDownList;
        cmD.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(cmD);

        lblDate = new Label();
        lblDate.Text = "指定日期：";
        lblDate.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        lblDate.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(lblDate);

        dtpDate = new DateTimePicker();
        dtpDate.Width = (int)(340 * scale);
        dtpDate.Format = DateTimePickerFormat.Custom;
        dtpDate.CustomFormat = "yyyy-MM-dd";
        dtpDate.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(dtpDate);

        cmType.SelectedIndexChanged += (s, e) => {
            bool isLoop = false;
            if (cmType.Text == "循環") {
                isLoop = true;
            }
            lblCycle.Visible = isLoop;
            cmM.Visible = isLoop;
            cmD.Visible = isLoop;
            
            lblDate.Visible = !isLoop;
            dtpDate.Visible = !isLoop;
        };
        cmType.SelectedIndex = 0; 

        cmM.SelectedIndexChanged += (s, e) => {
            cmD.Items.Clear();
            if(cmM.Text == "每天") { 
                cmD.Items.Add("每日"); 
                cmD.Enabled = false; 
            } else if(cmM.Text == "每週") { 
                cmD.Items.Add("一");
                cmD.Items.Add("二");
                cmD.Items.Add("三");
                cmD.Items.Add("四");
                cmD.Items.Add("五");
                cmD.Items.Add("六");
                cmD.Items.Add("日");
                cmD.Enabled = true; 
            } else { 
                for(int i = 1; i <= 31; i++) {
                    cmD.Items.Add(i.ToString());
                }
                cmD.Items.Add("月底"); 
                cmD.Enabled = true; 
            }
            if (cmD.Items.Count > 0) {
                cmD.SelectedIndex = 0;
            }
        }; 
        cmM.SelectedIndex = 0;

        Label l4 = new Label();
        l4.Text = "觸發時間：";
        l4.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        l4.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(l4);

        dtpTime = new DateTimePicker();
        dtpTime.Width = (int)(340 * scale);
        dtpTime.Format = DateTimePickerFormat.Custom;
        dtpTime.CustomFormat = "HH:mm";
        dtpTime.ShowUpDown = true;
        dtpTime.Value = DateTime.Today.AddHours(9);
        dtpTime.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(dtpTime);

        Button btn = new Button();
        btn.Text = "建立任務";
        btn.Width = (int)(340 * scale);
        btn.Height = (int)(45 * scale);
        btn.BackColor = UITheme.AppleBlue;
        btn.ForeColor = UITheme.CardWhite;
        btn.FlatStyle = FlatStyle.Flat;
        btn.Margin = new Padding(0, (int)(25 * scale), 0, 0);
        btn.Font = UITheme.GetFont(11f, FontStyle.Bold);
        btn.Cursor = Cursors.Hand;
        btn.FlatAppearance.BorderSize = 0;

        btn.Click += (s, e) => { 
            if(!string.IsNullOrWhiteSpace(txtN.Text)) { 
                string monthVal = "特定日期";
                if (cmType.Text == "循環") {
                    monthVal = cmM.Text;
                }
                
                string dateVal = dtpDate.Value.ToString("yyyy-MM-dd");
                if (cmType.Text == "循環") {
                    dateVal = cmD.Text;
                }
                
                string timeVal = dtpTime.Value.ToString("HH:mm");
                
                parent.AddNewTask(txtN.Text, monthVal, dateVal, timeVal, txtNote.Text, cmType.Text); 
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
        this.parent = p; 
        this.task = t; 
        float scale = this.DeviceDpi / 96f;
        this.Text = "調整任務"; 
        this.Width = (int)(420 * scale); 
        this.Height = (int)(680 * scale); 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true; 
        this.BackColor = UITheme.BgGray;

        FlowLayoutPanel f = new FlowLayoutPanel();
        f.Dock = DockStyle.Fill;
        f.FlowDirection = FlowDirection.TopDown;
        f.Padding = new Padding((int)(25 * scale));

        Label l1 = new Label();
        l1.Text = "任務名稱：";
        l1.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(l1); 

        txtN = new TextBox();
        txtN.Width = (int)(340 * scale);
        txtN.Text = t.Name;
        txtN.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(txtN);

        Label l2 = new Label();
        l2.Text = "詳細說明 (註)：";
        l2.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        l2.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(l2);

        txtNote = new TextBox();
        txtNote.Width = (int)(340 * scale);
        txtNote.Height = (int)(80 * scale);
        txtNote.Multiline = true;
        txtNote.AcceptsReturn = true;
        txtNote.Text = t.Note;
        txtNote.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(txtNote);
        
        Label l3 = new Label();
        l3.Text = "任務類型：";
        l3.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        l3.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(l3);

        cmType = new ComboBox();
        cmType.Width = (int)(340 * scale);
        cmType.DropDownStyle = ComboBoxStyle.DropDownList;
        cmType.Font = UITheme.GetFont(10.5f);
        cmType.Items.Add("循環");
        cmType.Items.Add("單次");
        cmType.Items.Add("到期日");
        cmType.Text = t.TaskType; 
        f.Controls.Add(cmType);

        lblCycle = new Label();
        lblCycle.Text = "週期類型：";
        lblCycle.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        lblCycle.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(lblCycle);
        
        cmM = new ComboBox();
        cmM.Width = (int)(340 * scale);
        cmM.DropDownStyle = ComboBoxStyle.DropDownList;
        cmM.Font = UITheme.GetFont(10.5f);
        cmM.Items.Add("每天");
        cmM.Items.Add("每週");
        cmM.Items.Add("每月");
        for(int k = 1; k <= 12; k++) {
            cmM.Items.Add(k.ToString() + "月");
        }
        f.Controls.Add(cmM);

        cmD = new ComboBox();
        cmD.Width = (int)(340 * scale);
        cmD.DropDownStyle = ComboBoxStyle.DropDownList;
        cmD.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(cmD);

        lblDate = new Label();
        lblDate.Text = "指定日期：";
        lblDate.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        lblDate.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(lblDate);
        
        dtpDate = new DateTimePicker();
        dtpDate.Width = (int)(340 * scale);
        dtpDate.Format = DateTimePickerFormat.Custom;
        dtpDate.CustomFormat = "yyyy-MM-dd";
        dtpDate.Font = UITheme.GetFont(10.5f);
        f.Controls.Add(dtpDate);

        cmType.SelectedIndexChanged += (s, e) => {
            bool isLoop = false;
            if (cmType.Text == "循環") {
                isLoop = true;
            }
            lblCycle.Visible = isLoop;
            cmM.Visible = isLoop;
            cmD.Visible = isLoop;
            
            lblDate.Visible = !isLoop;
            dtpDate.Visible = !isLoop;
        };

        if (t.MonthStr == "特定日期") {
            DateTime d;
            if (DateTime.TryParseExact(t.DateStr, "yyyy-MM-dd", null, System.Globalization.DateTimeStyles.None, out d)) {
                dtpDate.Value = d;
            }
        } else {
            cmM.Text = t.MonthStr;
        }

        cmM.SelectedIndexChanged += (s, e) => {
            cmD.Items.Clear();
            if(cmM.Text == "每天") { 
                cmD.Items.Add("每日"); 
            } else if(cmM.Text == "每週") { 
                cmD.Items.Add("一");
                cmD.Items.Add("二");
                cmD.Items.Add("三");
                cmD.Items.Add("四");
                cmD.Items.Add("五");
                cmD.Items.Add("六");
                cmD.Items.Add("日");
            } else { 
                for(int k = 1; k <= 31; k++) {
                    cmD.Items.Add(k.ToString());
                }
                cmD.Items.Add("月底"); 
            }
            if(cmD.Items.Count > 0) {
                cmD.SelectedIndex = 0;
            }
        }; 
        
        if (t.MonthStr != "特定日期") {
            cmD.Text = t.DateStr;
        }

        Label l4 = new Label();
        l4.Text = "觸發時間：";
        l4.Margin = new Padding(0, (int)(15 * scale), 0, 0);
        l4.Font = UITheme.GetFont(10f, FontStyle.Bold);
        f.Controls.Add(l4);

        dtpTime = new DateTimePicker();
        dtpTime.Width = (int)(340 * scale);
        dtpTime.Format = DateTimePickerFormat.Custom;
        dtpTime.CustomFormat = "HH:mm";
        dtpTime.ShowUpDown = true;
        dtpTime.Font = UITheme.GetFont(10.5f);
        
        DateTime dtv;
        if(DateTime.TryParseExact(t.TimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtv)) {
            dtpTime.Value = dtv;
        }
        f.Controls.Add(dtpTime);

        Button btn = new Button();
        btn.Text = "儲存修改";
        btn.Width = (int)(340 * scale);
        btn.Height = (int)(45 * scale);
        btn.BackColor = UITheme.AppleGreen;
        btn.ForeColor = UITheme.CardWhite;
        btn.FlatStyle = FlatStyle.Flat;
        btn.Margin = new Padding(0, (int)(25 * scale), 0, 0);
        btn.Font = UITheme.GetFont(11f, FontStyle.Bold);
        btn.Cursor = Cursors.Hand;
        btn.FlatAppearance.BorderSize = 0;

        btn.Click += (s, e) => { 
            t.Name = txtN.Text;
            t.TaskType = cmType.Text;
            
            string monthVal = "特定日期";
            if (cmType.Text == "循環") {
                monthVal = cmM.Text;
            }
            t.MonthStr = monthVal;
            
            string dateVal = dtpDate.Value.ToString("yyyy-MM-dd");
            if (cmType.Text == "循環") {
                dateVal = cmD.Text;
            }
            t.DateStr = dateVal;
            
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
        this.Text = "全域排程設定"; 
        this.Width = (int)(380 * scale); 
        this.Height = (int)(380 * scale); 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.TopMost = true; 
        this.BackColor = UITheme.BgGray;

        FlowLayoutPanel f = new FlowLayoutPanel();
        f.Dock = DockStyle.Fill;
        f.FlowDirection = FlowDirection.TopDown;
        f.Padding = new Padding((int)(20 * scale));
        
        FlowLayoutPanel r1 = new FlowLayoutPanel();
        r1.AutoSize = true;

        Label l1 = new Label();
        l1.Text = "所有任務提前";
        l1.AutoSize = true;
        l1.Margin = new Padding(0, (int)(5 * scale), 0, 0);
        l1.Font = UITheme.GetFont(10.5f);
        r1.Controls.Add(l1);

        cmAdv = new ComboBox();
        cmAdv.Width = (int)(70 * scale);
        cmAdv.DropDownStyle = ComboBoxStyle.DropDownList;
        cmAdv.Font = UITheme.GetFont(10.5f);
        for (int i = 0; i <= 7; i++) {
            cmAdv.Items.Add(i.ToString());
        }
        cmAdv.Text = p.advanceDays.ToString(); 
        r1.Controls.Add(cmAdv); 

        Label l2 = new Label();
        l2.Text = "天加入待辦";
        l2.AutoSize = true;
        l2.Margin = new Padding(0, (int)(5 * scale), 0, 0);
        l2.Font = UITheme.GetFont(10.5f);
        r1.Controls.Add(l2);

        f.Controls.Add(r1);
        
        FlowLayoutPanel r2 = new FlowLayoutPanel();
        r2.AutoSize = true;
        r2.Margin = new Padding(0, (int)(15 * scale), 0, (int)(15 * scale));

        Label l3 = new Label();
        l3.Text = "視窗摘要提醒：";
        l3.AutoSize = true;
        l3.Margin = new Padding(0, (int)(5 * scale), 0, 0);
        l3.Font = UITheme.GetFont(10.5f);
        r2.Controls.Add(l3);

        cmDig = new ComboBox();
        cmDig.Width = (int)(100 * scale);
        cmDig.DropDownStyle = ComboBoxStyle.DropDownList;
        cmDig.Font = UITheme.GetFont(10.5f);
        cmDig.Items.Add("不提醒"); 
        cmDig.Items.Add("每週一"); 
        cmDig.Items.Add("每月1號");
        cmDig.Text = p.digestType;
        r2.Controls.Add(cmDig);

        dtp = new DateTimePicker();
        dtp.Width = (int)(80 * scale);
        dtp.Format = DateTimePickerFormat.Custom;
        dtp.CustomFormat = "HH:mm";
        dtp.ShowUpDown = true;
        dtp.Font = UITheme.GetFont(10.5f);
        
        DateTime dtv2;
        if(DateTime.TryParseExact(p.digestTimeStr, "HH:mm", null, System.Globalization.DateTimeStyles.None, out dtv2)) {
            dtp.Value = dtv2;
        }
        r2.Controls.Add(dtp); 
        f.Controls.Add(r2);
        
        Label line = new Label();
        line.AutoSize = false;
        line.Height = 2;
        line.Width = (int)(320 * scale);
        line.BorderStyle = BorderStyle.Fixed3D;
        line.Margin = new Padding(0, (int)(5 * scale), 0, (int)(20 * scale));
        f.Controls.Add(line);

        FlowLayoutPanel r3 = new FlowLayoutPanel();
        r3.AutoSize = true;
        r3.Margin = new Padding(0, 0, 0, (int)(25 * scale));

        Label l4 = new Label();
        l4.Text = "背景掃描頻率：";
        l4.AutoSize = true;
        l4.Margin = new Padding(0, (int)(5 * scale), 0, 0);
        l4.Font = UITheme.GetFont(10.5f);
        r3.Controls.Add(l4);

        cmScan = new ComboBox();
        cmScan.Width = (int)(120 * scale);
        cmScan.DropDownStyle = ComboBoxStyle.DropDownList;
        cmScan.Font = UITheme.GetFont(10.5f);
        cmScan.Items.Add("即時");
        cmScan.Items.Add("1分鐘");
        cmScan.Items.Add("5分鐘");
        cmScan.Items.Add("10分鐘");
        cmScan.Items.Add("1小時");
        cmScan.Items.Add("12小時");
        cmScan.Items.Add("1天");
        cmScan.Text = p.scanFrequency;
        r3.Controls.Add(cmScan); 

        f.Controls.Add(r3);

        Button btn = new Button();
        btn.Text = "儲存所有設定";
        btn.Width = (int)(320 * scale);
        btn.Height = (int)(45 * scale);
        btn.BackColor = UITheme.AppleBlue;
        btn.ForeColor = UITheme.CardWhite;
        btn.FlatStyle = FlatStyle.Flat;
        btn.Font = UITheme.GetFont(11f, FontStyle.Bold);
        btn.Cursor = Cursors.Hand;
        btn.FlatAppearance.BorderSize = 0;
        
        btn.Click += (s, e) => { 
            int advDays = 0;
            int.TryParse(cmAdv.Text, out advDays);
            p.UpdateGlobalSettings(cmDig.Text, dtp.Value.ToString("HH:mm"), advDays, cmScan.Text); 
            this.Close(); 
        };
        f.Controls.Add(btn); 
        this.Controls.Add(f);
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

        TableLayoutPanel header = new TableLayoutPanel();
        header.Dock = DockStyle.Top;
        header.Height = (int)(70 * scale);
        header.BackColor = UITheme.CardWhite;
        header.ColumnCount = 4;
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(110 * scale)));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(110 * scale)));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(110 * scale)));

        Label lbl = new Label();
        lbl.Text = "週期任務排程總覽";
        lbl.Dock = DockStyle.Fill;
        lbl.TextAlign = ContentAlignment.MiddleLeft;
        lbl.Padding = new Padding((int)(20 * scale),0,0,0);
        lbl.Font = UITheme.GetFont(16f, FontStyle.Bold);
        lbl.ForeColor = UITheme.TextMain;
        
        Button btnImport = new Button();
        btnImport.Text = "匯入 Excel";
        btnImport.Dock = DockStyle.Fill;
        btnImport.BackColor = UITheme.AppleBlue;
        btnImport.ForeColor = UITheme.CardWhite;
        btnImport.FlatStyle = FlatStyle.Flat;
        btnImport.Margin = new Padding((int)(5*scale), (int)(15*scale), (int)(5*scale), (int)(15*scale));
        btnImport.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnImport.Cursor = Cursors.Hand;
        btnImport.FlatAppearance.BorderSize = 0;
        btnImport.Click += (s, e) => ExecuteImportExcel();

        Button btnExport = new Button();
        btnExport.Text = "匯出 Excel";
        btnExport.Dock = DockStyle.Fill;
        btnExport.BackColor = UITheme.AppleGreen;
        btnExport.ForeColor = UITheme.CardWhite;
        btnExport.FlatStyle = FlatStyle.Flat;
        btnExport.Margin = new Padding((int)(5*scale), (int)(15*scale), (int)(5*scale), (int)(15*scale));
        btnExport.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnExport.Cursor = Cursors.Hand;
        btnExport.FlatAppearance.BorderSize = 0;
        btnExport.Click += (s, e) => ExecuteExportExcel();

        Button btnPrint = new Button();
        btnPrint.Text = "導出 PDF";
        btnPrint.Dock = DockStyle.Fill;
        btnPrint.BackColor = Color.Gray;
        btnPrint.ForeColor = UITheme.CardWhite;
        btnPrint.FlatStyle = FlatStyle.Flat;
        btnPrint.Margin = new Padding((int)(5*scale), (int)(15*scale), (int)(15*scale), (int)(15*scale));
        btnPrint.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        btnPrint.Cursor = Cursors.Hand;
        btnPrint.FlatAppearance.BorderSize = 0;
        btnPrint.Click += (s, e) => ExecuteExportPDF();

        header.Controls.Add(lbl, 0, 0); 
        header.Controls.Add(btnImport, 1, 0); 
        header.Controls.Add(btnExport, 2, 0); 
        header.Controls.Add(btnPrint, 3, 0); 
        this.Controls.Add(header);

        flow = new FlowLayoutPanel();
        flow.Dock = DockStyle.Fill;
        flow.AutoScroll = true;
        flow.Padding = new Padding((int)(20 * scale));
        flow.FlowDirection = FlowDirection.TopDown;
        flow.WrapContents = false;
        
        flow.Resize += (s, e) => { 
            int w = flow.ClientSize.Width - (int)(40 * scale); 
            if (w > 0) {
                flow.SuspendLayout(); 
                foreach (Control c in flow.Controls) {
                    if (c is Panel) {
                        c.Width = w;
                    }
                }
                flow.ResumeLayout(true);
            }
        };
        
        this.Controls.Add(flow);
        flow.BringToFront();
        RefreshData();
    }

    private void ExecuteExportExcel() {
        using (SaveFileDialog sfd = new SaveFileDialog()) {
            sfd.Filter = "Excel 活頁簿|*.xlsx";
            sfd.FileName = $"週期任務總覽_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            if (sfd.ShowDialog() == DialogResult.OK) {
                try {
                    using (var workbook = new XLWorkbook()) {
                        var mainSheet = workbook.Worksheets.Add("週期任務清單");
                        var dataSheet = workbook.Worksheets.Add("DataValidation");
                        dataSheet.Hide(); 

                        List<string> times = new List<string>();
                        for (int h = 0; h < 24; h++) {
                            times.Add($"{h:D2}:00");
                            times.Add($"{h:D2}:30");
                        }
                        for (int i = 0; i < times.Count; i++) {
                            dataSheet.Cell(i + 1, 1).Value = times[i];
                        }

                        string[] dateArr = "每日,一,二,三,四,五,六,日,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,月底".Split(',');
                        for (int i = 0; i < dateArr.Length; i++) {
                            dataSheet.Cell(i + 1, 2).Value = dateArr[i];
                        }

                        string[] typeArr = {"循環", "單次", "到期日"};
                        for (int i = 0; i < typeArr.Length; i++) {
                            dataSheet.Cell(i + 1, 3).Value = typeArr[i];
                        }

                        string[] cycleArr = "每天,每週,每月,1月,2月,3月,4月,5月,6月,7月,8月,9月,10月,11月,12月,特定日期".Split(',');
                        for (int i = 0; i < cycleArr.Length; i++) {
                            dataSheet.Cell(i + 1, 4).Value = cycleArr[i];
                        }

                        mainSheet.Cell(1, 1).Value = "任務名稱";
                        mainSheet.Cell(1, 2).Value = "任務類型";
                        mainSheet.Cell(1, 3).Value = "週期類型";
                        mainSheet.Cell(1, 4).Value = "指定日期";
                        mainSheet.Cell(1, 5).Value = "觸發時間";
                        mainSheet.Cell(1, 6).Value = "備註";
                        
                        var headerRange = mainSheet.Range("A1:F1");
                        headerRange.Style.Font.Bold = true;
                        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                        int row = 2;
                        foreach (var t in parentControl.tasks) {
                            mainSheet.Cell(row, 1).Value = t.Name;
                            mainSheet.Cell(row, 2).Value = t.TaskType;
                            mainSheet.Cell(row, 3).Value = t.MonthStr;
                            mainSheet.Cell(row, 4).Value = t.DateStr;
                            mainSheet.Cell(row, 5).SetValue("'" + t.TimeStr); 
                            mainSheet.Cell(row, 6).Value = t.Note;
                            row++;
                        }

                        mainSheet.Column(1).Width = 35; 
                        for (int i = 2; i <= 6; i++) {
                            mainSheet.Column(i).Width = 15;
                        }
                        mainSheet.Rows().Height = 22;

                        int maxValRow = row > 100 ? row + 100 : 500;

                        // 【完美修復】全部寫入隱藏工作表，使用標準 Excel 公式，並嚴格加上 XLAllowedValues.List 防止 XML 損毀
                        var valB = mainSheet.Range($"B2:B{maxValRow}").CreateDataValidation();
                        valB.AllowedValues = XLAllowedValues.List;
                        valB.List($"=DataValidation!$C$1:$C${typeArr.Length}");
                        valB.InCellDropdown = true;
                        valB.ShowErrorMessage = false;

                        var valC = mainSheet.Range($"C2:C{maxValRow}").CreateDataValidation();
                        valC.AllowedValues = XLAllowedValues.List;
                        valC.List($"=DataValidation!$D$1:$D${cycleArr.Length}");
                        valC.InCellDropdown = true;
                        valC.ShowErrorMessage = false;

                        var valD = mainSheet.Range($"D2:D{maxValRow}").CreateDataValidation();
                        valD.AllowedValues = XLAllowedValues.List;
                        valD.List($"=DataValidation!$B$1:$B${dateArr.Length}");
                        valD.InCellDropdown = true;
                        valD.ShowErrorMessage = false;

                        var valE = mainSheet.Range($"E2:E{maxValRow}").CreateDataValidation();
                        valE.AllowedValues = XLAllowedValues.List;
                        valE.List($"=DataValidation!$A$1:$A${times.Count}");
                        valE.InCellDropdown = true;
                        valE.ShowErrorMessage = false;

                        mainSheet.Activate();
                        workbook.SaveAs(sfd.FileName);
                        MessageBox.Show("Excel 檔案已成功導出！\n\n(已在B~E欄自動建立快速下拉選單)", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                } catch (Exception ex) {
                    MessageBox.Show("匯出時發生錯誤：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    private void ExecuteImportExcel() {
        using (OpenFileDialog ofd = new OpenFileDialog()) {
            ofd.Filter = "Excel 活頁簿|*.xlsx";
            ofd.Title = "選擇要匯入的排程檔案";
            if (ofd.ShowDialog() == DialogResult.OK) {
                try {
                    using (var workbook = new XLWorkbook(ofd.FileName)) {
                        var sheet = workbook.Worksheets.First();
                        var rows = sheet.RangeUsed().RowsUsed().Skip(1);

                        List<App_RecurringTasks.RecurringTask> importList = new List<App_RecurringTasks.RecurringTask>();

                        foreach (var r in rows) {
                            string name = r.Cell(1).GetString();
                            if (string.IsNullOrWhiteSpace(name)) continue;

                            App_RecurringTasks.RecurringTask t = new App_RecurringTasks.RecurringTask();
                            t.Name = name;
                            t.TaskType = r.Cell(2).GetString();
                            t.MonthStr = r.Cell(3).GetString();
                            t.DateStr = r.Cell(4).GetString();
                            t.TimeStr = r.Cell(5).GetString().TrimStart('\'');
                            t.Note = r.Cell(6).GetString();
                            
                            importList.Add(t);
                        }

                        var resultStats = parentControl.BulkImportOrUpdate(importList);

                        MessageBox.Show(
                            $"Excel 匯入完成！\n\n新增了 {resultStats.Item1} 筆任務\n更新了 {resultStats.Item2} 筆任務", 
                            "匯入成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                        RefreshData();
                    }
                } catch (Exception ex) {
                    MessageBox.Show("檔案格式有誤或被其他程式鎖定。\n詳細錯誤：" + ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

                List<App_RecurringTasks.RecurringTask> dailyList = tasks.Where(t => t.MonthStr == "每天").ToList();
                addGroup("每天觸發", dailyList);
                
                List<App_RecurringTasks.RecurringTask> weeklyList = tasks.Where(t => t.MonthStr == "每週").ToList();
                addGroup("每週觸發", weeklyList);
                
                List<App_RecurringTasks.RecurringTask> monthlyList = tasks.Where(t => t.MonthStr == "每月").ToList();
                addGroup("每月觸發", monthlyList);
                
                for (int i = 1; i <= 12; i++) {
                    string monthTarget = i.ToString() + "月";
                    List<App_RecurringTasks.RecurringTask> monthSpecList = tasks.Where(t => t.MonthStr == monthTarget).ToList();
                    addGroup(i.ToString() + "月 限定", monthSpecList);
                }
                
                List<App_RecurringTasks.RecurringTask> specificList = tasks.Where(t => t.MonthStr == "特定日期").ToList();
                addGroup("特定日期 (單次/到期日)", specificList);

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
            
            List<App_RecurringTasks.RecurringTask> dailyList = tasks.Where(t => t.MonthStr == "每天").ToList();
            AddGroup("每天觸發", dailyList);
            
            List<App_RecurringTasks.RecurringTask> weeklyList = tasks.Where(t => t.MonthStr == "每週").ToList();
            AddGroup("每週觸發", weeklyList);
            
            List<App_RecurringTasks.RecurringTask> monthlyList = tasks.Where(t => t.MonthStr == "每月").ToList();
            AddGroup("每月觸發", monthlyList);
            
            for (int i = 1; i <= 12; i++) {
                string monthTarget = i.ToString() + "月";
                List<App_RecurringTasks.RecurringTask> monthSpecList = tasks.Where(t => t.MonthStr == monthTarget).ToList();
                AddGroup(i.ToString() + "月 限定", monthSpecList);
            }
            
            List<App_RecurringTasks.RecurringTask> specificList = tasks.Where(t => t.MonthStr == "特定日期").ToList();
            AddGroup("特定日期 (單次/到期日)", specificList);
            
        } finally {
            flow.ResumeLayout();
            isRefreshing = false;
        }
    }

    private void AddGroup(string header, List<App_RecurringTasks.RecurringTask> sub) {
        if (sub.Count == 0) return;

        Panel gb = new Panel();
        gb.AutoSize = true;
        gb.Width = flow.ClientSize.Width - (int)(40 * scale);
        gb.Margin = new Padding((int)(10 * scale), (int)(10 * scale), (int)(10 * scale), (int)(25 * scale));
        gb.Padding = new Padding((int)(15 * scale));
        gb.BackColor = UITheme.CardWhite;
        
        gb.Paint += (s, e) => {
            UITheme.DrawRoundedBackground(e.Graphics, new Rectangle(0, 0, gb.Width - 1, gb.Height - 1), (int)(12 * scale), UITheme.CardWhite);
            using (var pen = new Pen(Color.FromArgb(220, 220, 220), 1)) {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, gb.Width - 1, gb.Height - 1), (int)(12 * scale)));
            }
        };

        FlowLayoutPanel inner = new FlowLayoutPanel();
        inner.Dock = DockStyle.Fill;
        inner.FlowDirection = FlowDirection.TopDown;
        inner.WrapContents = false;
        inner.AutoSize = true;
        inner.BackColor = Color.Transparent;
        
        Label titleLbl = new Label();
        titleLbl.Text = header;
        titleLbl.Font = UITheme.GetFont(13f, FontStyle.Bold);
        titleLbl.ForeColor = UITheme.AppleBlue;
        titleLbl.AutoSize = true;
        titleLbl.Margin = new Padding(0, 0, 0, (int)(15 * scale));
        
        inner.Controls.Add(titleLbl);

        foreach (var t in sub) {
            TableLayoutPanel row = new TableLayoutPanel();
            row.Width = gb.Width - (int)(40 * scale);
            row.AutoSize = true;
            row.ColumnCount = 4;
            row.Margin = new Padding(0, 0, 0, (int)(10 * scale));
            
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(45 * scale))); 
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(45 * scale)));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(45 * scale))); 
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            Button bE = new Button();
            bE.Text = "調";
            bE.Height = (int)(32 * scale);
            bE.Dock = DockStyle.Top;
            bE.BackColor = UITheme.AppleBlue;
            bE.ForeColor = UITheme.CardWhite;
            bE.FlatStyle = FlatStyle.Flat;
            bE.Font = UITheme.GetFont(9f, FontStyle.Bold);
            bE.Cursor = Cursors.Hand;
            bE.FlatAppearance.BorderSize = 0;
            bE.Click += (s, e) => { 
                EditRecurringTaskWindow win = new EditRecurringTaskWindow(parentControl, t);
                win.ShowDialog(); 
                RefreshData(); 
            };

            Button bD = new Button();
            bD.Text = "✕";
            bD.Height = (int)(32 * scale);
            bD.Dock = DockStyle.Top;
            bD.BackColor = UITheme.AppleRed;
            bD.ForeColor = UITheme.CardWhite;
            bD.FlatStyle = FlatStyle.Flat;
            bD.Font = UITheme.GetFont(9f, FontStyle.Bold);
            bD.Cursor = Cursors.Hand;
            bD.FlatAppearance.BorderSize = 0;
            bD.Click += (s, e) => { 
                DialogResult dr = MessageBox.Show("確定移除？", "確認", MessageBoxButtons.OKCancel);
                if (dr == DialogResult.OK) { 
                    parentControl.DeleteTask(t); 
                    RefreshData(); 
                } 
            };

            Button bN = new Button();
            bN.Text = "註";
            bN.Height = (int)(32 * scale);
            bN.Dock = DockStyle.Top;
            bN.FlatStyle = FlatStyle.Flat;
            bN.Font = UITheme.GetFont(9f, FontStyle.Bold);
            bN.Cursor = Cursors.Hand;
            bN.FlatAppearance.BorderSize = 0;
            
            if (!string.IsNullOrEmpty(t.Note)) { 
                bN.BackColor = UITheme.AppleYellow; 
                bN.ForeColor = Color.Black; 
            } else { 
                bN.BackColor = UITheme.BgGray; 
                bN.ForeColor = UITheme.TextSub; 
            }
            
            bN.Click += (s, e) => { 
                Form nf = new Form();
                nf.Width = (int)(420 * scale);
                nf.Height = (int)(380 * scale);
                nf.Text = "任務備註";
                nf.StartPosition = FormStartPosition.CenterScreen;
                nf.TopMost = true;
                nf.BackColor = UITheme.BgGray;

                TextBox nt = new TextBox();
                nt.Left = (int)(15 * scale);
                nt.Top = (int)(50 * scale);
                nt.Width = (int)(370 * scale);
                nt.Height = (int)(200 * scale);
                nt.Multiline = true;
                nt.AcceptsReturn = true;
                nt.Text = t.Note;
                nt.Font = UITheme.GetFont(10.5f);

                Button nb = new Button();
                nb.Text = "儲存";
                nb.Left = (int)(285 * scale);
                nb.Top = (int)(280 * scale);
                nb.Width = (int)(100 * scale);
                nb.Height = (int)(40 * scale);
                nb.DialogResult = DialogResult.OK;
                nb.FlatStyle = FlatStyle.Flat;
                nb.BackColor = UITheme.AppleBlue;
                nb.ForeColor = UITheme.CardWhite;
                nb.Font = UITheme.GetFont(10f, FontStyle.Bold);
                nb.FlatAppearance.BorderSize = 0;

                Label nLbl = new Label();
                nLbl.Text = "【" + t.Name + "】";
                nLbl.Left = (int)(15 * scale);
                nLbl.Top = (int)(15 * scale);
                nLbl.AutoSize = true;
                nLbl.Font = UITheme.GetFont(11f, FontStyle.Bold);

                nf.Controls.Add(nLbl);
                nf.Controls.Add(nt);
                nf.Controls.Add(nb);

                DialogResult res = nf.ShowDialog();
                if (res == DialogResult.OK) { 
                    t.Note = nt.Text; 
                    parentControl.UpdateTaskDb(t); 
                    RefreshData(); 
                }
            };

            string typeTag = $"[{t.TaskType}] ";
            string timeInfo = t.MonthStr == "特定日期" ? $"[{t.DateStr} {t.TimeStr}]" : $"[{t.TimeStr}] {t.DateStr}";

            Label rowLbl = new Label();
            rowLbl.Text = typeTag + timeInfo + "  " + t.Name;
            rowLbl.Dock = DockStyle.Fill;
            rowLbl.TextAlign = ContentAlignment.MiddleLeft;
            rowLbl.AutoSize = true;
            rowLbl.Padding = new Padding(0, (int)(8 * scale), 0, (int)(8 * scale));
            rowLbl.Font = UITheme.GetFont(10.5f);

            row.Controls.Add(bE, 0, 0); 
            row.Controls.Add(bD, 1, 0); 
            row.Controls.Add(bN, 2, 0);
            row.Controls.Add(rowLbl, 3, 0);
            
            inner.Controls.Add(row);
        }
        gb.Controls.Add(inner); 
        flow.Controls.Add(gb);
    }
}
