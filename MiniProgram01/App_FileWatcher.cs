// ============================================================
// FILE: MiniProgram01/App_FileWatcher.cs
// ============================================================

using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;
using System.Linq;
using Microsoft.Data.Sqlite;

public class App_FileWatcher : UserControl {
    private MainForm parentForm;
    private FlowLayoutPanel cardPanel;

    private List<FileSystemWatcher> watchers = new List<FileSystemWatcher>();
    private Dictionary<string, string[]> pathPairs = new Dictionary<string, string[]>();
    
    private Dictionary<string, DateTime> lastProcessedTimes = new Dictionary<string, DateTime>();
    private Dictionary<string, DateTime> lastUINotifications = new Dictionary<string, DateTime>();
    
    private readonly object lockObj = new object();
    private System.Windows.Forms.Timer retentionTimer; 
    private float scale;

    public App_FileWatcher(MainForm mainForm, ContextMenuStrip menu) {
        this.parentForm = mainForm;
        this.scale = this.DeviceDpi / 96f;
        this.BackColor = UITheme.BgGray;
        this.Padding = new Padding((int)(5 * scale));

        // --- 頂部控制列 ---
        TableLayoutPanel header = new TableLayoutPanel();
        header.Dock = DockStyle.Top;
        header.Height = (int)(50 * scale);
        header.ColumnCount = 3;
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(100 * scale))); 
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(100 * scale)));

        Label lblTitle = new Label();
        lblTitle.Text = "異動紀錄清單";
        lblTitle.Font = UITheme.GetFont(12f, FontStyle.Bold);
        lblTitle.ForeColor = UITheme.TextMain;
        lblTitle.Dock = DockStyle.Fill;
        lblTitle.TextAlign = ContentAlignment.MiddleLeft;
        lblTitle.Padding = new Padding((int)(10 * scale),0,0,0);
        
        Button btnClear = new Button();
        btnClear.Text = "一鍵清除";
        btnClear.Dock = DockStyle.Fill;
        btnClear.FlatStyle = FlatStyle.Flat;
        btnClear.Margin = new Padding((int)(2 * scale),(int)(8 * scale),(int)(2 * scale),(int)(8 * scale));
        btnClear.Cursor = Cursors.Hand;
        btnClear.BackColor = UITheme.CardWhite;
        btnClear.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnClear.FlatAppearance.BorderColor = Color.LightGray;
        btnClear.Click += (s, e) => cardPanel.Controls.Clear();
        
        Button btnGoSet = new Button();
        btnGoSet.Text = "設定參數";
        btnGoSet.Dock = DockStyle.Fill;
        btnGoSet.FlatStyle = FlatStyle.Flat;
        btnGoSet.BackColor = Color.Gainsboro;
        btnGoSet.ForeColor = UITheme.TextMain;
        btnGoSet.Margin = new Padding((int)(2 * scale),(int)(8 * scale),(int)(8 * scale),(int)(8 * scale));
        btnGoSet.Cursor = Cursors.Hand;
        btnGoSet.Font = UITheme.GetFont(10f, FontStyle.Bold);
        btnGoSet.FlatAppearance.BorderSize = 0;
        btnGoSet.Click += (s, e) => { new MonitorSettingsWindow(this).Show(); };

        header.Controls.Add(lblTitle, 0, 0);
        header.Controls.Add(btnClear, 1, 0);
        header.Controls.Add(btnGoSet, 2, 0);
        this.Controls.Add(header);

        // --- 提示卡片容器 ---
        cardPanel = new FlowLayoutPanel();
        cardPanel.Dock = DockStyle.Fill;
        cardPanel.AutoScroll = true;
        cardPanel.FlowDirection = FlowDirection.TopDown;
        cardPanel.WrapContents = false;
        cardPanel.BackColor = UITheme.BgGray;
        
        cardPanel.Resize += (s, e) => {
            int safeWidth = cardPanel.ClientSize.Width - (int)(15 * scale);
            if (safeWidth > 0) {
                cardPanel.SuspendLayout(); // 【優化】避免卡頓
                foreach (Control c in cardPanel.Controls) {
                    if (c is Panel) c.Width = safeWidth;
                }
                cardPanel.ResumeLayout(true);
            }
        };

        this.Controls.Add(cardPanel);
        cardPanel.BringToFront();

        LoadConfigFromDb();

        retentionTimer = new System.Windows.Forms.Timer() { Interval = 3600000, Enabled = true };
        retentionTimer.Tick += (s, e) => RunRetentionSweep();
        RunRetentionSweep(); 
    }

    public Dictionary<string, string[]> GetPathPairs() { return pathPairs; }

    public void AddNewTask(string src, string dsts, string method, string freq, string depth, string syncMode, string retention, string customName) {
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "INSERT INTO FileWatchers (SourcePath, DestPath, Method, Frequency, Depth, SyncMode, Retention, CustomName) VALUES (@S, @D, @M, @F, @Dp, @Sy, @R, @C)";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@S", src); cmd.Parameters.AddWithValue("@D", dsts); 
                cmd.Parameters.AddWithValue("@M", method); cmd.Parameters.AddWithValue("@F", freq);
                cmd.Parameters.AddWithValue("@Dp", depth); cmd.Parameters.AddWithValue("@Sy", syncMode);
                cmd.Parameters.AddWithValue("@R", retention); cmd.Parameters.AddWithValue("@C", customName);
                cmd.ExecuteNonQuery();
            }
        }
        LoadConfigFromDb(); 
        MessageBox.Show("監控任務已成功建立！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public void UpdateTask(string oldSrc, string newSrc, string dsts, string method, string freq, string depth, string syncMode, string retention, string customName) {
        if (!string.Equals(oldSrc, newSrc, StringComparison.OrdinalIgnoreCase) && pathPairs.ContainsKey(newSrc)) {
            MessageBox.Show("新的來源路徑已存在監控清單中！", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            using (var transaction = conn.BeginTransaction()) {
                using (var cmdDel = new SqliteCommand("DELETE FROM FileWatchers WHERE SourcePath=@OldSrc", conn, transaction)) {
                    cmdDel.Parameters.AddWithValue("@OldSrc", oldSrc);
                    cmdDel.ExecuteNonQuery();
                }
                string sql = "INSERT INTO FileWatchers (SourcePath, DestPath, Method, Frequency, Depth, SyncMode, Retention, CustomName) VALUES (@S, @D, @M, @F, @Dp, @Sy, @R, @C)";
                using (var cmdIn = new SqliteCommand(sql, conn, transaction)) {
                    cmdIn.Parameters.AddWithValue("@S", newSrc); cmdIn.Parameters.AddWithValue("@D", dsts);
                    cmdIn.Parameters.AddWithValue("@M", method); cmdIn.Parameters.AddWithValue("@F", freq);
                    cmdIn.Parameters.AddWithValue("@Dp", depth); cmdIn.Parameters.AddWithValue("@Sy", syncMode);
                    cmdIn.Parameters.AddWithValue("@R", retention); cmdIn.Parameters.AddWithValue("@C", customName);
                    cmdIn.ExecuteNonQuery();
                }
                transaction.Commit();
            }
        }
        LoadConfigFromDb(); 
        MessageBox.Show("任務修改成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public void DeleteTask(string key) {
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            using (var cmd = new SqliteCommand("DELETE FROM FileWatchers WHERE SourcePath=@Key", conn)) {
                cmd.Parameters.AddWithValue("@Key", key);
                cmd.ExecuteNonQuery();
            }
        }
        LoadConfigFromDb(); 
    }

    private void LoadConfigFromDb() {
        pathPairs.Clear();
        using (var conn = DbHelper.GetConnection()) {
            conn.Open();
            string sql = "SELECT SourcePath, DestPath, Method, Frequency, Depth, SyncMode, Retention, CustomName FROM FileWatchers";
            using (var cmd = new SqliteCommand(sql, conn)) {
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) {
                        string src = reader.GetString(0);
                        string[] data = new string[] {
                            src, reader.GetString(1), reader.GetString(2), reader.GetString(3), 
                            reader.GetString(4), reader.GetString(5), reader.GetString(6), 
                            reader.IsDBNull(7) ? "" : reader.GetString(7)
                        };
                        pathPairs[src] = data;
                    }
                }
            }
        }
        ReloadAllWatchers();
    }

    public void OpenListWindow() {
        new MonitorListWindow(this).Show();
    }

    private void ReloadAllWatchers() {
        foreach (var w in watchers) { 
            w.EnableRaisingEvents = false; 
            w.Dispose(); 
        }
        watchers.Clear();
        foreach (var val in pathPairs.Values) { 
            StartWatcherFromArray(val); 
        }
    }

    private bool IsSamePath(string p1, string p2) {
        if (string.IsNullOrEmpty(p1) || string.IsNullOrEmpty(p2)) return false;
        return string.Equals(p1.TrimEnd('\\', '/'), p2.TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);
    }

    private void StartWatcherFromArray(string[] p) {
        if (p.Length < 6) return;
        string src = p[0], depth = p[4], syncMode = p[5];
        string[] dsts = p[1].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
        
        try {
            if (Directory.Exists(src)) {
                FileSystemWatcher wSrc = new FileSystemWatcher(src) { 
                    IncludeSubdirectories = (depth != "僅本層"), 
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size,
                    EnableRaisingEvents = true 
                };
                wSrc.Changed += OnFileEvent; 
                wSrc.Created += OnFileEvent; 
                watchers.Add(wSrc);
            }

            if (syncMode == "雙向同步") {
                foreach (string dst in dsts) {
                    if (!Directory.Exists(dst)) Directory.CreateDirectory(dst);
                    FileSystemWatcher wDst = new FileSystemWatcher(dst) { 
                        IncludeSubdirectories = (depth != "僅本層"), 
                        NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName | NotifyFilters.Size,
                        EnableRaisingEvents = true 
                    };
                    wDst.Changed += OnFileEvent; 
                    wDst.Created += OnFileEvent; 
                    watchers.Add(wDst);
                }
            }
        } catch { }
    }

    private void OnFileEvent(object sender, FileSystemEventArgs e) {
        if (Directory.Exists(e.FullPath)) return;
        
        string checkName = Path.GetFileName(e.FullPath);
        
        if (checkName.StartsWith("~") || 
            checkName.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase) ||
            checkName.Equals("Thumbs.db", StringComparison.OrdinalIgnoreCase) ||
            checkName.Equals("desktop.ini", StringComparison.OrdinalIgnoreCase) ||
            checkName.Equals(".DS_Store", StringComparison.OrdinalIgnoreCase)) 
        { return; }

        string triggerRoot = ((FileSystemWatcher)sender).Path;
        string[] taskLine = null; 
        string triggerMatchedPath = "";

        foreach (var val in pathPairs.Values) {
            if (IsSamePath(triggerRoot, val[0])) { 
                taskLine = val; triggerMatchedPath = val[0]; break; 
            }
            if (val[5] == "雙向同步") {
                string[] multiDsts = val[1].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                foreach(string d in multiDsts) {
                    if (IsSamePath(triggerRoot, d)) { 
                        taskLine = val; triggerMatchedPath = d; break; 
                    }
                }
                if (taskLine != null) break;
            }
        }
        if (taskLine == null) return; 

        string cleanRoot = triggerMatchedPath.TrimEnd('\\', '/') + Path.DirectorySeparatorChar;
        if (!e.FullPath.StartsWith(cleanRoot, StringComparison.OrdinalIgnoreCase)) return;

        string relPath = e.FullPath.Substring(cleanRoot.Length);
        int targetDepth = ParseDepth(taskLine[4]);
        if (targetDepth != -1) {
            int currentDepth = relPath.Split(new char[] { '\\', '/' }).Length - 1;
            if (currentDepth > targetDepth) return;
        }

        List<string> targetFiles = new List<string>();
        string srcPath = taskLine[0];
        string[] dstPaths = taskLine[1].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

        if (!IsSamePath(triggerMatchedPath, srcPath)) {
            targetFiles.Add(Path.Combine(srcPath, relPath));
        }

        foreach (string d in dstPaths) {
            if (!IsSamePath(triggerMatchedPath, d)) {
                targetFiles.Add(Path.Combine(d, relPath));
            }
        }

        if (taskLine[5] == "雙向同步") {
            targetFiles.RemoveAll(tf => {
                if (File.Exists(tf)) {
                    if (Math.Abs((File.GetLastWriteTime(e.FullPath) - File.GetLastWriteTime(tf)).TotalSeconds) < 2) return true;
                }
                return false;
            });
        }
        if (targetFiles.Count == 0) return;

        lock (lockObj) {
            if (lastProcessedTimes.ContainsKey(e.FullPath) && (DateTime.Now - lastProcessedTimes[e.FullPath]).TotalMilliseconds < 800) return;
            lastProcessedTimes[e.FullPath] = DateTime.Now;
        }
        
        string customName = taskLine[7];
        DoBackup(e.FullPath, targetFiles, taskLine[2] == "顯示在監控", taskLine[3], relPath, customName);
    }

    private int ParseDepth(string d) {
        if (d == "無限層") return -1; if (d == "僅本層") return 0;
        if (d == "第一層") return 1; if (d == "第二層") return 2; if (d == "第三層") return 3; if (d == "第十層") return 10;
        string clean = d.Replace("層", "").Replace("第", "").Trim();
        return int.TryParse(clean, out int res) ? res : -1;
    }

    private void DoBackup(string srcFile, List<string> targetFiles, bool showUI, string freqStr, string relPath, string customName) {
        ThreadPool.QueueUserWorkItem(_ => {
            int sec; 
            if(!int.TryParse(freqStr.Replace("秒", "").Trim(), out sec)) sec = 1;
            Thread.Sleep(sec * 1000); 

            bool copySuccess = false;
            
            foreach (string tFile in targetFiles) {
                for(int i = 0; i < 3; i++) { 
                    try {
                        string dir = Path.GetDirectoryName(tFile);
                        if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                        File.Copy(srcFile, tFile, true); 
                        copySuccess = true; 
                        break;
                    } catch { Thread.Sleep(800); }
                }
            }

            if (copySuccess && showUI) {
                string[] pathParts = relPath.Split(new char[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
                bool isInsideFolder = pathParts.Length > 1;
                string displayFileName = isInsideFolder ? "📁 " + pathParts[0] + " (群組)" : Path.GetFileName(srcFile);

                bool shouldUpdateUI = false;
                lock (lockObj) {
                    if (!lastUINotifications.ContainsKey(displayFileName) || (DateTime.Now - lastUINotifications[displayFileName]).TotalSeconds > 2) {
                        lastUINotifications[displayFileName] = DateTime.Now;
                        shouldUpdateUI = true;
                    }
                }

                if (!shouldUpdateUI) return; 

                try {
                    this.Invoke(new Action(() => {
                        parentForm.AlertTab(0); 
                        string cardUniqueName = displayFileName;
                        
                        foreach (Control ctrl in cardPanel.Controls) {
                            if (ctrl.Name == cardUniqueName) {
                                cardPanel.Controls.SetChildIndex(ctrl, 0);
                                return; 
                            }
                        }
                        
                        Panel c = new Panel();
                        c.Name = cardUniqueName;
                        c.Width = (int)(340 * scale);
                        c.AutoSize = true;
                        c.BackColor = UITheme.CardWhite;
                        c.Margin = new Padding((int)(5 * scale), (int)(5 * scale), (int)(5 * scale), (int)(3 * scale));
                        
                        c.Paint += (s, ev) => {
                            UITheme.DrawRoundedBackground(ev.Graphics, new Rectangle(0, 0, c.Width - 1, c.Height - 1), (int)(8 * scale), UITheme.CardWhite);
                            using (var pen = new Pen(Color.FromArgb(230, 230, 230), 1)) {
                                ev.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                                ev.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, c.Width - 1, c.Height - 1), (int)(8 * scale)));
                            }
                        };
                        
                        TableLayoutPanel tlp = new TableLayoutPanel();
                        tlp.Dock = DockStyle.Fill;
                        tlp.ColumnCount = 2;
                        tlp.RowCount = 1;
                        tlp.AutoSize = true;
                        tlp.Padding = new Padding((int)(8 * scale));
                        tlp.BackColor = Color.Transparent;

                        tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(65 * scale))); 
                        tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); 

                        FlowLayoutPanel btnPnl = new FlowLayoutPanel();
                        btnPnl.FlowDirection = FlowDirection.TopDown;
                        btnPnl.AutoSize = true;
                        btnPnl.Margin = new Padding(0);
                        
                        Button bView = new Button();
                        bView.Text = "查看";
                        bView.Width = (int)(55 * scale);
                        bView.Height = (int)(28 * scale);
                        bView.Margin = new Padding(0, (int)(2 * scale), 0, (int)(5 * scale));
                        bView.BackColor = UITheme.AppleBlue;
                        bView.ForeColor = UITheme.CardWhite;
                        bView.FlatStyle = FlatStyle.Flat;
                        bView.Cursor = Cursors.Hand;
                        bView.Font = UITheme.GetFont(9f, FontStyle.Bold);
                        bView.FlatAppearance.BorderSize = 0;
                        bView.Click += (s, e2) => {
                            try { System.Diagnostics.Process.Start("explorer.exe", "/select,\"" + srcFile + "\""); } catch { }
                            cardPanel.Controls.Remove(c); c.Dispose(); 
                        };
                        
                        Button bClose = new Button();
                        bClose.Text = "X";
                        bClose.Width = (int)(55 * scale);
                        bClose.Height = (int)(28 * scale);
                        bClose.BackColor = UITheme.AppleRed;
                        bClose.ForeColor = UITheme.CardWhite;
                        bClose.FlatStyle = FlatStyle.Flat;
                        bClose.Cursor = Cursors.Hand;
                        bClose.Font = UITheme.GetFont(9f, FontStyle.Bold);
                        bClose.FlatAppearance.BorderSize = 0;
                        bClose.Click += (s, e2) => { cardPanel.Controls.Remove(c); c.Dispose(); };
                        
                        btnPnl.Controls.Add(bView); 
                        btnPnl.Controls.Add(bClose);

                        string displayLoc = string.IsNullOrWhiteSpace(customName) ? (isInsideFolder ? pathParts[0] : relPath) : customName;
                        Label lbl = new Label();
                        lbl.Text = displayFileName + "\n位置: " + displayLoc;
                        lbl.Dock = DockStyle.Fill;
                        lbl.AutoSize = true;
                        lbl.Padding = new Padding((int)(10 * scale), (int)(5 * scale), 0, 0);
                        lbl.TextAlign = ContentAlignment.MiddleLeft;
                        lbl.Font = UITheme.GetFont(10.5f);
                        lbl.ForeColor = UITheme.TextMain;
                        
                        tlp.Controls.Add(btnPnl, 0, 0); 
                        tlp.Controls.Add(lbl, 1, 0); 
                        c.Controls.Add(tlp);
                        
                        cardPanel.Controls.Add(c); 
                        cardPanel.Controls.SetChildIndex(c, 0);
                        
                        if (cardPanel.Controls.Count > 15) { 
                            var old = cardPanel.Controls[15]; 
                            cardPanel.Controls.RemoveAt(15); old.Dispose(); 
                        }
                    }));
                } catch { } 
            }
        });
    }

    private void RunRetentionSweep() {
        foreach (var val in pathPairs.Values) {
            if (val[5] == "單向備份" && val[6] != "永久") {
                int months = 0; 
                string mStr = val[6].Replace("個月", "").Replace("月", "").Trim();
                if (int.TryParse(mStr, out months) && months > 0) {
                    string[] dsts = val[1].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (string dst in dsts) {
                        if (Directory.Exists(dst)) {
                            try {
                                DateTime threshold = DateTime.Now.AddMonths(-months);
                                foreach (string file in Directory.GetFiles(dst, "*", SearchOption.AllDirectories)) {
                                    if (File.GetLastWriteTime(file) < threshold) { 
                                        try { File.Delete(file); } catch { } 
                                    }
                                }
                            } catch { }
                        }
                    }
                }
            }
        }
    }
}

public class MonitorSettingsWindow : Form {
    private App_FileWatcher parentWatcher;
    private TextBox txtCustomName, txtSource, txtBackup1, txtBackup2, txtBackup3;
    private ComboBox cmbMethod, cmbFreq, cmbDepth, cmbSync, cmbRetain;
    private string editingKey = null;
    private float scale;

    public MonitorSettingsWindow(App_FileWatcher watcher, string existingKey = null, string[] existingData = null) {
        this.parentWatcher = watcher;
        this.editingKey = existingKey;
        this.scale = this.DeviceDpi / 96f;

        this.Text = string.IsNullOrEmpty(existingKey) ? "新增監控設定" : "編輯監控設定";
        this.Width = (int)(430 * scale);
        this.AutoSize = true;
        this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = UITheme.BgGray;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.TopMost = true; 

        FlowLayoutPanel mainFlow = new FlowLayoutPanel();
        mainFlow.FlowDirection = FlowDirection.TopDown;
        mainFlow.WrapContents = false;
        mainFlow.Padding = new Padding((int)(20 * scale));
        mainFlow.AutoSize = true;
        
        txtCustomName = AddTextRow(mainFlow, "自訂顯示名稱：");
        txtSource = AddPathRow(mainFlow, "來源路徑：");
        
        txtBackup1 = AddPathRow(mainFlow, "備份路徑 1 (必填)：");
        txtBackup2 = AddPathRow(mainFlow, "備份路徑 2 (多重備份/選填)：");
        txtBackup3 = AddPathRow(mainFlow, "備份路徑 3 (多重備份/選填)：");
        
        cmbSync = AddComboRow(mainFlow, "同步模式：", new string[] { "單向備份", "雙向同步" });
        cmbRetain = AddComboRow(mainFlow, "自動刪除(單向)：", new string[] { "永久", "1個月", "3個月", "6個月", "12個月" });
        cmbMethod = AddComboRow(mainFlow, "通知方式：", new string[] { "顯示在監控", "隱藏背後執行" });
        cmbFreq = AddComboRow(mainFlow, "延遲執行：", new string[] { "0秒", "1秒", "3秒", "5秒", "10秒" });
        cmbDepth = AddComboRow(mainFlow, "監控深度：", new string[] { "無限層", "僅本層", "第一層", "第二層", "第三層" });

        Button btnSave = new Button();
        btnSave.Text = string.IsNullOrEmpty(existingKey) ? "新增任務" : "儲存修改";
        btnSave.Width = (int)(360 * scale);
        btnSave.Height = (int)(45 * scale);
        btnSave.BackColor = UITheme.AppleBlue;
        btnSave.ForeColor = UITheme.CardWhite;
        btnSave.FlatStyle = FlatStyle.Flat;
        btnSave.Margin = new Padding(0, (int)(20 * scale), 0, 0);
        btnSave.Font = UITheme.GetFont(11f, FontStyle.Bold);
        btnSave.Cursor = Cursors.Hand;
        btnSave.FlatAppearance.BorderSize = 0;
        
        if (existingData != null && existingData.Length >= 8) {
            txtSource.Text = existingData[0];
            
            string[] dests = existingData[1].Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            if (dests.Length > 0) txtBackup1.Text = dests[0];
            if (dests.Length > 1) txtBackup2.Text = dests[1];
            if (dests.Length > 2) txtBackup3.Text = dests[2];

            cmbMethod.Text = existingData[2];
            cmbFreq.Text = existingData[3];
            cmbDepth.Text = existingData[4];
            cmbSync.Text = existingData[5];
            cmbRetain.Text = existingData[6];
            txtCustomName.Text = existingData[7];
        }

        btnSave.Click += (s, e) => {
            if (string.IsNullOrWhiteSpace(txtSource.Text) || string.IsNullOrWhiteSpace(txtBackup1.Text)) {
                MessageBox.Show("來源與第一備份路徑不可為空！"); return;
            }
            
            var paths = new List<string> { txtBackup1.Text.Trim() };
            if (!string.IsNullOrWhiteSpace(txtBackup2.Text)) paths.Add(txtBackup2.Text.Trim());
            if (!string.IsNullOrWhiteSpace(txtBackup3.Text)) paths.Add(txtBackup3.Text.Trim());
            string combinedDsts = string.Join("|", paths);

            if (string.IsNullOrEmpty(editingKey)) {
                parentWatcher.AddNewTask(txtSource.Text, combinedDsts, cmbMethod.Text, cmbFreq.Text, cmbDepth.Text, cmbSync.Text, cmbRetain.Text, txtCustomName.Text);
            } else {
                parentWatcher.UpdateTask(editingKey, txtSource.Text, combinedDsts, cmbMethod.Text, cmbFreq.Text, cmbDepth.Text, cmbSync.Text, cmbRetain.Text, txtCustomName.Text);
            }
            this.Close();
        };

        mainFlow.Controls.Add(btnSave);

        if (string.IsNullOrEmpty(existingKey)) {
            Button btnList = new Button();
            btnList.Text = "管理現有任務";
            btnList.Width = (int)(360 * scale);
            btnList.Height = (int)(45 * scale);
            btnList.BackColor = Color.Gray;
            btnList.ForeColor = UITheme.CardWhite;
            btnList.FlatStyle = FlatStyle.Flat;
            btnList.Margin = new Padding(0, (int)(10 * scale), 0, 0);
            btnList.Font = UITheme.GetFont(11f, FontStyle.Bold);
            btnList.Cursor = Cursors.Hand;
            btnList.FlatAppearance.BorderSize = 0;

            btnList.Click += (s, e) => { parentWatcher.OpenListWindow(); this.Close(); };
            mainFlow.Controls.Add(btnList);
        }
        this.Controls.Add(mainFlow);
    }

    private TextBox AddTextRow(FlowLayoutPanel container, string label) {
        Label lbl = new Label();
        lbl.Text = label;
        lbl.AutoSize = true;
        lbl.Margin = new Padding(0, (int)(8 * scale), 0, (int)(3 * scale));
        lbl.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        container.Controls.Add(lbl);

        TextBox tb = new TextBox();
        tb.Width = (int)(360 * scale);
        tb.Font = UITheme.GetFont(10.5f);
        container.Controls.Add(tb);
        return tb;
    }

    private TextBox AddPathRow(FlowLayoutPanel container, string label) {
        Label lbl = new Label();
        lbl.Text = label;
        lbl.AutoSize = true;
        lbl.Margin = new Padding(0, (int)(8 * scale), 0, (int)(3 * scale));
        lbl.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        container.Controls.Add(lbl);

        FlowLayoutPanel row = new FlowLayoutPanel();
        row.AutoSize = true;
        row.Margin = new Padding(0, 0, 0, (int)(10 * scale));
        row.WrapContents = false;

        TextBox tb = new TextBox();
        tb.Width = (int)(250 * scale);
        tb.Font = UITheme.GetFont(10.5f);
        
        Button btnSel = new Button();
        btnSel.Text = "選";
        btnSel.Width = (int)(45 * scale);
        btnSel.Height = (int)(32 * scale);
        btnSel.FlatStyle = FlatStyle.Flat;
        btnSel.Cursor = Cursors.Hand;
        btnSel.BackColor = UITheme.CardWhite;
        btnSel.Font = UITheme.GetFont(9.5f);
        btnSel.FlatAppearance.BorderColor = Color.LightGray;
        btnSel.Click += (s, e) => { using(FolderBrowserDialog fbd = new FolderBrowserDialog()) { if(fbd.ShowDialog() == DialogResult.OK) tb.Text = fbd.SelectedPath; } };
        
        Button btnPaste = new Button();
        btnPaste.Text = "貼";
        btnPaste.Width = (int)(45 * scale);
        btnPaste.Height = (int)(32 * scale);
        btnPaste.FlatStyle = FlatStyle.Flat;
        btnPaste.Cursor = Cursors.Hand;
        btnPaste.BackColor = UITheme.CardWhite;
        btnPaste.Font = UITheme.GetFont(9.5f);
        btnPaste.FlatAppearance.BorderColor = Color.LightGray;
        btnPaste.Click += (s, e) => { string p = Clipboard.GetText().Trim(' ', '\"'); if (!string.IsNullOrEmpty(p)) tb.Text = p; };
        
        row.Controls.Add(tb); 
        row.Controls.Add(btnSel); 
        row.Controls.Add(btnPaste);
        container.Controls.Add(row);

        return tb;
    }

    private ComboBox AddComboRow(FlowLayoutPanel container, string label, string[] items) {
        Label lbl = new Label();
        lbl.Text = label;
        lbl.AutoSize = true;
        lbl.Margin = new Padding(0, (int)(8 * scale), 0, (int)(3 * scale));
        lbl.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        container.Controls.Add(lbl);

        ComboBox cb = new ComboBox();
        cb.Width = (int)(360 * scale);
        cb.Font = UITheme.GetFont(10.5f);
        cb.DropDownStyle = ComboBoxStyle.DropDownList;
        cb.Items.AddRange(items);

        if (items.Length > 0) cb.SelectedIndex = 0;
        container.Controls.Add(cb);
        return cb;
    }
}

public class MonitorListWindow : Form {
    private App_FileWatcher parentWatcher;
    private FlowLayoutPanel flow;
    private float scale;

    public MonitorListWindow(App_FileWatcher watcher) {
        this.parentWatcher = watcher;
        this.scale = this.DeviceDpi / 96f;

        this.Text = "管理現有任務";
        this.Width = (int)(550 * scale);
        this.Height = (int)(550 * scale);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.BackColor = UITheme.BgGray;
        this.TopMost = true; 

        flow = new FlowLayoutPanel();
        flow.Dock = DockStyle.Fill;
        flow.AutoScroll = true;
        flow.FlowDirection = FlowDirection.TopDown;
        flow.WrapContents = false;
        flow.Padding = new Padding((int)(15 * scale));
        
        flow.Resize += (s, e) => {
            int w = flow.ClientSize.Width - (int)(30 * scale);
            if (w > 0) {
                flow.SuspendLayout(); // 【優化】避免縮放卡頓
                foreach (Control c in flow.Controls) {
                    if (c is Panel) c.Width = w;
                }
                flow.ResumeLayout(true);
            }
        };
        
        this.Controls.Add(flow);
        RefreshList();
    }

    private void RefreshList() {
        flow.Controls.Clear();
        var pairs = parentWatcher.GetPathPairs();
        
        if (pairs.Count == 0) {
            Label noDataLbl = new Label();
            noDataLbl.Text = "目前沒有任何監控任務。";
            noDataLbl.AutoSize = true;
            noDataLbl.Margin = new Padding((int)(10 * scale));
            noDataLbl.Font = UITheme.GetFont(11f);
            noDataLbl.ForeColor = UITheme.TextMain;
            flow.Controls.Add(noDataLbl);
            return;
        }

        foreach (var kvp in pairs) {
            string key = kvp.Key;
            string[] parts = kvp.Value;
            
            string customName = parts[7];
            string displayName = string.IsNullOrWhiteSpace(customName) ? parts[0] : $"{customName}\n({parts[0]})";

            string destDisplay = parts[1].Replace("|", "\n=> ");

            Panel card = new Panel();
            card.AutoSize = true;
            card.MinimumSize = new Size(0, (int)(70 * scale));
            card.BackColor = UITheme.CardWhite;
            card.Margin = new Padding(0, 0, 0, (int)(3 * scale));
            card.Padding = new Padding((int)(12 * scale));
            card.Width = flow.ClientSize.Width > (int)(30 * scale) ? flow.ClientSize.Width - (int)(30 * scale) : (int)(400 * scale);
            
            card.Paint += (s, ev) => {
                UITheme.DrawRoundedBackground(ev.Graphics, new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(10 * scale), UITheme.CardWhite);
                using (var pen = new Pen(Color.FromArgb(230, 230, 230), 1)) {
                    ev.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    ev.Graphics.DrawPath(pen, UITheme.CreateRoundedRectanglePath(new Rectangle(0, 0, card.Width - 1, card.Height - 1), (int)(10 * scale)));
                }
            };

            TableLayoutPanel tlp = new TableLayoutPanel();
            tlp.Dock = DockStyle.Fill;
            tlp.ColumnCount = 2;
            tlp.RowCount = 1;
            tlp.AutoSize = true;
            tlp.BackColor = Color.Transparent;

            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            tlp.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(130 * scale)));

            Label lbl = new Label();
            lbl.Text = displayName + "\n=> " + destDisplay;
            lbl.Dock = DockStyle.Fill;
            lbl.AutoSize = true;
            lbl.Font = UITheme.GetFont(10.5f);
            lbl.ForeColor = UITheme.TextMain;
            lbl.Padding = new Padding(0, 0, 0, (int)(5 * scale));
            
            FlowLayoutPanel btnPanel = new FlowLayoutPanel();
            btnPanel.Dock = DockStyle.Fill;
            btnPanel.FlowDirection = FlowDirection.LeftToRight;
            btnPanel.WrapContents = false;
            btnPanel.Margin = new Padding(0);
            
            Button btnEdit = new Button();
            btnEdit.Text = "編輯";
            btnEdit.Width = (int)(55 * scale);
            btnEdit.Height = (int)(35 * scale);
            btnEdit.BackColor = UITheme.AppleBlue;
            btnEdit.ForeColor = UITheme.CardWhite;
            btnEdit.FlatStyle = FlatStyle.Flat;
            btnEdit.Cursor = Cursors.Hand;
            btnEdit.Font = UITheme.GetFont(10f, FontStyle.Bold);
            btnEdit.FlatAppearance.BorderSize = 0;

            btnEdit.Click += (s, e) => { new MonitorSettingsWindow(parentWatcher, key, parts).Show(); this.Close(); };
            
            Button btnDel = new Button();
            btnDel.Text = "刪除";
            btnDel.Width = (int)(55 * scale);
            btnDel.Height = (int)(35 * scale);
            btnDel.BackColor = UITheme.AppleRed;
            btnDel.ForeColor = UITheme.CardWhite;
            btnDel.FlatStyle = FlatStyle.Flat;
            btnDel.Cursor = Cursors.Hand;
            btnDel.Font = UITheme.GetFont(10f, FontStyle.Bold);
            btnDel.FlatAppearance.BorderSize = 0;

            btnDel.Click += (s, e) => {
                if (MessageBox.Show("確定要刪除此監控任務嗎？", "確認刪除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) {
                    parentWatcher.DeleteTask(key); RefreshList();
                }
            };
            
            btnPanel.Controls.Add(btnEdit); 
            btnPanel.Controls.Add(btnDel);

            tlp.Controls.Add(lbl, 0, 0); 
            tlp.Controls.Add(btnPanel, 1, 0);

            card.Controls.Add(tlp); 
            flow.Controls.Add(card);
        }
    }
}
