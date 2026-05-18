// ============================================================
// FILE: MiniProgram01/MainForm.cs
// ============================================================

using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Text.Json; // 用於簡單的序列化儲存快捷鍵設定

public class MainForm : Form {
    public NotifyIcon trayIcon;
    public ContextMenuStrip trayMenu;
    private TabControl tabControl;
    private string appName = "MiniProgram01";
    private HashSet<int> alertTabs = new HashSet<int>();
    private System.Windows.Forms.Timer flashTimer;
    private bool flashState = false;

    // --- 快捷鍵與 Windows 訊息相關 API ---
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    [DllImport("user32.dll", SetLastError = true)]
    static extern uint RegisterWindowMessage(string lpString);

    private const int HOTKEY_ID_MAIN = 9000; // 固定為 Ctrl+1 呼叫主程式
    private const uint MOD_CONTROL = 0x0002;
    private const uint VK_1 = 0x31; 
    private const int WM_HOTKEY = 0x0312;

    // 用來接收 Program.cs 發送的喚醒廣播
    private readonly uint WM_SHOWME = RegisterWindowMessage("WM_SHOW_MINIPROGRAM01");

    // 【修改點】儲存 10 組自訂快捷鍵的資料結構
    public class CustomHotkeyInfo {
        public string Name { get; set; } = "";
        public string Path { get; set; } = "";
        public uint Modifiers { get; set; } = 0; // MOD_CONTROL, MOD_ALT, MOD_SHIFT
        public uint Vk { get; set; } = 0;        // 虛擬鍵碼
        public string ModifierStr { get; set; } = "Ctrl";
        public string KeyStr { get; set; } = "None";
    }

    public Dictionary<int, CustomHotkeyInfo> customHotkeys = new Dictionary<int, CustomHotkeyInfo>();
    private const int HOTKEY_BASE_ID = 9001; // 從 9001 ~ 9010 供自訂快捷鍵使用

    public App_FileWatcher fileWatcherApp;
    public App_TodoList todoApp;
    public App_TodoList planApp; 
    public App_TodoList scheduleApp; 
    public App_RecurringTasks recurringApp;
    public App_Shortcuts shortcutsApp;
    public App_Screenshot screenshotApp;

    public MainForm() {
        LoadHotkeySettings();

        this.Text = "整合通知中心";
        float scale = this.DeviceDpi / 96f;
        this.Width = (int)(520 * scale); 
        this.Height = (int)(600 * scale); 
        this.FormBorderStyle = FormBorderStyle.Sizable; 
        this.MinimumSize = new Size((int)(450 * scale), (int)(500 * scale)); 
        this.StartPosition = FormStartPosition.Manual;
        this.TopMost = true;
        this.ShowInTaskbar = false; 
        this.BackColor = UITheme.BgGray;

        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        this.Left = area.Right - this.Width - 10;
        this.Top = area.Bottom - this.Height - 10;

        BuildTrayMenu();
        
        trayIcon = new NotifyIcon();
        trayIcon.Icon = SystemIcons.Application; 
        trayIcon.ContextMenuStrip = trayMenu;
        trayIcon.Visible = true;
        trayIcon.Text = "整合通知中心";
        trayIcon.DoubleClick += (s, e) => ShowAppWindow();

        // 註冊呼叫主程式的固定快捷鍵
        RegisterHotKey(this.Handle, HOTKEY_ID_MAIN, MOD_CONTROL, VK_1);
        
        // 註冊自訂的快捷鍵
        RegisterAllCustomHotkeys();

        tabControl = new TabControl();
        tabControl.Dock = DockStyle.Fill;
        tabControl.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        tabControl.ItemSize = new Size((int)(68 * scale), (int)(38 * scale));
        tabControl.Padding = new Point(0, 0);
        tabControl.SizeMode = TabSizeMode.Fixed; 
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.DrawItem += TabControl_DrawItem;
        tabControl.SelectedIndexChanged += (s, e) => { 
            alertTabs.Remove(tabControl.SelectedIndex);
            tabControl.Invalidate(); 
        };
        this.Controls.Add(tabControl);

        fileWatcherApp = new App_FileWatcher(this, trayMenu);
        todoApp = new App_TodoList(this, "todo", "待辦清單");
        planApp = new App_TodoList(this, "plan", "待規清單");
        scheduleApp = new App_TodoList(this, "schedule", "行程清單"); 
        
        todoApp.TargetLists.Add("待規", planApp);
        todoApp.TargetLists.Add("行程", scheduleApp);
        planApp.TargetLists.Add("待辦", todoApp);
        planApp.TargetLists.Add("行程", scheduleApp);
        scheduleApp.TargetLists.Add("待辦", todoApp);
        scheduleApp.TargetLists.Add("待規", planApp);

        recurringApp = new App_RecurringTasks(this, todoApp); 
        shortcutsApp = new App_Shortcuts(this);
        screenshotApp = new App_Screenshot(this);

        fileWatcherApp.Dock = DockStyle.Fill;
        todoApp.Dock = DockStyle.Fill;
        planApp.Dock = DockStyle.Fill;
        scheduleApp.Dock = DockStyle.Fill; 
        recurringApp.Dock = DockStyle.Fill;
        shortcutsApp.Dock = DockStyle.Fill;
        screenshotApp.Dock = DockStyle.Fill;

        tabControl.TabPages.Add(new TabPage("監控") { BackColor = UITheme.BgGray }); 
        tabControl.TabPages.Add(new TabPage("待辦") { BackColor = UITheme.BgGray }); 
        tabControl.TabPages.Add(new TabPage("待規") { BackColor = UITheme.BgGray }); 
        tabControl.TabPages.Add(new TabPage("行程") { BackColor = UITheme.BgGray }); 
        tabControl.TabPages.Add(new TabPage("週期") { BackColor = UITheme.BgGray }); 
        tabControl.TabPages.Add(new TabPage("捷徑") { BackColor = UITheme.BgGray }); 
        tabControl.TabPages.Add(new TabPage("截圖") { BackColor = UITheme.BgGray }); 

        tabControl.TabPages[0].Controls.Add(fileWatcherApp);
        tabControl.TabPages[1].Controls.Add(todoApp);
        tabControl.TabPages[2].Controls.Add(planApp); 
        tabControl.TabPages[3].Controls.Add(scheduleApp); 
        tabControl.TabPages[4].Controls.Add(recurringApp);
        tabControl.TabPages[5].Controls.Add(shortcutsApp);
        tabControl.TabPages[6].Controls.Add(screenshotApp);

        flashTimer = new System.Windows.Forms.Timer() { Interval = 500 };
        flashTimer.Tick += (s, e) => {
            if (alertTabs.Count == 0) { 
                flashTimer.Stop();
                flashState = false; 
            } else { 
                flashState = !flashState;
                tabControl.Invalidate(); 
            }
        };
    }

    private void BuildTrayMenu() {
        if (trayMenu == null) trayMenu = new ContextMenuStrip();
        trayMenu.Items.Clear();

        ToolStripMenuItem startupItem = new ToolStripMenuItem("開機自動啟動", null, ToggleStartup);
        startupItem.Checked = IsRunOnStartup();
        trayMenu.Items.Add(startupItem);
        trayMenu.Items.Add(new ToolStripSeparator());
        
        trayMenu.Items.Add("顯示主視窗 (Ctrl+1)", null, (s, e) => ShowAppWindow());

        // 動態將設定好的快捷鍵名稱加入選單，供點擊執行
        for (int i = 0; i < 10; i++) {
            if (customHotkeys.TryGetValue(i, out CustomHotkeyInfo info)) {
                if (!string.IsNullOrWhiteSpace(info.Name) && !string.IsNullOrWhiteSpace(info.Path)) {
                    string displayText = $"{info.Name} ({info.ModifierStr}+{info.KeyStr})";
                    int captureId = i;
                    trayMenu.Items.Add(displayText, null, (s, e) => LaunchExternalApp(captureId));
                }
            }
        }
        
        trayMenu.Items.Add(new ToolStripSeparator());
        trayMenu.Items.Add("快捷鍵程式設定", null, (s, e) => OpenPathSettingsWindow());
        trayMenu.Items.Add(new ToolStripSeparator());
        trayMenu.Items.Add("完全退出", null, (s, e) => { trayIcon.Visible = false; Environment.Exit(0); });
    }

    private void LoadHotkeySettings() {
        customHotkeys.Clear();
        for (int i = 0; i < 10; i++) {
            string json = DbHelper.GetSetting($"CustomHotkey_{i}", "");
            if (!string.IsNullOrEmpty(json)) {
                try {
                    customHotkeys[i] = JsonSerializer.Deserialize<CustomHotkeyInfo>(json);
                } catch {
                    customHotkeys[i] = new CustomHotkeyInfo();
                }
            } else {
                customHotkeys[i] = new CustomHotkeyInfo();
            }
        }
    }

    public void SaveAndApplyHotkeySettings(Dictionary<int, CustomHotkeyInfo> newSettings) {
        // 先註銷舊的
        for (int i = 0; i < 10; i++) {
            UnregisterHotKey(this.Handle, HOTKEY_BASE_ID + i);
        }

        customHotkeys = newSettings;
        for (int i = 0; i < 10; i++) {
            string json = JsonSerializer.Serialize(customHotkeys[i]);
            DbHelper.SetSetting($"CustomHotkey_{i}", json);
        }

        // 重新註冊
        RegisterAllCustomHotkeys();
        BuildTrayMenu(); // 更新右鍵選單
    }

    private void RegisterAllCustomHotkeys() {
        for (int i = 0; i < 10; i++) {
            if (customHotkeys.TryGetValue(i, out CustomHotkeyInfo info)) {
                if (info.Modifiers > 0 && info.Vk > 0 && !string.IsNullOrWhiteSpace(info.Path)) {
                    bool success = RegisterHotKey(this.Handle, HOTKEY_BASE_ID + i, info.Modifiers, info.Vk);
                    if (!success) {
                        // 註冊失敗通常是因為組合鍵被其他程式佔用
                        Console.WriteLine($"Hotkey registration failed for index {i}");
                    }
                }
            }
        }
    }

    private void TabControl_DrawItem(object sender, DrawItemEventArgs e) {
        TabPage page = tabControl.TabPages[e.Index];
        bool isSelected = e.Index == tabControl.SelectedIndex;
        bool isAlert = alertTabs.Contains(e.Index);

        using (SolidBrush bgBrush = new SolidBrush(UITheme.BgGray)) { 
            e.Graphics.FillRectangle(bgBrush, e.Bounds); 
        }

        Color textColor = isSelected ? UITheme.AppleBlue : UITheme.TextSub;
        if (isAlert && flashState && !isSelected) {
            textColor = UITheme.AppleRed;
        }

        using (SolidBrush textBrush = new SolidBrush(textColor)) {
            StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            e.Graphics.DrawString(page.Text, e.Font, textBrush, e.Bounds, sf);
        }

        if (isSelected) {
            float scale = this.DeviceDpi / 96f;
            using (SolidBrush lineBrush = new SolidBrush(UITheme.AppleBlue)) {
                e.Graphics.FillRectangle(lineBrush, e.Bounds.Left + 8, e.Bounds.Bottom - (int)(4 * scale), e.Bounds.Width - 16, (int)(4 * scale));
            }
        }
    }

    public void AlertTab(int tabIndex) {
        if (tabControl.SelectedIndex != tabIndex) {
            alertTabs.Add(tabIndex);
            if (!flashTimer.Enabled) flashTimer.Start();
        }
    }

    public void ClearTabAlert(int tabIndex) {
        alertTabs.Remove(tabIndex);
        tabControl.Invalidate();
    }

    public void ShowAppWindow(int tabIndex = -1) {
        if (tabIndex >= 0 && tabIndex < tabControl.TabPages.Count) {
            tabControl.SelectedIndex = tabIndex;
        }
        this.Show(); 
        this.WindowState = FormWindowState.Normal; 
        this.Activate(); 
    }

    private void HideAppWindow() { this.Hide(); }

    private void ToggleStartup(object sender, EventArgs e) {
        ToolStripMenuItem item = sender as ToolStripMenuItem;
        if (item == null) return;
        bool newState = !item.Checked;
        SetRunOnStartup(newState);
        item.Checked = newState;
    }

    private bool IsRunOnStartup() {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", false)) {
            return key?.GetValue(appName) != null;
        }
    }

    private void SetRunOnStartup(bool enable) {
        using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true)) {
            if (enable) {
                string launcherPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "MyTool.exe");
                launcherPath = Path.GetFullPath(launcherPath); 
                if (File.Exists(launcherPath)) {
                    key.SetValue(appName, launcherPath);
                } else {
                    key.SetValue(appName, Application.ExecutablePath);
                }
            }
            else key.DeleteValue(appName, false);
        }
    }

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_SHOWME) {
            ShowAppWindow();
        }
        if (m.Msg == WM_HOTKEY) {
            int id = m.WParam.ToInt32();
            if (id == HOTKEY_ID_MAIN) { 
                ShowAppWindow(); 
            } else if (id >= HOTKEY_BASE_ID && id < HOTKEY_BASE_ID + 10) { 
                LaunchExternalApp(id - HOTKEY_BASE_ID); 
            }
        }
        base.WndProc(ref m);
    }

    private void LaunchExternalApp(int index) {
        if (customHotkeys.TryGetValue(index, out CustomHotkeyInfo info) && !string.IsNullOrWhiteSpace(info.Path)) {
            try {
                if (File.Exists(info.Path) || Directory.Exists(info.Path) || info.Path.StartsWith("http")) {
                    ProcessStartInfo psi = new ProcessStartInfo() { FileName = info.Path, UseShellExecute = true };
                    // 只有是執行檔且非網址時，才嘗試設定工作目錄
                    if (File.Exists(info.Path) && !info.Path.StartsWith("http")) {
                        psi.WorkingDirectory = Path.GetDirectoryName(info.Path);
                    }
                    Process.Start(psi);
                } else {
                    MessageBox.Show($"找不到目標：\n{info.Path}\n\n請確認路徑或網址是否正確。", "執行失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            } catch (Exception ex) {
                MessageBox.Show($"啟動時發生錯誤：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }

    private void OpenPathSettingsWindow() {
        using (var form = new HotkeyPathSettingsForm(customHotkeys, this)) {
            form.ShowDialog();
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e) {
        if (e.CloseReason == CloseReason.UserClosing) { 
            e.Cancel = true; 
            HideAppWindow(); 
        } 
        base.OnFormClosing(e);
    }

    protected override void OnResize(EventArgs e) {
        if (this.WindowState == FormWindowState.Minimized) { HideAppWindow(); } 
        base.OnResize(e);
    }

    protected override void Dispose(bool disposing) {
        if (disposing) {
            UnregisterHotKey(this.Handle, HOTKEY_ID_MAIN);
            for (int i = 0; i < 10; i++) {
                UnregisterHotKey(this.Handle, HOTKEY_BASE_ID + i);
            }
            if (trayIcon != null) { trayIcon.Visible = false; trayIcon.Dispose(); }
            if (flashTimer != null) { flashTimer.Dispose(); }
        }
        base.Dispose(disposing);
    }
}

public class HotkeyPathSettingsForm : Form {
    private Dictionary<int, MainForm.CustomHotkeyInfo> currentSettings;
    private MainForm parentForm;
    private float scale;

    // 存放 UI 控制項，以便取值
    private class RowControls {
        public TextBox txtName;
        public ComboBox cmbModifier;
        public ComboBox cmbKey;
        public TextBox txtPath;
    }
    private Dictionary<int, RowControls> uiRows = new Dictionary<int, RowControls>();

    public HotkeyPathSettingsForm(Dictionary<int, MainForm.CustomHotkeyInfo> settings, MainForm parent) {
        // 深拷貝，避免取消時影響原設定
        this.currentSettings = new Dictionary<int, MainForm.CustomHotkeyInfo>();
        foreach (var kvp in settings) {
            this.currentSettings[kvp.Key] = new MainForm.CustomHotkeyInfo {
                Name = kvp.Value.Name,
                Path = kvp.Value.Path,
                Modifiers = kvp.Value.Modifiers,
                Vk = kvp.Value.Vk,
                ModifierStr = kvp.Value.ModifierStr,
                KeyStr = kvp.Value.KeyStr
            };
        }
        this.parentForm = parent;
        this.scale = this.DeviceDpi / 96f;

        this.Text = "自訂快捷鍵設定 (最多10組)";
        this.Width = (int)(750 * scale);
        this.Height = (int)(650 * scale); 
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.BackColor = UITheme.BgGray;

        InitializeUI();
    }

    private void InitializeUI() {
        FlowLayoutPanel panel = new FlowLayoutPanel() {
            Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown,
            Padding = new Padding((int)(15 * scale)), AutoScroll = true, WrapContents = false
        };

        // 標題列
        TableLayoutPanel header = new TableLayoutPanel {
            Width = (int)(680 * scale), Height = (int)(30 * scale), ColumnCount = 5,
            Margin = new Padding(0, 0, 0, (int)(10 * scale))
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(30 * scale)));  // ID
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(150 * scale))); // 自訂名稱
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(80 * scale)));  // 修飾鍵
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(100 * scale))); // 組合鍵
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));                // 程式路徑
        
        Label h1 = new Label { Text = "#", Font = UITheme.GetFont(10f, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter };
        Label h2 = new Label { Text = "自訂顯示名稱", Font = UITheme.GetFont(10f, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft };
        Label h3 = new Label { Text = "修飾鍵", Font = UITheme.GetFont(10f, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft };
        Label h4 = new Label { Text = "按鍵", Font = UITheme.GetFont(10f, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft };
        Label h5 = new Label { Text = "目標路徑 (執行檔 / 網址)", Font = UITheme.GetFont(10f, FontStyle.Bold), TextAlign = ContentAlignment.MiddleLeft };
        
        header.Controls.Add(h1, 0, 0); header.Controls.Add(h2, 1, 0);
        header.Controls.Add(h3, 2, 0); header.Controls.Add(h4, 3, 0); header.Controls.Add(h5, 4, 0);
        panel.Controls.Add(header);

        // 產生 10 組設定列
        for (int i = 0; i < 10; i++) {
            AddRowUI(panel, i);
        }

        Panel bottomPanel = new Panel() { Dock = DockStyle.Bottom, Height = (int)(70 * scale) };
        
        Button btnSave = new Button() { 
            Text = "儲存並套用", Width = (int)(140 * scale), Height = (int)(40 * scale), 
            Left = (int)(250 * scale), Top = (int)(10 * scale),
            BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite,
            FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(10f, FontStyle.Bold), Cursor = Cursors.Hand
        };
        btnSave.FlatAppearance.BorderSize = 0;

        Button btnCancel = new Button() { 
            Text = "取消", Width = (int)(100 * scale), Height = (int)(40 * scale), 
            Left = (int)(400 * scale), Top = (int)(10 * scale),
            BackColor = UITheme.CardWhite, ForeColor = UITheme.TextMain,
            FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(10f), Cursor = Cursors.Hand
        };
        btnCancel.FlatAppearance.BorderColor = Color.LightGray;

        btnSave.Click += BtnSave_Click;
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        bottomPanel.Controls.Add(btnSave); bottomPanel.Controls.Add(btnCancel);
        this.Controls.Add(panel); this.Controls.Add(bottomPanel);
    }

    private void AddRowUI(FlowLayoutPanel parent, int index) {
        TableLayoutPanel row = new TableLayoutPanel {
            Width = (int)(680 * scale), Height = (int)(38 * scale), ColumnCount = 6,
            Margin = new Padding(0, 0, 0, (int)(8 * scale))
        };
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(30 * scale)));  // ID
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(145 * scale))); // 名稱
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(75 * scale)));  // 修飾
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(95 * scale)));  // 按鍵
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));                // 路徑
        row.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, (int)(55 * scale)));  // 瀏覽按鈕

        MainForm.CustomHotkeyInfo info = currentSettings[index];

        Label lblNum = new Label { Text = (index + 1).ToString(), Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, Font = UITheme.GetFont(10f) };
        
        TextBox txtName = new TextBox { Width = (int)(140 * scale), Font = UITheme.GetFont(10f), Text = info.Name, Margin = new Padding(0, (int)(6 * scale), 0, 0) };
        
        ComboBox cmbMod = new ComboBox { Width = (int)(70 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10f), Margin = new Padding(0, (int)(6 * scale), 0, 0) };
        cmbMod.Items.AddRange(new string[] { "None", "Ctrl", "Alt", "Shift" });
        cmbMod.Text = string.IsNullOrEmpty(info.ModifierStr) ? "None" : info.ModifierStr;

        ComboBox cmbKey = new ComboBox { Width = (int)(90 * scale), DropDownStyle = ComboBoxStyle.DropDownList, Font = UITheme.GetFont(10f), Margin = new Padding(0, (int)(6 * scale), 0, 0) };
        cmbKey.Items.Add("None");
        for (char c = 'A'; c <= 'Z'; c++) cmbKey.Items.Add(c.ToString());
        for (int i = 0; i <= 9; i++) cmbKey.Items.Add(i.ToString());
        for (int i = 1; i <= 12; i++) cmbKey.Items.Add($"F{i}");
        cmbKey.Text = string.IsNullOrEmpty(info.KeyStr) ? "None" : info.KeyStr;

        TextBox txtPath = new TextBox { Dock = DockStyle.Fill, Font = UITheme.GetFont(10f), Text = info.Path, Margin = new Padding(0, (int)(6 * scale), (int)(5 * scale), 0) };
        
        Button btnBrowse = new Button { 
            Text = "瀏覽", Dock = DockStyle.Fill, BackColor = UITheme.CardWhite, 
            FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(9f), Cursor = Cursors.Hand,
            Margin = new Padding(0, (int)(4 * scale), 0, (int)(4 * scale))
        };
        btnBrowse.FlatAppearance.BorderColor = Color.LightGray;
        btnBrowse.Click += (s, e) => {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "執行檔 (*.exe)|*.exe|所有檔案 (*.*)|*.*" }) {
                if (ofd.ShowDialog() == DialogResult.OK) txtPath.Text = ofd.FileName;
            }
        };

        uiRows[index] = new RowControls { txtName = txtName, cmbModifier = cmbMod, cmbKey = cmbKey, txtPath = txtPath };

        row.Controls.Add(lblNum, 0, 0);
        row.Controls.Add(txtName, 1, 0);
        row.Controls.Add(cmbMod, 2, 0);
        row.Controls.Add(cmbKey, 3, 0);
        row.Controls.Add(txtPath, 4, 0);
        row.Controls.Add(btnBrowse, 5, 0);

        parent.Controls.Add(row);
    }

    private void BtnSave_Click(object sender, EventArgs e) {
        // 驗證並轉換設定
        for (int i = 0; i < 10; i++) {
            RowControls row = uiRows[i];
            MainForm.CustomHotkeyInfo info = currentSettings[i];

            info.Name = row.txtName.Text.Trim();
            info.Path = row.txtPath.Text.Trim();
            info.ModifierStr = row.cmbModifier.Text;
            info.KeyStr = row.cmbKey.Text;

            // 轉換 Modifier
            info.Modifiers = 0;
            if (info.ModifierStr == "Ctrl") info.Modifiers = 0x0002;
            else if (info.ModifierStr == "Alt") info.Modifiers = 0x0001;
            else if (info.ModifierStr == "Shift") info.Modifiers = 0x0004;

            // 轉換虛擬鍵碼 (VK)
            info.Vk = 0;
            if (info.KeyStr != "None") {
                if (info.KeyStr.Length == 1) {
                    char c = info.KeyStr[0];
                    if (c >= 'A' && c <= 'Z') info.Vk = (uint)c;
                    else if (c >= '0' && c <= '9') info.Vk = (uint)c;
                } else if (info.KeyStr.StartsWith("F")) {
                    if (int.TryParse(info.KeyStr.Substring(1), out int fNum) && fNum >= 1 && fNum <= 12) {
                        info.Vk = (uint)(0x70 + fNum - 1); // VK_F1 是 0x70
                    }
                }
            }

            // 簡單防呆：如果有填路徑但沒填名稱，給預設名；有填名稱或路徑但沒設快捷鍵，給提示。
            if (!string.IsNullOrEmpty(info.Path) && string.IsNullOrEmpty(info.Name)) {
                info.Name = $"捷徑 {i + 1}";
            }
        }

        parentForm.SaveAndApplyHotkeySettings(currentSettings);
        MessageBox.Show("自訂快捷鍵設定已儲存並生效！\n您可隨時透過右鍵選單呼叫這些程式。", "設定成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        this.DialogResult = DialogResult.OK; 
        this.Close();
    }
}
