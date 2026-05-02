using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;

public class MainForm : Form {
    public NotifyIcon trayIcon;
    public ContextMenuStrip trayMenu;
    private TabControl tabControl;
    private string appName = "MiniProgram01";
    private HashSet<int> alertTabs = new HashSet<int>();
    private System.Windows.Forms.Timer flashTimer;
    private bool flashState = false;

    // --- 快捷鍵相關 API 與常數 ---
    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int HOTKEY_ID = 9000;         // Ctrl+1: 喚醒視窗
    private const int HOTKEY_IMAGE_ID = 9001;   // Ctrl+2: 圖片小工具
    private const int HOTKEY_SAFETY_ID = 9002;  // Ctrl+3: Safety System
    private const int HOTKEY_GTASK_ID = 9008;   // Ctrl+9: G-Task

    private const uint MOD_CONTROL = 0x0002;
    private const uint VK_1 = 0x31; 
    private const uint VK_2 = 0x32; 
    private const uint VK_3 = 0x33;
    private const uint VK_9 = 0x39;
    private const int WM_HOTKEY = 0x0312;

    // 用於儲存捷徑路徑的設定
    private Dictionary<int, string> hotkeyPaths = new Dictionary<int, string>();

    public App_FileWatcher fileWatcherApp;
    public App_TodoList todoApp;
    public App_TodoList planApp; 
    public App_TodoList scheduleApp; // 【新增】：行程清單
    public App_RecurringTasks recurringApp;
    public App_Shortcuts shortcutsApp;
    public App_Screenshot screenshotApp;

    public MainForm() {
        LoadPathSettings();

        this.Text = "整合通知中心";
        // 動態計算 DPI 確保視窗初始大小合理
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

        // --- 設定常駐右鍵選單 ---
        trayMenu = new ContextMenuStrip();
        ToolStripMenuItem startupItem = new ToolStripMenuItem("開機自動啟動", null, ToggleStartup);
        startupItem.Checked = IsRunOnStartup();
        trayMenu.Items.Add(startupItem);
        trayMenu.Items.Add(new ToolStripSeparator());
        
        trayMenu.Items.Add("顯示主視窗 (Ctrl+1)", null, (s, e) => ShowAppWindow());
        trayMenu.Items.Add("圖片小工具 (Ctrl+2)", null, (s, e) => LaunchExternalApp(HOTKEY_IMAGE_ID));
        trayMenu.Items.Add("啟動 Safety System (Ctrl+3)", null, (s, e) => LaunchExternalApp(HOTKEY_SAFETY_ID));
        trayMenu.Items.Add("啟動 G-Task (Ctrl+9)", null, (s, e) => LaunchExternalApp(HOTKEY_GTASK_ID));
        
        trayMenu.Items.Add(new ToolStripSeparator());
        trayMenu.Items.Add("快捷鍵程式設定", null, (s, e) => OpenPathSettingsWindow());
        trayMenu.Items.Add(new ToolStripSeparator());
        trayMenu.Items.Add("完全退出", null, (s, e) => { trayIcon.Visible = false; Environment.Exit(0); });
        
        trayIcon = new NotifyIcon();
        trayIcon.Icon = SystemIcons.Application; 
        trayIcon.ContextMenuStrip = trayMenu;
        trayIcon.Visible = true;
        trayIcon.Text = "整合通知中心";
        trayIcon.DoubleClick += (s, e) => ShowAppWindow();

        // --- 註冊全域快捷鍵 ---
        RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL, VK_1);
        RegisterHotKey(this.Handle, HOTKEY_IMAGE_ID, MOD_CONTROL, VK_2);
        RegisterHotKey(this.Handle, HOTKEY_SAFETY_ID, MOD_CONTROL, VK_3); 
        RegisterHotKey(this.Handle, HOTKEY_GTASK_ID, MOD_CONTROL, VK_9); 

        // --- 初始化 TabControl (iOS 風格) ---
        tabControl = new TabControl();
        tabControl.Dock = DockStyle.Fill;
        tabControl.Font = UITheme.GetFont(10.5f, FontStyle.Bold);
        // 【修改】：DPI 動態調整 Tab 大小，將寬度從 82 縮小到 68 以塞入第七個分頁
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

        // 初始化各模組
        fileWatcherApp = new App_FileWatcher(this, trayMenu);
        todoApp = new App_TodoList(this, "todo", "待辦清單");
        planApp = new App_TodoList(this, "plan", "待規清單");
        scheduleApp = new App_TodoList(this, "schedule", "行程清單"); // 【新增】：行程清單初始化
        
        // 【修改】：設定三個清單之間的互連動態選單
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
        scheduleApp.Dock = DockStyle.Fill; // 【新增】
        recurringApp.Dock = DockStyle.Fill;
        shortcutsApp.Dock = DockStyle.Fill;
        screenshotApp.Dock = DockStyle.Fill;

        tabControl.TabPages.Add(new TabPage("監控") { BackColor = UITheme.BgGray }); // 0
        tabControl.TabPages.Add(new TabPage("待辦") { BackColor = UITheme.BgGray }); // 1
        tabControl.TabPages.Add(new TabPage("待規") { BackColor = UITheme.BgGray }); // 2
        tabControl.TabPages.Add(new TabPage("行程") { BackColor = UITheme.BgGray }); // 3 【新增】
        tabControl.TabPages.Add(new TabPage("週期") { BackColor = UITheme.BgGray }); // 4
        tabControl.TabPages.Add(new TabPage("捷徑") { BackColor = UITheme.BgGray }); // 5
        tabControl.TabPages.Add(new TabPage("截圖") { BackColor = UITheme.BgGray }); // 6

        tabControl.TabPages[0].Controls.Add(fileWatcherApp);
        tabControl.TabPages[1].Controls.Add(todoApp);
        tabControl.TabPages[2].Controls.Add(planApp); 
        tabControl.TabPages[3].Controls.Add(scheduleApp); // 【新增】
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

    // --- 資料庫設定載入 ---
    private void LoadPathSettings() {
        hotkeyPaths[HOTKEY_IMAGE_ID] = DbHelper.GetSetting($"Hotkey_{HOTKEY_IMAGE_ID}", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MiniImage", "MiniImageStudio.exe"));
        hotkeyPaths[HOTKEY_GTASK_ID] = DbHelper.GetSetting($"Hotkey_{HOTKEY_GTASK_ID}", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GTask", "GTaskNexus.exe"));
        hotkeyPaths[HOTKEY_SAFETY_ID] = DbHelper.GetSetting($"Hotkey_{HOTKEY_SAFETY_ID}", Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "SafetySystem", "Safety_System.exe"));
    }

    public void SavePathSettings(Dictionary<int, string> newPaths) {
        hotkeyPaths = newPaths;
        foreach (var kvp in hotkeyPaths) {
            DbHelper.SetSetting($"Hotkey_{kvp.Key}", kvp.Value);
        }
    }

    // --- Tab 繪製：iOS 乾淨風格 ---
    private void TabControl_DrawItem(object sender, DrawItemEventArgs e) {
        TabPage page = tabControl.TabPages[e.Index];
        bool isSelected = e.Index == tabControl.SelectedIndex;
        bool isAlert = alertTabs.Contains(e.Index);

        // 畫背景
        using (SolidBrush bgBrush = new SolidBrush(UITheme.BgGray)) { 
            e.Graphics.FillRectangle(bgBrush, e.Bounds); 
        }

        Color textColor = isSelected ? UITheme.AppleBlue : UITheme.TextSub;
        
        // 閃爍警告色
        if (isAlert && flashState && !isSelected) {
            textColor = UITheme.AppleRed;
        }

        // 繪製文字
        using (SolidBrush textBrush = new SolidBrush(textColor)) {
            StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            e.Graphics.DrawString(page.Text, e.Font, textBrush, e.Bounds, sf);
        }

        // iOS 風格的選中指示線
        if (isSelected) {
            float scale = this.DeviceDpi / 96f;
            using (SolidBrush lineBrush = new SolidBrush(UITheme.AppleBlue)) {
                // 因應寬度變窄，微調畫線的起始與寬度
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
            // 原本是 Application.ExecutablePath (這會指向 Library/MyTool.exe)
            // 修改為指向外層的啟動器路徑
            string launcherPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "MyTool.exe");
            // 取得絕對路徑並確認檔案存在
            launcherPath = Path.GetFullPath(launcherPath); 
            
            if (File.Exists(launcherPath)) {
                key.SetValue(appName, launcherPath);
            } else {
                // 如果找不到外層啟動器(例如開發環境)，就用目前的
                key.SetValue(appName, Application.ExecutablePath);
            }
        }
        else key.DeleteValue(appName, false);
    }
}

    protected override void WndProc(ref Message m) {
        if (m.Msg == WM_HOTKEY) {
            int id = m.WParam.ToInt32();
            if (id == HOTKEY_ID) { ShowAppWindow(); } 
            else { LaunchExternalApp(id); }
        }
        base.WndProc(ref m);
    }

    private void LaunchExternalApp(int hotkeyId) {
        if (hotkeyPaths.TryGetValue(hotkeyId, out string path) && !string.IsNullOrWhiteSpace(path)) {
            try {
                if (File.Exists(path)) {
                    Process.Start(new ProcessStartInfo() { FileName = path, UseShellExecute = true });
                } else {
                    MessageBox.Show($"找不到程式：\n{path}\n\n請確認路徑是否正確。", "執行失敗", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            } catch (Exception ex) {
                MessageBox.Show($"啟動程式時發生錯誤：\n{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        } else {
            MessageBox.Show("尚未設定此快捷鍵的程式路徑！\n請對常駐圖示點擊右鍵 ->「快捷鍵程式設定」中設定。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void OpenPathSettingsWindow() {
        using (var form = new HotkeyPathSettingsForm(hotkeyPaths, this)) {
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
            UnregisterHotKey(this.Handle, HOTKEY_ID);
            UnregisterHotKey(this.Handle, HOTKEY_IMAGE_ID);
            UnregisterHotKey(this.Handle, HOTKEY_SAFETY_ID);
            UnregisterHotKey(this.Handle, HOTKEY_GTASK_ID);
            if (trayIcon != null) { trayIcon.Visible = false; trayIcon.Dispose(); }
            if (flashTimer != null) { flashTimer.Dispose(); }
        }
        base.Dispose(disposing);
    }
}

// =======================================================
// 快捷鍵路徑設定視窗 (DPI 與 iOS UI 升級)
// =======================================================
public class HotkeyPathSettingsForm : Form {
    private Dictionary<int, string> currentPaths;
    private Dictionary<int, TextBox> textboxes = new Dictionary<int, TextBox>();
    private MainForm parentForm;

    public HotkeyPathSettingsForm(Dictionary<int, string> paths, MainForm parent) {
        this.currentPaths = new Dictionary<int, string>(paths);
        this.parentForm = parent;

        float scale = this.DeviceDpi / 96f;
        this.Text = "快捷鍵程式路徑設定";
        this.Width = (int)(550 * scale);
        this.Height = (int)(380 * scale);
        this.StartPosition = FormStartPosition.CenterScreen;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.BackColor = UITheme.BgGray;

        InitializeUI(scale);
    }

    private void InitializeUI(float scale) {
        FlowLayoutPanel panel = new FlowLayoutPanel() {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            Padding = new Padding((int)(20 * scale)),
            AutoScroll = true
        };

        AddPathSettingUI(panel, 9001, "Ctrl + 2 (圖片小工具):", scale);
        AddPathSettingUI(panel, 9002, "Ctrl + 3 (Safety System):", scale);
        AddPathSettingUI(panel, 9008, "Ctrl + 9 (G-Task):", scale);

        Panel bottomPanel = new Panel() { Dock = DockStyle.Bottom, Height = (int)(70 * scale) };
        
        Button btnSave = new Button() { 
            Text = "儲存設定", 
            Width = (int)(120 * scale), Height = (int)(40 * scale), 
            Left = (int)(150 * scale), Top = (int)(10 * scale),
            BackColor = UITheme.AppleBlue, ForeColor = UITheme.CardWhite,
            FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(10f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btnSave.FlatAppearance.BorderSize = 0;

        Button btnCancel = new Button() { 
            Text = "取消", 
            Width = (int)(100 * scale), Height = (int)(40 * scale), 
            Left = (int)(280 * scale), Top = (int)(10 * scale),
            BackColor = UITheme.CardWhite, ForeColor = UITheme.TextMain,
            FlatStyle = FlatStyle.Flat, Font = UITheme.GetFont(10f),
            Cursor = Cursors.Hand
        };
        btnCancel.FlatAppearance.BorderColor = Color.LightGray;

        btnSave.Click += BtnSave_Click;
        btnCancel.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

        bottomPanel.Controls.Add(btnSave);
        bottomPanel.Controls.Add(btnCancel);

        this.Controls.Add(panel);
        this.Controls.Add(bottomPanel);
    }

    private void AddPathSettingUI(FlowLayoutPanel parent, int id, string labelText, float scale) {
        Panel row = new Panel() { Width = (int)(480 * scale), Height = (int)(65 * scale), Margin = new Padding(0, 0, 0, (int)(10 * scale)) };
        
        Label lbl = new Label() { 
            Text = labelText, AutoSize = true, 
            Location = new Point((int)(5 * scale), (int)(5 * scale)), 
            Font = UITheme.GetFont(10f, FontStyle.Bold), ForeColor = UITheme.TextMain 
        };
        
        TextBox txt = new TextBox() { 
            Width = (int)(380 * scale), 
            Location = new Point((int)(5 * scale), (int)(30 * scale)), 
            Text = currentPaths.ContainsKey(id) ? currentPaths[id] : "",
            Font = UITheme.GetFont(10f)
        };
        
        Button btnBrowse = new Button() { 
            Text = "瀏覽...", 
            Width = (int)(80 * scale), Height = (int)(28 * scale), 
            Location = new Point((int)(395 * scale), (int)(29 * scale)),
            BackColor = UITheme.CardWhite, ForeColor = UITheme.TextMain,
            FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand,
            Font = UITheme.GetFont(9f)
        };
        btnBrowse.FlatAppearance.BorderColor = Color.LightGray;

        btnBrowse.Click += (s, e) => {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "執行檔 (*.exe)|*.exe|所有檔案 (*.*)|*.*" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    txt.Text = ofd.FileName;
                }
            }
        };

        textboxes[id] = txt;
        
        row.Controls.Add(lbl);
        row.Controls.Add(txt);
        row.Controls.Add(btnBrowse);
        parent.Controls.Add(row);
    }

    private void BtnSave_Click(object sender, EventArgs e) {
        foreach (var kvp in textboxes) {
            currentPaths[kvp.Key] = kvp.Value.Text;
        }
        parentForm.SavePathSettings(currentPaths);
        MessageBox.Show("快捷鍵路徑設定已儲存！\n下次啟動時將自動載入。", "設定成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        this.DialogResult = DialogResult.OK;
        this.Close();
    }
}
