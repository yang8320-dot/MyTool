using System;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;

public static class Program {
    // 宣告一個全域的 Mutex，確保單一執行個體
    private static Mutex mutex = new Mutex(true, "{A93B72C1-D264-48E3-8E5F-B4A1F6C8D9E2}");

    // --- 註冊自訂的 Windows 訊息，用於喚醒視窗 ---
    [DllImport("user32.dll", SetLastError = true)]
    static extern uint RegisterWindowMessage(string lpString);
    
    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
    
    public static readonly uint WM_SHOWME = RegisterWindowMessage("WM_SHOW_MINIPROGRAM01");
    public static readonly IntPtr HWND_BROADCAST = new IntPtr(0xffff);

    [STAThread] 
    public static void Main() { 
        if (mutex.WaitOne(TimeSpan.Zero, true)) {
            try {
                Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
                Application.EnableVisualStyles(); 
                Application.SetCompatibleTextRenderingDefault(false);

                // 初始化本地資料庫
                DbHelper.InitializeDatabase();

                Application.Run(new MainForm()); 
            } finally {
                mutex.ReleaseMutex();
            }
        } else {
            // 【優化】如果程式已在執行，廣播喚醒訊號給第一台程式，讓它彈到最上層
            PostMessage(HWND_BROADCAST, WM_SHOWME, IntPtr.Zero, IntPtr.Zero);
            return;
        }
    }
}
