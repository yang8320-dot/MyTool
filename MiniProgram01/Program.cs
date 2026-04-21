using System;
using System.Windows.Forms;
using System.Threading;

public static class Program {
    // 宣告一個全域的 Mutex，確保單一執行個體
    private static Mutex mutex = new Mutex(true, "{A93B72C1-D264-48E3-8E5F-B4A1F6C8D9E2}");

    [STAThread] 
    public static void Main() { 
        if (mutex.WaitOne(TimeSpan.Zero, true)) {
            try {
                // 【DPI 升級】使用現代的 PerMonitorV2 確保在 125%, 150% 縮放時文字與介面完美清晰
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
            return; // 已經在執行中，退出
        }
    }
}
