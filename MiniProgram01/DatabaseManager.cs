/*
 * 檔案功能：SQLite 資料庫全域管理器 (微軟官方 Microsoft.Data.Sqlite 升級版)
 */

using System;
using Microsoft.Data.Sqlite; // 【修正】使用微軟官方命名空間
using System.IO;
using System.Windows.Forms;

public static class DatabaseManager
{
    private static string dbPath = Path.Combine(Application.StartupPath, "MainDB.sqlite");
    // 【修正】新版不需要寫 Version=3
    private static string connectionString = $"Data Source={dbPath};";

    public static void InitializeDatabase()
    {
        using (var conn = GetConnection())
        {
            conn.Open(); // 新版套件只要 Open 就會自動建立實體檔案
            
            using (var cmd = new SqliteCommand(@"CREATE TABLE IF NOT EXISTS TodoList (Id TEXT PRIMARY KEY, ListName TEXT NOT NULL, Content TEXT NOT NULL, CreatedDate TEXT NOT NULL, Color TEXT DEFAULT 'Black', Note TEXT DEFAULT '')", conn)) { cmd.ExecuteNonQuery(); }
            try { using (var cmd = new SqliteCommand("ALTER TABLE TodoList ADD COLUMN Color TEXT DEFAULT 'Black'", conn)) { cmd.ExecuteNonQuery(); } } catch { }
            try { using (var cmd = new SqliteCommand("ALTER TABLE TodoList ADD COLUMN Note TEXT DEFAULT ''", conn)) { cmd.ExecuteNonQuery(); } } catch { }
            using (var cmd = new SqliteCommand(@"CREATE TABLE IF NOT EXISTS Shortcuts (Id TEXT PRIMARY KEY, Name TEXT NOT NULL, TargetPath TEXT NOT NULL, SortOrder INTEGER NOT NULL)", conn)) { cmd.ExecuteNonQuery(); }
            using (var cmd = new SqliteCommand(@"CREATE TABLE IF NOT EXISTS FileWatcher (SourcePath TEXT PRIMARY KEY, TargetPath TEXT NOT NULL, SyncMethod TEXT, Frequency TEXT, Depth TEXT, SyncMode TEXT, Retention TEXT, CustomName TEXT)", conn)) { cmd.ExecuteNonQuery(); }
            using (var cmd = new SqliteCommand(@"CREATE TABLE IF NOT EXISTS RecurringTasks (Id TEXT PRIMARY KEY, Name TEXT NOT NULL, MonthStr TEXT, DateStr TEXT, TimeStr TEXT, LastTriggeredDate TEXT, Note TEXT, TaskType TEXT)", conn)) { cmd.ExecuteNonQuery(); }
            using (var cmd = new SqliteCommand(@"CREATE TABLE IF NOT EXISTS GlobalSettings (SettingKey TEXT PRIMARY KEY, SettingValue TEXT)", conn)) { cmd.ExecuteNonQuery(); }
        }
    }

    public static SqliteConnection GetConnection()
    {
        return new SqliteConnection(connectionString);
    }
}
