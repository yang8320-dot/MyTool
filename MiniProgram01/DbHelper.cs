using System;
using System.IO;
using Microsoft.Data.Sqlite;

public static class DbHelper {
    private static readonly string DbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MiniProgramData.db");
    private static readonly string ConnectionString = $"Data Source={DbPath};";

    public static SqliteConnection GetConnection() {
        return new SqliteConnection(ConnectionString);
    }

    public static void InitializeDatabase() {
        using (var conn = GetConnection()) {
            conn.Open();

            // 待辦與待規清單表
            string createTasksTable = @"
                CREATE TABLE IF NOT EXISTS Tasks (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ListType TEXT NOT NULL,      -- 'todo' 或 'plan'
                    Text TEXT NOT NULL,
                    Color TEXT NOT NULL,
                    Note TEXT,
                    CreatedTime DATETIME NOT NULL,
                    OrderIndex INTEGER NOT NULL DEFAULT 0
                );";

            // 週期任務表
            string createRecurringTable = @"
                CREATE TABLE IF NOT EXISTS RecurringTasks (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL,
                    MonthStr TEXT,
                    DateStr TEXT,
                    TimeStr TEXT,
                    TaskType TEXT,
                    Note TEXT,
                    LastTriggeredDate TEXT,
                    OrderIndex INTEGER NOT NULL DEFAULT 0
                );";

            // 捷徑表
            string createShortcutsTable = @"
                CREATE TABLE IF NOT EXISTS Shortcuts (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL,
                    Path TEXT NOT NULL,
                    OrderIndex INTEGER NOT NULL DEFAULT 0
                );";

            // 監控任務表
            string createFileWatcherTable = @"
                CREATE TABLE IF NOT EXISTS FileWatchers (
                    SourcePath TEXT PRIMARY KEY,
                    DestPath TEXT,
                    Method TEXT,
                    Frequency TEXT,
                    Depth TEXT,
                    SyncMode TEXT,
                    Retention TEXT,
                    CustomName TEXT
                );";

            // 系統全域設定表 (取代 hotkey_paths_config.txt 等)
            string createSettingsTable = @"
                CREATE TABLE IF NOT EXISTS Settings (
                    SettingKey TEXT PRIMARY KEY,
                    SettingValue TEXT
                );";

            using (var cmd = new SqliteCommand(createTasksTable, conn)) cmd.ExecuteNonQuery();
            using (var cmd = new SqliteCommand(createRecurringTable, conn)) cmd.ExecuteNonQuery();
            using (var cmd = new SqliteCommand(createShortcutsTable, conn)) cmd.ExecuteNonQuery();
            using (var cmd = new SqliteCommand(createFileWatcherTable, conn)) cmd.ExecuteNonQuery();
            using (var cmd = new SqliteCommand(createSettingsTable, conn)) cmd.ExecuteNonQuery();
        }
    }

    // 簡易的 Key-Value 存取設定功能
    public static string GetSetting(string key, string defaultValue = "") {
        using (var conn = GetConnection()) {
            conn.Open();
            using (var cmd = new SqliteCommand("SELECT SettingValue FROM Settings WHERE SettingKey = @Key", conn)) {
                cmd.Parameters.AddWithValue("@Key", key);
                var result = cmd.ExecuteScalar();
                return result != null ? result.ToString() : defaultValue;
            }
        }
    }

    public static void SetSetting(string key, string value) {
        using (var conn = GetConnection()) {
            conn.Open();
            string sql = @"
                INSERT INTO Settings (SettingKey, SettingValue) 
                VALUES (@Key, @Value) 
                ON CONFLICT(SettingKey) DO UPDATE SET SettingValue = @Value;";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.Parameters.AddWithValue("@Key", key);
                cmd.Parameters.AddWithValue("@Value", value ?? "");
                cmd.ExecuteNonQuery();
            }
        }
    }
}
