// ============================================================
// FILE: MiniProgram01/DbHelper.cs
// ============================================================

using System;
using System.IO;
using Microsoft.Data.Sqlite;

public static class DbHelper {
    // 統一使用 MiniProgramData.db 作為唯一資料庫
    private static readonly string DbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MiniProgramData.db");
    private static readonly string ConnectionString = $"Data Source={DbPath};";

    public static SqliteConnection GetConnection() {
        return new SqliteConnection(ConnectionString);
    }

    public static void InitializeDatabase() {
        using (var conn = GetConnection()) {
            conn.Open();

            // ==========================================
            // 1. 建立所有核心資料表 (如果不存在的話)
            // ==========================================
            
            // 待辦、待規、行程清單表 (透過 ListType 區分)
            string createTasksTable = @"
                CREATE TABLE IF NOT EXISTS Tasks (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ListType TEXT NOT NULL,      -- 'todo', 'plan', 'schedule'
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

            // 系統全域設定表 (取代文字檔)
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

            // ==========================================
            // 2. 吸收 DatabaseManager.cs 的優點：無痛升級機制
            // ==========================================
            // 確保舊版資料庫也能擁有後期擴充的新欄位，不會引發找不到欄位的錯誤
            SafeAddColumn(conn, "Tasks", "Note", "TEXT");
            SafeAddColumn(conn, "Tasks", "Color", "TEXT NOT NULL DEFAULT 'Black'");
            SafeAddColumn(conn, "RecurringTasks", "TaskType", "TEXT DEFAULT '循環'");
            SafeAddColumn(conn, "FileWatchers", "CustomName", "TEXT");
        }
    }

    // --- 安全新增欄位的共用方法 ---
    private static void SafeAddColumn(SqliteConnection conn, string tableName, string columnName, string columnDef) {
        try {
            string sql = $"ALTER TABLE {tableName} ADD COLUMN {columnName} {columnDef}";
            using (var cmd = new SqliteCommand(sql, conn)) {
                cmd.ExecuteNonQuery();
            }
        } catch {
            // 如果欄位已經存在，SQLite 會拋出 Exception。
            // 我們直接捕捉並忽略，這代表該資料庫已經是最新版本。
        }
    }

    // ==========================================
    // 3. 全域 Key-Value 設定存取方法
    // ==========================================
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
