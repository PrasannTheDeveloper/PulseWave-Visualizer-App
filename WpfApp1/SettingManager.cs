using System;
using System.Data.SQLite;
using System.IO;

namespace WpfApp1
{
    public class SettingsManager
    {
        private readonly string connectionString;
        private readonly string dbPath;

        public SettingsManager()
        {
            // Create database in application data folder
            string appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "AudioVisualizer");
            Directory.CreateDirectory(appDataPath);

            dbPath = Path.Combine(appDataPath, "settings.db");
            connectionString = $"Data Source={dbPath};Version=3;";

            InitializeDatabase();
        }

        private void InitializeDatabase()
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();

                    string createTableQuery = @"
                        CREATE TABLE IF NOT EXISTS Settings (
                            Id INTEGER PRIMARY KEY AUTOINCREMENT,
                            SettingName TEXT UNIQUE NOT NULL,
                            SettingValue TEXT NOT NULL,
                            LastUpdated DATETIME DEFAULT CURRENT_TIMESTAMP
                        )";

                    using (var command = new SQLiteCommand(createTableQuery, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Database initialization error: {ex.Message}");
            }
        }

        public void SaveSetting(string name, object value)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();

                    string query = @"
                        INSERT OR REPLACE INTO Settings (SettingName, SettingValue, LastUpdated) 
                        VALUES (@name, @value, CURRENT_TIMESTAMP)";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@name", name);
                        command.Parameters.AddWithValue("@value", value.ToString());
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving setting {name}: {ex.Message}");
            }
        }

        public T GetSetting<T>(string name, T defaultValue)
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();

                    string query = "SELECT SettingValue FROM Settings WHERE SettingName = @name";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@name", name);

                        var result = command.ExecuteScalar();
                        if (result != null)
                        {
                            return (T)Convert.ChangeType(result, typeof(T));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading setting {name}: {ex.Message}");
            }

            return defaultValue;
        }

        public void SaveWindowSettings(double left, double top, double width, double height, System.Windows.WindowState windowState)
        {
            SaveSetting("WindowLeft", left);
            SaveSetting("WindowTop", top);
            SaveSetting("WindowWidth", width);
            SaveSetting("WindowHeight", height);
            SaveSetting("WindowState", (int)windowState);
        }

        public (double left, double top, double width, double height, System.Windows.WindowState windowState) LoadWindowSettings()
        {
            return (
                GetSetting("WindowLeft", 100.0),
                GetSetting("WindowTop", 100.0),
                GetSetting("WindowWidth", 1200.0),
                GetSetting("WindowHeight", 800.0),
                (System.Windows.WindowState)GetSetting("WindowState", (int)System.Windows.WindowState.Maximized)
            );
        }

        public void ClearAllSettings()
        {
            try
            {
                using (var connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();

                    string query = "DELETE FROM Settings";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        command.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error clearing settings: {ex.Message}");
            }
        }
    }
}