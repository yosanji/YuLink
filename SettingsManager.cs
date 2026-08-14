using System;
using System.IO;
using System.Xml.Serialization;

namespace PPTWebBrowserAddIn
{
    public class AppSettings
    {
        public bool DisableWebSecurity { get; set; }
        public int ConsolePort { get; set; }
        public bool RemoteControlEnabled { get; set; }
        public string DefaultUrl { get; set; }

        public AppSettings()
        {
            DisableWebSecurity = true;
            ConsolePort = 8888;
            RemoteControlEnabled = false;
            DefaultUrl = "https://bing.com";
        }
    }

    public static class SettingsManager
    {
        private static readonly string SettingsPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PPTWebBrowserAddIn",
            "settings.xml"
        );

        private static AppSettings _current = new AppSettings();
        public static AppSettings Current 
        { 
            get { return _current; } 
            private set { _current = value; } 
        }

        static SettingsManager()
        {
            Load();
        }

        public static void Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var serializer = new XmlSerializer(typeof(AppSettings));
                    using (var reader = new StreamReader(SettingsPath))
                    {
                        Current = (AppSettings)serializer.Deserialize(reader);
                    }
                }
                else
                {
                    Current = new AppSettings();
                }
            }
            catch
            {
                Current = new AppSettings();
            }
        }

        public static void Save()
        {
            try
            {
                string directory = Path.GetDirectoryName(SettingsPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var serializer = new XmlSerializer(typeof(AppSettings));
                using (var writer = new StreamWriter(SettingsPath))
                {
                    serializer.Serialize(writer, Current);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Failed to save settings: " + ex.Message);
            }
        }
    }
}
