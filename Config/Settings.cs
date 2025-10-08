using System;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace UcsInspectorperu
{
    [DataContract]
    public sealed class Settings
    {
        // === Valores (con defaults) ===
        [DataMember] public double StepMm = 1.0;
        [DataMember] public double StepDeg = 1.0;
        [DataMember] public bool PasteAsDelta = true;

        [DataMember] public int Left = 200, Top = 150, Width = 700, Height = 440;

        [DataMember] public bool UseGlobalHotkey = true;
        [DataMember] public string GlobalHotkey = "Ctrl+Shift+U";

        [DataMember] public bool StartWithLastUcs = true;
        [DataMember] public int RecentMax = 3;
        [DataMember] public string[] RecentUcs = new string[0];

        // === Ruta del JSON ===
        public static string Path
        {
            get
            {
                var baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return System.IO.Path.Combine(baseDir, "UcsInspectorperu", "settings.json");
            }
        }

        // === Load / Save ===
        public static Settings Load()
        {
            try
            {
                var p = Path;
                if (!File.Exists(p)) return CreateDefault();

                using (var fs = File.OpenRead(p))
                {
                    var ser = new DataContractJsonSerializer(typeof(Settings));
                    return (Settings)ser.ReadObject(fs);
                }
            }
            catch { return CreateDefault(); }
        }

        public void Save()
        {
            try
            {
                var dir = System.IO.Path.GetDirectoryName(Path);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                using (var fs = File.Create(Path))
                {
                    var ser = new DataContractJsonSerializer(typeof(Settings));
                    ser.WriteObject(fs, this);
                }
            }
            catch { /* sin bloquear Inventor */ }
        }

        public static Settings CreateDefault()
        {
            var s = new Settings();
            s.Save();
            return s;
        }
    }
}
