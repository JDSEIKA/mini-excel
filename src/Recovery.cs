using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace CalcRapido
{
    internal sealed class RecoveryInfo
    {
        public string SnapshotPath;
        public string MetadataPath;
        public string OriginalPath;
        public string DisplayName;
        public DateTime SavedAt;
    }

    internal static class RecoveryManager
    {
        private static string _directory;
        public static string LastError { get; internal set; }

        private static string DirectoryPath
        {
            get
            {
                if (_directory != null) return _directory;
                string[] candidates = new string[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "MiniExcel", "Recovery"),
                    Path.Combine(Application.StartupPath, ".mini-excel-recovery")
                };
                foreach (string candidate in candidates)
                {
                    try
                    {
                        Directory.CreateDirectory(candidate);
                        string probe = Path.Combine(candidate, ".write-test");
                        File.WriteAllText(probe, "ok");
                        File.Delete(probe);
                        _directory = candidate;
                        return _directory;
                    }
                    catch { }
                }
                _directory = candidates[candidates.Length - 1];
                return _directory;
            }
        }

        public static string SnapshotPath
        {
            get { return Path.Combine(DirectoryPath, "session.recovery.xlsx"); }
        }

        public static string MetadataPath
        {
            get { return Path.Combine(DirectoryPath, "session.recovery.txt"); }
        }

        public static bool TryRead(out RecoveryInfo info)
        {
            info = null;
            try
            {
                string meta = MetadataPath;
                string snapshot = SnapshotPath;
                if (!File.Exists(meta) || !File.Exists(snapshot)) return false;
                byte[] rawMeta = File.ReadAllBytes(meta);
                string metaText;
                try
                {
                    metaText = Encoding.UTF8.GetString(ProtectedData.Unprotect(rawMeta, null,
                        DataProtectionScope.CurrentUser));
                }
                catch
                {
                    // Compatibilidad con recuperaciones antiguas en texto plano.
                    metaText = Encoding.UTF8.GetString(rawMeta);
                }
                Dictionary<string, string> values = new Dictionary<string, string>();
                foreach (string line in metaText.Split(new string[] { "\r\n", "\n" },
                    StringSplitOptions.RemoveEmptyEntries))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    values[line.Substring(0, eq)] = line.Substring(eq + 1);
                }
                string encodedPath, encodedName, encodedTime;
                values.TryGetValue("path", out encodedPath);
                values.TryGetValue("name", out encodedName);
                values.TryGetValue("time", out encodedTime);
                info = new RecoveryInfo
                {
                    SnapshotPath = snapshot,
                    MetadataPath = meta,
                    OriginalPath = Decode(encodedPath),
                    DisplayName = Decode(encodedName),
                    SavedAt = ParseTime(encodedTime)
                };
                if (string.IsNullOrEmpty(info.DisplayName))
                    info.DisplayName = "Libro recuperado";
                return true;
            }
            catch
            {
                info = null;
                return false;
            }
        }

        public static void Save(Workbook workbook, string originalPath, string displayName)
        {
            LastError = null;
            if (workbook == null) return;
            string directory = DirectoryPath;
            if (!Directory.Exists(directory)) return;
            string plain = Path.Combine(directory, "~recovery_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                XlsxIO.Save(workbook, plain);
                byte[] snapshot = File.ReadAllBytes(plain);
                byte[] sealedSnapshot = ProtectedData.Protect(snapshot, null,
                    DataProtectionScope.CurrentUser);
                AtomicIO.Save(SnapshotPath, delegate(string temp)
                {
                    File.WriteAllBytes(temp, sealedSnapshot);
                });
            }
            finally
            {
                try { if (File.Exists(plain)) File.Delete(plain); } catch { }
            }
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("version=2");
            sb.AppendLine("path=" + Encode(originalPath));
            sb.AppendLine("name=" + Encode(displayName));
            sb.AppendLine("time=" + DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture));
            byte[] meta = ProtectedData.Protect(Encoding.UTF8.GetBytes(sb.ToString()), null,
                DataProtectionScope.CurrentUser);
            AtomicIO.Save(MetadataPath, delegate(string temp)
            {
                File.WriteAllBytes(temp, meta);
            });
        }

        public static Workbook LoadSnapshot(RecoveryInfo info)
        {
            if (info == null || !File.Exists(info.SnapshotPath))
                throw new FileNotFoundException("No existe el snapshot de recuperación.");
            byte[] raw = File.ReadAllBytes(info.SnapshotPath);
            byte[] plain;
            try
            {
                plain = ProtectedData.Unprotect(raw, null, DataProtectionScope.CurrentUser);
            }
            catch
            {
                // Compatibilidad con snapshots antiguos sin cifrar.
                plain = raw;
            }
            string temp = Path.Combine(DirectoryPath,
                "~restore_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                File.WriteAllBytes(temp, plain);
                Workbook wb = XlsxIO.Load(temp);
                wb.Recalculate();
                return wb;
            }
            finally
            {
                try { if (File.Exists(temp)) File.Delete(temp); } catch { }
            }
        }

        public static void Clear()
        {
            try { if (File.Exists(SnapshotPath)) File.Delete(SnapshotPath); } catch { }
            try { if (File.Exists(MetadataPath)) File.Delete(MetadataPath); } catch { }
        }

        private static string Encode(string value)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value ?? ""));
        }

        private static string Decode(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            try { return Encoding.UTF8.GetString(Convert.FromBase64String(value)); }
            catch { return ""; }
        }

        private static DateTime ParseTime(string value)
        {
            DateTime parsed;
            if (DateTime.TryParse(value, CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind, out parsed)) return parsed.ToLocalTime();
            return DateTime.Now;
        }
    }
}
