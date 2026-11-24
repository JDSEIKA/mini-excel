using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace CalcRapido
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += delegate(object sender, ThreadExceptionEventArgs e)
            {
                ReportError(e.Exception);
            };
            AppDomain.CurrentDomain.UnhandledException += delegate(object sender, UnhandledExceptionEventArgs e)
            {
                ReportError(e.ExceptionObject as Exception);
            };

            string path = null;
            if (args != null)
            {
                foreach (string arg in args)
                {
                    if (!string.IsNullOrEmpty(arg) && !arg.StartsWith("-", StringComparison.Ordinal) && File.Exists(arg))
                    {
                        path = arg;
                        break;
                    }
                }
            }
            try
            {
                Application.Run(new MainForm(path));
            }
            catch (Exception ex)
            {
                ReportError(ex);
            }
        }

        private static void ReportError(Exception ex)
        {
            if (ex == null) return;
            string logPath = null;
            try
            {
                logPath = Path.Combine(Application.StartupPath, "error.log");
                FileInfo logInfo = new FileInfo(logPath);
                if (logInfo.Exists && logInfo.Length > 1024 * 1024)
                {
                    string old = logPath + ".old";
                    try { if (File.Exists(old)) File.Delete(old); File.Move(logPath, old); }
                    catch { }
                }
                File.AppendAllText(logPath,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + ex + Environment.NewLine);
            }
            catch
            {
                try
                {
                    logPath = Path.Combine(Path.GetTempPath(), "CalcRapido-error.log");
                    File.AppendAllText(logPath,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + ex + Environment.NewLine);
                }
                catch { }
            }
            string detail = ex.Message;
            if (logPath != null) detail += "\n\nDetalle en:\n" + logPath;
            MessageBox.Show("Ocurrió un error inesperado:\n\n" + detail,
                "Mini Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
