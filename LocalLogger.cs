using System;
using System.IO;
using System.Text;
using System.Threading;

namespace Copilot_Nudge_App
{
    public static class LocalLogger
    {
        private static readonly object _lock = new object();
        private static string _logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");

        /// <summary>
        /// Initializes logger folder (optional call).
        /// </summary>
        public static void Initialize(string? directory = null)
        {
            if (!string.IsNullOrWhiteSpace(directory))
                _logDirectory = directory;

            if (!Directory.Exists(_logDirectory))
                Directory.CreateDirectory(_logDirectory);
        }

        /// <summary>
        /// Writes a message to today's log file.
        /// </summary>
        public static void Log(string message)
        {
            try
            {
                WriteLogInternal($"[INFO] {message}");
            }
            catch { /* ensure logging never breaks the app */ }
        }

        /// <summary>
        /// Writes an exception + message to today's log file.
        /// </summary>
        public static void LogException(string message, Exception ex)
        {
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine($"[ERROR] {message}");
                sb.AppendLine($"Exception: {ex.GetType().FullName}");
                sb.AppendLine($"Message: {ex.Message}");
                sb.AppendLine($"StackTrace: {ex.StackTrace}");
                WriteLogInternal(sb.ToString());
            }
            catch { /* swallow errors to avoid recursive logging failures */ }
        }

        private static void WriteLogInternal(string text)
        {
            lock (_lock)
            {
                if (!Directory.Exists(_logDirectory))
                    Directory.CreateDirectory(_logDirectory);

                string filePath = Path.Combine(_logDirectory,
                    $"log_{DateTime.Now:yyyy-MM-dd}.txt");

                string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}  {text}{Environment.NewLine}";
                File.AppendAllText(filePath, entry, Encoding.UTF8);
            }
        }
    }

}
