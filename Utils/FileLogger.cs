using System;
using System.IO;

namespace MsgToPdfConverter.Utils
{
    public static class FileLogger
    {
        private static readonly object _lock = new object();
        private static string _logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MsgToPdfConverter.log");

        public static void Log(string message)
        {
            try
            {
                var logLine = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n";
                lock (_lock)
                {
                    File.AppendAllText(_logFilePath, logLine);
                }
            }
            catch { /* Swallow logging errors */ }
        }
    }
}
