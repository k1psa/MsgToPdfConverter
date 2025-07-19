#if DEBUG
using System;
using System.IO;

namespace MsgToPdfConverter
{
    public static class DebugLogger
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "MsgToPdfConverter_debuglog.txt");

        public static void Log(string message)
        {
            try
            {
                File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\n");
            }
            catch { /* Ignore logging errors in debug */ }
        }
    }
}
#endif
