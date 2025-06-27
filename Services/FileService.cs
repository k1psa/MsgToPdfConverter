using System;
using System.IO;
using System.Threading;

namespace MsgToPdfConverter.Services
{
    public static class FileService
    {
        /// <summary>
        /// Robust file deletion with retries (for temp files, not user files)
        /// </summary>
        public static void RobustDeleteFile(string filePath, int maxRetries = 5, int delayMs = 500)
        {
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                        Thread.Sleep(100);
                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"[CLEANUP] Successfully deleted temp file: {filePath}");
                            return;
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[CLEANUP] File does not exist, skipping deletion: {filePath}");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[CLEANUP] Error deleting temp file (attempt {i + 1}/{maxRetries}): {filePath} - {ex.Message}");
                    if (i == maxRetries - 1)
                    {
                        Console.WriteLine($"[CLEANUP] Failed to delete temp file after {maxRetries} attempts: {filePath}");
                    }
                    else
                    {
                        Thread.Sleep(delayMs);
                    }
                }
            }
        }
    }
}
