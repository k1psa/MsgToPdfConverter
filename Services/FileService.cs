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
                            #if DEBUG
                            DebugLogger.Log($"[CLEANUP] Successfully deleted temp file: {filePath}");
                            #endif
                            return;
                        }
                    }
                    else
                    {
                        #if DEBUG
                        DebugLogger.Log($"[CLEANUP] File does not exist, skipping deletion: {filePath}");
                        #endif
                        return;
                    }
                }
                catch (Exception ex)
                {
                    #if DEBUG
                    DebugLogger.Log($"[CLEANUP] Error deleting temp file (attempt {i + 1}/{maxRetries}): {filePath} - {ex.Message}");
                    if (i == maxRetries - 1)
                    {
                        DebugLogger.Log($"[CLEANUP] Failed to delete temp file after {maxRetries} attempts: {filePath}");
                    }
                    #endif
                    if (i != maxRetries - 1)
                    {
                        Thread.Sleep(delayMs);
                    }
                }
            }
        }

        /// <summary>
        /// Sanitize filename to remove illegal characters
        /// </summary>
        public static string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return "untitled.msg";

            // Remove illegal characters from filename
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                fileName = fileName.Replace(c, '_');
            }

            // Also remove some other problematic characters
            fileName = fileName.Replace(":", "_").Replace("?", "_").Replace("*", "_");

            // Ensure it's not too long (Windows has a 255 character limit for filenames)
            if (fileName.Length > 200)
            {
                string extension = Path.GetExtension(fileName);
                string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                fileName = nameWithoutExt.Substring(0, 200 - extension.Length) + extension;
            }

            return fileName;
        }

        /// <summary>
        /// Moves a file to the Windows Recycle Bin using Microsoft.VisualBasic.FileIO
        /// </summary>
        public static void MoveFileToRecycleBin(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(filePath, Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
                    #if DEBUG
                    DebugLogger.Log($"[RECYCLE] Moved to recycle bin: {filePath}");
                    #endif
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[RECYCLE] Error moving file to recycle bin: {filePath} - {ex.Message}");
                DebugLogger.Log($"[RECYCLE] Falling back to regular delete");
                #endif
                try
                {
                    File.Delete(filePath);
                }
                catch (Exception deleteEx)
                {
                    #if DEBUG
                    DebugLogger.Log($"[RECYCLE] Error deleting file: {filePath} - {deleteEx.Message}");
                    #endif
                }
            }
        }

        /// <summary>
        /// Recursively copy a directory and all its contents
        /// </summary>
        public static void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"Source directory does not exist or could not be found: {sourceDirName}");

            DirectoryInfo[] dirs = dir.GetDirectories();
            Directory.CreateDirectory(destDirName);

            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string tempPath = Path.Combine(destDirName, file.Name);
                file.CopyTo(tempPath, true);
            }

            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string tempPath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, tempPath, copySubDirs);
                }
            }
        }
    }
}
