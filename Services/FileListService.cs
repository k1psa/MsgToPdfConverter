using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MsgToPdfConverter.Services
{
    public class FileListService
    {
        public List<string> AddFiles(List<string> currentFiles, IEnumerable<string> newFiles)
        {

            string[] supportedExtensions = new[] { ".msg", ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".7z", ".jpg", ".jpeg", ".png", ".bmp", ".gif" };
            var extToHashSet = new Dictionary<string, HashSet<string>>();
            var result = new List<string>();

            // Helper to compute file hash with robust error handling
            string GetFileHash(string path)
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] GetFileHash: Path is null or empty.");
                    #endif
                    return null;
                }
                if (!File.Exists(path))
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] GetFileHash: File does not exist: {path}");
                    #endif
                    return null;
                }
                try
                {
                    using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        using (var sha256 = System.Security.Cryptography.SHA256.Create())
                        {
                            var hash = sha256.ComputeHash(stream);
                            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                        }
                    }
                }
                catch (Exception ex)
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] GetFileHash: Exception for {path}: {ex.Message}");
                    #endif
                    return null;
                }
            }


            // Add current files first, tracking their hashes by extension
            foreach (var file in currentFiles)
            {
                if (string.IsNullOrWhiteSpace(file))
                {
                    #if DEBUG
                    DebugLogger.Log("[FileListService] Skipping null or empty file path in currentFiles.");
                    #endif
                    continue;
                }
                if (!File.Exists(file))
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] File does not exist: {file}");
                    #endif
                    continue;
                }
                string ext = Path.GetExtension(file)?.ToLowerInvariant();
                if (string.IsNullOrEmpty(ext) || !supportedExtensions.Contains(ext))
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] Unsupported extension: {file}");
                    #endif
                    continue;
                }
                string hash = GetFileHash(file);
                if (hash == null)
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] Could not compute hash for: {file}. Skipping file.");
                    #endif
                    continue;
                }
                if (!extToHashSet.ContainsKey(ext))
                    extToHashSet[ext] = new HashSet<string>();
                if (!extToHashSet[ext].Contains(hash))
                {
                    extToHashSet[ext].Add(hash);
                    result.Add(file);
                }
                else
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] Duplicate file skipped (same hash and extension): {file}");
                    #endif
                }
            }

            // Add new files, skip if hash already exists for the same extension
            foreach (var file in newFiles)
            {
                if (string.IsNullOrWhiteSpace(file))
                {
                    #if DEBUG
                    DebugLogger.Log("[FileListService] Skipping null or empty file path in newFiles.");
                    #endif
                    continue;
                }
                if (!File.Exists(file))
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] File does not exist: {file}");
                    #endif
                    continue;
                }
                string ext = Path.GetExtension(file)?.ToLowerInvariant();
                if (string.IsNullOrEmpty(ext) || !supportedExtensions.Contains(ext))
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] Unsupported extension: {file}");
                    #endif
                    continue;
                }
                string hash = GetFileHash(file);
                if (hash == null)
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] Could not compute hash for: {file}. Skipping file.");
                    #endif
                    continue;
                }
                if (!extToHashSet.ContainsKey(ext))
                    extToHashSet[ext] = new HashSet<string>();
                if (!extToHashSet[ext].Contains(hash))
                {
                    extToHashSet[ext].Add(hash);
                    result.Add(file);
                }
                else
                {
                    #if DEBUG
                    DebugLogger.Log($"[FileListService] Duplicate file skipped (same hash and extension): {file}");
                    #endif
                }
            }
            return result;
        }

        public List<string> AddFilesFromDirectory(List<string> currentFiles, string directory)
        {
            if (!Directory.Exists(directory)) return currentFiles;
            
            string[] supportedExtensions = new[] { ".msg", ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".7z", ".jpg", ".jpeg", ".png", ".bmp", ".gif" };
            var allSupportedFiles = new List<string>();
            
            foreach (var ext in supportedExtensions)
            {
                var files = Directory.GetFiles(directory, "*" + ext, SearchOption.AllDirectories);
                allSupportedFiles.AddRange(files);
            }
            
            return AddFiles(currentFiles, allSupportedFiles);
        }

        public List<string> RemoveFiles(List<string> currentFiles, IEnumerable<string> filesToRemove)
        {
            var set = new HashSet<string>(currentFiles);
            foreach (var file in filesToRemove)
            {
                set.Remove(file);
            }
            return set.ToList();
        }

        public List<string> ClearFiles()
        {
            return new List<string>();
        }
    }
}
