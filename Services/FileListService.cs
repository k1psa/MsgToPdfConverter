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

            // Use a HashSet to store hashes for deduplication
            var hashSet = new HashSet<string>();
            var result = new List<string>();

            // Helper to compute file hash
            string GetFileHash(string path)
            {
                try
                {
                    using (var stream = File.OpenRead(path))
                    {
                        using (var sha256 = System.Security.Cryptography.SHA256.Create())
                        {
                            var hash = sha256.ComputeHash(stream);
                            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                        }
                    }
                }
                catch
                {
                    return null;
                }
            }

            // Add current files first, tracking their hashes
            foreach (var file in currentFiles)
            {
                if (File.Exists(file))
                {
                    string ext = Path.GetExtension(file).ToLowerInvariant();
                    if (supportedExtensions.Contains(ext))
                    {
                        string hash = GetFileHash(file);
                        if (hash != null && !hashSet.Contains(hash))
                        {
                            hashSet.Add(hash);
                            result.Add(file);
                        }
                    }
                }
            }

            // Add new files, skip if hash already exists (i.e. identical content)
            foreach (var file in newFiles)
            {
                if (File.Exists(file))
                {
                    string ext = Path.GetExtension(file).ToLowerInvariant();
                    if (supportedExtensions.Contains(ext))
                    {
                        string hash = GetFileHash(file);
                        if (hash != null && !hashSet.Contains(hash))
                        {
                            hashSet.Add(hash);
                            result.Add(file);
                        }
                    }
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
