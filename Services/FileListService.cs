using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MsgToPdfConverter.Services
{
    public class FileListService
    {
        public List<string> AddFiles(List<string> currentFiles, IEnumerable<string> newFiles)
        {
            var set = new HashSet<string>(currentFiles);
            string[] supportedExtensions = new[] { ".msg", ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".7z", ".jpg", ".jpeg", ".png", ".bmp", ".gif" };
            
            foreach (var file in newFiles)
            {
                if (!set.Contains(file) && File.Exists(file))
                {
                    string ext = Path.GetExtension(file).ToLowerInvariant();
                    if (supportedExtensions.Contains(ext))
                    {
                        set.Add(file);
                    }
                }
            }
            return set.ToList();
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
