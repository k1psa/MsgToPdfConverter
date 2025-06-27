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
            foreach (var file in newFiles)
            {
                if (!set.Contains(file) && File.Exists(file) && Path.GetExtension(file).ToLowerInvariant() == ".msg")
                {
                    set.Add(file);
                }
            }
            return set.ToList();
        }

        public List<string> AddFilesFromDirectory(List<string> currentFiles, string directory)
        {
            if (!Directory.Exists(directory)) return currentFiles;
            var msgFiles = Directory.GetFiles(directory, "*.msg", SearchOption.AllDirectories);
            return AddFiles(currentFiles, msgFiles);
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
