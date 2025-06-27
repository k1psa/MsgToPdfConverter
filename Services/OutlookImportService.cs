using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace MsgToPdfConverter.Services
{
    public class OutlookImportService
    {
        public List<string> ExtractMsgFilesFromDragEvent(IDataObject data, string outputFolder, Func<string, string> sanitizeFileName)
        {
            var tempFiles = new List<string>();
            var skippedFiles = new List<string>();
            var usedFilenames = new HashSet<string>();
            string[] fileNames = null;
            if (data.GetDataPresent("FileGroupDescriptorW"))
            {
                using (var stream = (MemoryStream)data.GetData("FileGroupDescriptorW"))
                {
                    fileNames = GetFileNamesFromFileGroupDescriptorW(stream);
                }
            }
            else if (data.GetDataPresent("FileGroupDescriptor"))
            {
                using (var stream = (MemoryStream)data.GetData("FileGroupDescriptor"))
                {
                    fileNames = GetFileNamesFromFileGroupDescriptor(stream);
                }
            }
            if (fileNames == null || fileNames.Length == 0)
                return tempFiles;
            for (int i = 0; i < fileNames.Length; i++)
            {
                string fileName = fileNames[i];
                if (!fileName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                    fileName += ".msg";
                fileName = sanitizeFileName(fileName);
                string destPath = Path.Combine(outputFolder, fileName);
                int counter = 1;
                while (File.Exists(destPath) || usedFilenames.Contains(Path.GetFileName(destPath)))
                {
                    string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                    string extension = Path.GetExtension(fileName);
                    string uniqueFileName = $"{nameWithoutExt}_{counter}{extension}";
                    destPath = Path.Combine(outputFolder, uniqueFileName);
                    counter++;
                }
                usedFilenames.Add(Path.GetFileName(destPath));
                bool success = false;
                if (fileNames.Length > 1)
                {
                    string indexedFormat = $"FileContents{i}";
                    if (data.GetDataPresent(indexedFormat))
                    {
                        using (var fileStream = (MemoryStream)data.GetData(indexedFormat))
                        {
                            if (fileStream != null && fileStream.Length > 0)
                            {
                                using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                                {
                                    fileStream.Position = 0;
                                    fileStream.WriteTo(fs);
                                }
                                tempFiles.Add(destPath);
                                success = true;
                            }
                        }
                    }
                }
                if (!success && data.GetDataPresent("FileContents"))
                {
                    using (var fileStream = (MemoryStream)data.GetData("FileContents"))
                    {
                        if (fileStream != null && fileStream.Length > 0)
                        {
                            using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                            {
                                fileStream.Position = 0;
                                fileStream.WriteTo(fs);
                            }
                            tempFiles.Add(destPath);
                            success = true;
                        }
                    }
                }
                if (!success)
                {
                    skippedFiles.Add(fileName);
                }
            }
            return tempFiles;
        }

        private string[] GetFileNamesFromFileGroupDescriptorW(Stream stream)
        {
            var fileNames = new List<string>();
            using (var reader = new BinaryReader(stream, System.Text.Encoding.Unicode))
            {
                stream.Position = 0;
                int count = reader.ReadInt32();
                for (int i = 0; i < count; i++)
                {
                    stream.Position = 4 + i * 592 + 76;
                    var nameBytes = reader.ReadBytes(520);
                    string name = System.Text.Encoding.Unicode.GetString(nameBytes).TrimEnd('\0');
                    fileNames.Add(name);
                }
            }
            return fileNames.ToArray();
        }

        private string[] GetFileNamesFromFileGroupDescriptor(Stream stream)
        {
            var fileNames = new List<string>();
            using (var reader = new BinaryReader(stream, System.Text.Encoding.Default))
            {
                stream.Position = 0;
                int count = reader.ReadInt32();
                for (int i = 0; i < count; i++)
                {
                    stream.Position = 4 + i * 592 + 76;
                    var nameBytes = reader.ReadBytes(260);
                    string name = System.Text.Encoding.Default.GetString(nameBytes).TrimEnd('\0');
                    fileNames.Add(name);
                }
            }
            return fileNames.ToArray();
        }
    }
}
