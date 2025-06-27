using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace MsgToPdfConverter.Services
{
    public class OutlookImportResult
    {
        public List<string> ExtractedFiles { get; set; } = new List<string>();
        public List<string> SkippedFiles { get; set; } = new List<string>();
    }

    public class OutlookImportService
    {
        public OutlookImportResult ExtractMsgFilesFromDragDrop(IDataObject data, string outputFolder, Func<string, string> sanitizeFileName)
        {
            var result = new OutlookImportResult();
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
                return result;
            var usedFilenames = new HashSet<string>();
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
                // Try indexed format first for multiple files
                if (fileNames.Length > 1)
                {
                    string indexedFormat = $"FileContents{i}";
                    if (data.GetDataPresent(indexedFormat))
                    {
                        try
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
                                    result.ExtractedFiles.Add(destPath);
                                    success = true;
                                }
                            }
                        }
                        catch { }
                    }
                }
                // Try non-indexed format as fallback
                if (!success && data.GetDataPresent("FileContents"))
                {
                    try
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
                                result.ExtractedFiles.Add(destPath);
                                success = true;
                            }
                        }
                    }
                    catch { }
                }
                // Try alternate Outlook formats (rare)
                if (!success)
                {
                    string[] altFormats = { "RenPrivateItem", "Attachment" };
                    foreach (var alt in altFormats)
                    {
                        if (data.GetDataPresent(alt))
                        {
                            try
                            {
                                using (var fileStream = (MemoryStream)data.GetData(alt))
                                using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                                {
                                    fileStream.WriteTo(fs);
                                }
                                result.ExtractedFiles.Add(destPath);
                                success = true;
                                break;
                            }
                            catch { }
                        }
                    }
                }
                // Try FileDrop as a last resort
                if (!success && data.GetDataPresent(DataFormats.FileDrop))
                {
                    try
                    {
                        string[] dropped = (string[])data.GetData(DataFormats.FileDrop);
                        foreach (var path in dropped)
                        {
                            if (File.Exists(path) && Path.GetExtension(path).ToLowerInvariant() == ".msg")
                            {
                                File.Copy(path, destPath, true);
                                result.ExtractedFiles.Add(destPath);
                                success = true;
                                break;
                            }
                        }
                    }
                    catch { }
                }
                // Try Outlook Interop fallback
                if (!success)
                {
                    try
                    {
                        var outlookApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                        if (outlookApp != null)
                        {
                            var explorer = outlookApp.ActiveExplorer();
                            if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                            {
                                int selectionIndex = Math.Min(i + 1, explorer.Selection.Count);
                                var mailItem = explorer.Selection[selectionIndex] as Microsoft.Office.Interop.Outlook.MailItem;
                                if (mailItem != null)
                                {
                                    string safeSubject = sanitizeFileName(mailItem.Subject ?? "untitled");
                                    string interopFileName = safeSubject;
                                    if (!interopFileName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                                        interopFileName += ".msg";
                                    string interopDestPath = Path.Combine(outputFolder, interopFileName);
                                    int interopCounter = 1;
                                    while (File.Exists(interopDestPath) || usedFilenames.Contains(Path.GetFileName(interopDestPath)))
                                    {
                                        string nameWithoutExt = Path.GetFileNameWithoutExtension(interopFileName);
                                        string extension = Path.GetExtension(interopFileName);
                                        string uniqueFileName = $"{nameWithoutExt}_{interopCounter}{extension}";
                                        interopDestPath = Path.Combine(outputFolder, uniqueFileName);
                                        interopCounter++;
                                    }
                                    usedFilenames.Add(Path.GetFileName(interopDestPath));
                                    mailItem.SaveAs(interopDestPath, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                                    result.ExtractedFiles.Add(interopDestPath);
                                    success = true;
                                }
                            }
                        }
                    }
                    catch { }
                }
                if (!success)
                {
                    result.SkippedFiles.Add(fileName);
                }
            }
            return result;
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
