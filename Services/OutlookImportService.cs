using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            
            System.Diagnostics.Debug.WriteLine("[OutlookImportService] Starting extraction...");
            
            // Skip all the complex stream extraction - go directly to Outlook Interop like attachments do
            try
            {
                var outlookApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                if (outlookApp != null)
                {
                    System.Diagnostics.Debug.WriteLine("[OutlookImportService] Outlook app found");
                    var explorer = outlookApp.ActiveExplorer();
                    if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Found {explorer.Selection.Count} selected items");
                        
                        // Process each selected email
                        for (int i = 1; i <= explorer.Selection.Count; i++)
                        {
                            var mailItem = explorer.Selection[i] as Microsoft.Office.Interop.Outlook.MailItem;
                            if (mailItem != null)
                            {
                                System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Processing email: {mailItem.Subject}");
                                
                                string safeSubject = sanitizeFileName(mailItem.Subject ?? "untitled");
                                string fileName = safeSubject + ".msg";
                                string destPath = Path.Combine(outputFolder, fileName);
                                int counter = 1;
                                while (File.Exists(destPath))
                                {
                                    string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                                    string uniqueFileName = $"{nameWithoutExt}_{counter}.msg";
                                    destPath = Path.Combine(outputFolder, uniqueFileName);
                                    counter++;
                                }
                                
                                System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Saving to: {destPath}");
                                mailItem.SaveAs(destPath, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                                result.ExtractedFiles.Add(destPath);
                                System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Successfully saved: {destPath}");
                            }
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("[OutlookImportService] No explorer or selection found");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[OutlookImportService] Outlook app not found");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Exception: {ex.Message}");
                // If Interop fails, add to skipped
                result.SkippedFiles.Add("Email could not be extracted: " + ex.Message);
            }
            
            System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Extraction complete. Found {result.ExtractedFiles.Count} files, skipped {result.SkippedFiles.Count}");
            return result;
        }

        public OutlookImportResult ExtractAttachmentsFromDragDrop(IDataObject data, string outputFolder, Func<string, string> sanitizeFileName)
        {
            var result = new OutlookImportResult();
            try
            {
                // Only support FileGroupDescriptorW (Unicode)
                if (data.GetDataPresent("FileGroupDescriptorW"))
                {
                    var fileGroupStream = (MemoryStream)data.GetData("FileGroupDescriptorW");
                    fileGroupStream.Position = 0;
                    var fileNames = GetFileNamesFromFileGroupDescriptorW(fileGroupStream);
                    System.Diagnostics.Debug.WriteLine($"[OutlookImportService] FileGroupDescriptorW present. Attachment count: {fileNames.Length}");
                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        string originalName = fileNames[i];
                        string safeName = sanitizeFileName(Path.GetFileName(originalName));
                        string destPath = Path.Combine(outputFolder, safeName);
                        int counter = 1;
                        while (File.Exists(destPath))
                        {
                            string nameWithoutExt = Path.GetFileNameWithoutExtension(safeName);
                            string ext = Path.GetExtension(safeName);
                            string uniqueFileName = $"{nameWithoutExt}_{counter}{ext}";
                            destPath = Path.Combine(outputFolder, uniqueFileName);
                            counter++;
                        }
                        // The actual file data is in the FileContents stream
                        string fileContentsFormat = i == 0 ? "FileContents" : $"FileContents{i}";
                        bool hasFileContents = data.GetDataPresent(fileContentsFormat);
                        System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Attachment {i}: {originalName}, FileContentsFormat: {fileContentsFormat}, HasFileContents: {hasFileContents}");
                        if (hasFileContents)
                        {
                            using (var fileStream = (MemoryStream)data.GetData(fileContentsFormat))
                            using (var outStream = File.Create(destPath))
                            {
                                fileStream.WriteTo(outStream);
                            }
                            result.ExtractedFiles.Add(destPath);
                            System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Saved attachment: {destPath}");
                        }
                        else
                        {
                            result.SkippedFiles.Add(originalName);
                            System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Skipped (no file data): {originalName}");
                        }
                    }
                }
                else
                {
                    result.SkippedFiles.Add("No FileGroupDescriptorW present");
                    System.Diagnostics.Debug.WriteLine($"[OutlookImportService] No FileGroupDescriptorW present in drop data.");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OutlookImportService] Exception extracting attachment: {ex.Message}");
                result.SkippedFiles.Add("Attachment could not be extracted: " + ex.Message);
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
                    // Offset changed from 76 to 72 to fix filename truncation
                    stream.Position = 4 + i * 592 + 72;
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
