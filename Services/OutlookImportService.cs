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
            
#if DEBUG
            DebugLogger.Log("[OutlookImportService] Starting extraction...");
#endif
            
            // Skip all the complex stream extraction - go directly to Outlook Interop like attachments do
            try
            {
                var outlookApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                if (outlookApp != null)
                {
#if DEBUG
                    DebugLogger.Log("[OutlookImportService] Outlook app found");
#endif
                    var explorer = outlookApp.ActiveExplorer();
                    if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                    {
#if DEBUG
                        DebugLogger.Log($"[OutlookImportService] Found {explorer.Selection.Count} selected items");
#endif
                        
                        // Process each selected email
                        for (int i = 1; i <= explorer.Selection.Count; i++)
                        {
                            var mailItem = explorer.Selection[i] as Microsoft.Office.Interop.Outlook.MailItem;
                            if (mailItem != null)
                            {
#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] Processing email: {mailItem.Subject}");
#endif
                                
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
                                
#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] Saving to: {destPath}");
#endif
                                mailItem.SaveAs(destPath, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                                result.ExtractedFiles.Add(destPath);
#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] Successfully saved: {destPath}");
#endif
                            }
                        }
                    }
                    else
                    {
#if DEBUG
                        DebugLogger.Log("[OutlookImportService] No explorer or selection found");
#endif
                    }
                }
                else
                {
#if DEBUG
                    DebugLogger.Log("[OutlookImportService] Outlook app not found");
#endif
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                DebugLogger.Log($"[OutlookImportService] Exception: {ex.Message}");
#endif
                // If Interop fails, add to skipped
                result.SkippedFiles.Add("Email could not be extracted: " + ex.Message);
            }
            
#if DEBUG
            DebugLogger.Log($"[OutlookImportService] Extraction complete. Found {result.ExtractedFiles.Count} files, skipped {result.SkippedFiles.Count}");
#endif
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
#if DEBUG
                    DebugLogger.Log($"[OutlookImportService] FileGroupDescriptorW present. Attachment count: {fileNames.Length}");
#endif
                    // Track hashes of already saved files in outputFolder
                    var existingFiles = Directory.GetFiles(outputFolder);
                    var existingHashes = new HashSet<string>();
                    foreach (var file in existingFiles)
                    {
                        try
                        {
                            using (var stream = File.OpenRead(file))
                            using (var sha256 = System.Security.Cryptography.SHA256.Create())
                            {
                                var hash = sha256.ComputeHash(stream);
                                string hashStr = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                                existingHashes.Add(hashStr);
                            }
                        }
                        catch { }
                    }
                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        string originalName = fileNames[i];
                        string safeName = sanitizeFileName(Path.GetFileName(originalName));
                        string destPath = Path.Combine(outputFolder, safeName);
                        int counter = 1;
                        // The actual file data is in the FileContents stream
                        string fileContentsFormat = i == 0 ? "FileContents" : $"FileContents{i}";
                        bool hasFileContents = data.GetDataPresent(fileContentsFormat);
#if DEBUG
                        DebugLogger.Log($"[OutlookImportService] Attachment {i}: {originalName}, FileContentsFormat: {fileContentsFormat}, HasFileContents: {hasFileContents}");
#endif
                        if (hasFileContents)
                        {
                            using (var fileStream = (MemoryStream)data.GetData(fileContentsFormat))
                            {
                                // Compute hash of incoming file
                                fileStream.Position = 0;
                                using (var sha256 = System.Security.Cryptography.SHA256.Create())
                                {
                                    var hash = sha256.ComputeHash(fileStream);
                                    string hashStr = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                                    if (existingHashes.Contains(hashStr))
                                    {
                                        result.SkippedFiles.Add(originalName);
#if DEBUG
                                        DebugLogger.Log($"[OutlookImportService] Skipped saving duplicate attachment (identical content): {originalName}");
#endif
                                        continue;
                                    }
                                    existingHashes.Add(hashStr);
                                }
                                fileStream.Position = 0;
                                // Find a unique file name
                                while (File.Exists(destPath))
                                {
                                    string nameWithoutExt = Path.GetFileNameWithoutExtension(safeName);
                                    string ext = Path.GetExtension(safeName);
                                    string uniqueFileName = $"{nameWithoutExt}_{counter}{ext}";
                                    destPath = Path.Combine(outputFolder, uniqueFileName);
                                    counter++;
                                }
                                using (var outStream = File.Create(destPath))
                                {
                                    fileStream.WriteTo(outStream);
                                }
                                result.ExtractedFiles.Add(destPath);
#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] Saved attachment: {destPath}");
#endif
                            }
                        }
                        else
                        {
                            result.SkippedFiles.Add(originalName);
#if DEBUG
                            DebugLogger.Log($"[OutlookImportService] Skipped (no file data): {originalName}");
#endif
                        }
                    }
                }
                else
                {
                    result.SkippedFiles.Add("No FileGroupDescriptorW present");
#if DEBUG
                    DebugLogger.Log($"[OutlookImportService] No FileGroupDescriptorW present in drop data.");
#endif
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                DebugLogger.Log($"[OutlookImportService] Exception extracting attachment: {ex.Message}");
#endif
                result.SkippedFiles.Add("Attachment could not be extracted: " + ex.Message);
            }
            return result;
        }

        public OutlookImportResult ExtractChildMsgFromDragDrop(IDataObject data, string outputFolder, Func<string, string> sanitizeFileName, string expectedFileName)
        {
            var result = new OutlookImportResult();
#if DEBUG
            DebugLogger.Log("[OutlookImportService] Starting child MSG extraction (hash-based)...");
#endif
            try
            {
                // 1. Get hash of dragged MSG data
                string draggedMsgHash = null;
                byte[] draggedBytes = null;
                Stream draggedStream = null;
                object fileContentsObj = null;
                object fileContents0Obj = null;
                bool fileContentsPresent = data.GetDataPresent("FileContents");
                bool fileContents0Present = data.GetDataPresent("FileContents0");

                if (fileContentsPresent)
                {
                    fileContentsObj = data.GetData("FileContents");
                    if (fileContentsObj == null)
                    {
#if DEBUG
                        DebugLogger.Log("[OutlookImportService] FileContents value is null");
#endif
                    }
                    else
                    {
#if DEBUG
                        DebugLogger.Log($"[OutlookImportService] FileContents type: {fileContentsObj.GetType().FullName}");
#endif
                        try
                        {
                            if (fileContentsObj is MemoryStream ms)
                            {

#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] FileContents MemoryStream length: {ms.Length}");
#endif
                                draggedStream = ms;
                            }
                            else if (fileContentsObj is Stream s)
                            {

#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] FileContents Stream length: {s.Length}");
#endif
                                draggedStream = s;
                            }
                            else if (fileContentsObj is byte[] arr)
                            {

#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] FileContents byte[] length: {arr.Length}");
#endif
                                draggedBytes = arr;
                            }
                            else
                            {

#if DEBUG
                                DebugLogger.Log("[OutlookImportService] FileContents is of unknown type");
#endif
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            DebugLogger.Log($"[OutlookImportService] Exception reading FileContents: {ex.Message}");
#endif
                        }
                    }
                }
                else if (fileContents0Present)
                {
                    fileContents0Obj = data.GetData("FileContents0");
                    if (fileContents0Obj == null)
                    {
#if DEBUG
                        DebugLogger.Log("[OutlookImportService] FileContents0 value is null");
#endif
                    }
                    else
                    {
#if DEBUG
                        DebugLogger.Log($"[OutlookImportService] FileContents0 type: {fileContents0Obj.GetType().FullName}");
#endif
                        try
                        {
                            if (fileContents0Obj is MemoryStream ms)
                            {

#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] FileContents0 MemoryStream length: {ms.Length}");
#endif
                                draggedStream = ms;
                            }
                            else if (fileContents0Obj is Stream s)
                            {

#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] FileContents0 Stream length: {s.Length}");
#endif
                                draggedStream = s;
                            }
                            else if (fileContents0Obj is byte[] arr)
                            {

#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] FileContents0 byte[] length: {arr.Length}");
#endif
                                draggedBytes = arr;
                            }
                            else
                            {

#if DEBUG
                                DebugLogger.Log("[OutlookImportService] FileContents0 is of unknown type");
#endif
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            DebugLogger.Log($"[OutlookImportService] Exception reading FileContents0: {ex.Message}");
#endif
                        }
                    }
                }
                if (draggedStream != null)
                {
                    draggedStream.Position = 0;
                    using (var sha256 = System.Security.Cryptography.SHA256.Create())
                    {
                        var hash = sha256.ComputeHash(draggedStream);
                        draggedMsgHash = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                    }
                }
                else if (draggedBytes != null)
                {
                    using (var sha256 = System.Security.Cryptography.SHA256.Create())
                    {
                        var hash = sha256.ComputeHash(draggedBytes);
                        draggedMsgHash = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                    }
                }
                var outlookApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                if (outlookApp != null)
                {
                    Microsoft.Office.Interop.Outlook.MailItem parentMailItem = null;
                    var inspector = outlookApp.ActiveInspector();
                    if (inspector != null && inspector.CurrentItem is Microsoft.Office.Interop.Outlook.MailItem mailItemInspector)
                    {
                        parentMailItem = mailItemInspector;
#if DEBUG
                        DebugLogger.Log("[OutlookImportService] Using ActiveInspector for child MSG extraction.");
#endif
                    }
                    else
                    {
                        var explorer = outlookApp.ActiveExplorer();
                        if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                        {
                            var selectedItem = explorer.Selection[1];
                            if (selectedItem is Microsoft.Office.Interop.Outlook.MailItem mailItemExplorer)
                            {
                                parentMailItem = mailItemExplorer;
#if DEBUG
                                DebugLogger.Log("[OutlookImportService] Using ActiveExplorer selection for child MSG extraction.");
#endif
                            }
                        }
                    }
                    if (parentMailItem != null)
                    {
                        if (draggedMsgHash != null)
                        {
                            bool found = false;
                            for (int i = 1; i <= parentMailItem.Attachments.Count; i++)
                            {
                                var attachment = parentMailItem.Attachments[i];
                                string attName = attachment.FileName?.Trim();
                                if (attachment.Type == Microsoft.Office.Interop.Outlook.OlAttachmentType.olEmbeddeditem && attName != null && attName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                                {
                                    // Save to temp file
                                    string tempPath = Path.GetTempFileName();
                                    string tempMsgPath = Path.ChangeExtension(tempPath, ".msg");
                                    try
                                    {
                                        attachment.SaveAsFile(tempMsgPath);
                                        // Compute hash
                                        using (var fileStream = File.OpenRead(tempMsgPath))
                                        using (var sha256 = System.Security.Cryptography.SHA256.Create())
                                        {
                                            var hash = sha256.ComputeHash(fileStream);
                                            string attHash = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();

#if DEBUG
                                            DebugLogger.Log($"[OutlookImportService] Attachment '{attName}' hash: {attHash}");
#endif
                                            if (attHash == draggedMsgHash)
                                            {
                                                // Found the correct attachment
                                                string safeFileName = sanitizeFileName(attName);
                                                string destPath = Path.Combine(outputFolder, safeFileName);
                                                int counter = 1;
                                                while (File.Exists(destPath))
                                                {
                                                    string nameWithoutExt = Path.GetFileNameWithoutExtension(safeFileName);
                                                    string ext = Path.GetExtension(safeFileName);
                                                    string uniqueFileName = $"{nameWithoutExt}_{counter}{ext}";
                                                    destPath = Path.Combine(outputFolder, uniqueFileName);
                                                    counter++;
                                                }
                                                File.Copy(tempMsgPath, destPath, true);
                                                result.ExtractedFiles.Add(destPath);

#if DEBUG
                                                DebugLogger.Log($"[OutlookImportService] Successfully saved child MSG (hash match): {destPath}");
#endif
                                                found = true;
                                                break;
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {

#if DEBUG
                                        DebugLogger.Log($"[OutlookImportService] Error saving/checking attachment '{attName}': {ex.Message}");
#endif
                                    }
                                    finally
                                    {
                                        try { if (File.Exists(tempMsgPath)) File.Delete(tempMsgPath); } catch { }
                                        try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                                    }
                                }
                            }
                            if (!found)
                            {
                                result.SkippedFiles.Add($"No matching child MSG found by hash for: {expectedFileName}");
                            }
                        }
                        else
                        {
                            // Fallback: extract all MSG attachments with matching name
                            int extractedCount = 0;
                            for (int i = 1; i <= parentMailItem.Attachments.Count; i++)
                            {
                                var attachment = parentMailItem.Attachments[i];
                                string attName = attachment.FileName?.Trim();
                                if (attachment.Type == Microsoft.Office.Interop.Outlook.OlAttachmentType.olEmbeddeditem && attName != null && attName.Equals(expectedFileName, StringComparison.OrdinalIgnoreCase))
                                {
                                    string tempPath = Path.GetTempFileName();
                                    string tempMsgPath = Path.ChangeExtension(tempPath, ".msg");
                                    try
                                    {
                                        attachment.SaveAsFile(tempMsgPath);
                                        string safeFileName = sanitizeFileName(attName);
                                        string destPath = Path.Combine(outputFolder, safeFileName);
                                        int counter = 1;
                                        while (File.Exists(destPath))
                                        {
                                            string nameWithoutExt = Path.GetFileNameWithoutExtension(safeFileName);
                                            string ext = Path.GetExtension(safeFileName);
                                            string uniqueFileName = $"{nameWithoutExt}_{counter}{ext}";
                                            destPath = Path.Combine(outputFolder, uniqueFileName);
                                            counter++;
                                        }
                                        File.Copy(tempMsgPath, destPath, true);
                                        result.ExtractedFiles.Add(destPath);

#if DEBUG
                                        DebugLogger.Log($"[OutlookImportService] Fallback: saved child MSG by name: {destPath}");
#endif
                                        extractedCount++;
                                    }
                                    catch (Exception ex)
                                    {
#if DEBUG
                                        DebugLogger.Log($"[OutlookImportService] Error saving fallback attachment '{attName}': {ex.Message}");
#endif
                                    }
                                    finally
                                    {
                                        try { if (File.Exists(tempMsgPath)) File.Delete(tempMsgPath); } catch { }
                                        try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
                                    }
                                }
                            }
                            if (extractedCount > 1)
                            {

#if DEBUG
                                DebugLogger.Log($"[OutlookImportService] Fallback: Multiple child MSGs with same name extracted: {extractedCount}");
#endif
                            }
                            if (extractedCount == 0)
                            {
                                result.SkippedFiles.Add($"No child MSG found by name for: {expectedFileName}");
                            }
                        }
                    }
                    else
                    {
                        result.SkippedFiles.Add("No active inspector, no valid explorer selection, or current item is not a MailItem");
                    }
                }
                else
                {
                    result.SkippedFiles.Add("Outlook app not found");
                }
            }
            catch (Exception ex)
            {
                result.SkippedFiles.Add("Child MSG could not be extracted: " + ex.Message);
            }

#if DEBUG
            DebugLogger.Log($"[OutlookImportService] Child MSG extraction complete. Found {result.ExtractedFiles.Count} files, skipped {result.SkippedFiles.Count}");
#endif
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
