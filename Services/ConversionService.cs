using MsgReader.Outlook;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace MsgToPdfConverter.Services
{
    public class ConversionService
    {
        public void ConvertMsgFilesToPdf(List<string> msgFilePaths, string outputDirectory)
        {
            foreach (var msgFilePath in msgFilePaths)
            {
                var email = new Storage.Message(msgFilePath);
                string pdfFileName = GenerateUniquePdfFileName(msgFilePath, outputDirectory, msgFilePaths);
                ConvertToPdf(email, pdfFileName);
            }
        }

        private void ConvertToPdf(Storage.Message email, string pdfFilePath)
        {
            using (var pdf = new PdfDocument())
            {
                var page = pdf.AddPage();
                using (var gfx = XGraphics.FromPdfPage(page))
                {
                    var font = new XFont("Verdana", 12);
                    double y = 20;
                    double x = 20;
                    double width = page.Width.Point - 40;
                    double height = 20;
                    gfx.DrawString($"Subject: {email.Subject}", font, XBrushes.Black, new XRect(x, y, width, height), XStringFormats.TopLeft); y += 25;
                    gfx.DrawString($"From: {email.Sender?.DisplayName}", font, XBrushes.Black, new XRect(x, y, width, height), XStringFormats.TopLeft); y += 25;
                    gfx.DrawString($"To: {email.GetEmailRecipients(0, false, false)}", font, XBrushes.Black, new XRect(x, y, width, height), XStringFormats.TopLeft); y += 25;
                    gfx.DrawString($"Date: {email.SentOn}", font, XBrushes.Black, new XRect(x, y, width, height), XStringFormats.TopLeft); y += 25;
                    gfx.DrawString($"Body:", font, XBrushes.Black, new XRect(x, y, width, height), XStringFormats.TopLeft); y += 25;
                    gfx.DrawString(email.BodyText, font, XBrushes.Black, new XRect(x, y, width, page.Height.Point - y - 20), XStringFormats.TopLeft);
                }
                pdf.Save(pdfFilePath);
            }
        }

        public (int Success, int Fail, int Processed, bool Cancelled) ConvertFilesWithAttachments(
            List<string> selectedFiles,
            string selectedOutputFolder,
            bool appendAttachments,
            bool extractOriginalOnly,
            bool deleteFilesAfterConversion,
            bool combineAllPdfs, // <--- new parameter
            EmailConverterService emailService,
            AttachmentService attachmentService,
            Action<int, int, int, string> updateProgress, // (processed, total, progress, statusText)
            Action<int, int> updateFileProgress, // (current, max) for per-file progress
            Func<bool> isCancellationRequested,
            Action<string> showMessageBox // (message)
            , List<string> generatedPdfs = null // optional: collect generated PDFs
        )
        {
            int success = 0, fail = 0, processed = 0;
            // Always use a single temp folder for all temp/intermediate files
            string baseTempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "MsgToPdfConverter");
            System.IO.Directory.CreateDirectory(baseTempDir);
            string sessionTempDir = baseTempDir;
            // Track all temp files created during this session
            var sessionTempFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                for (int i = 0; i < selectedFiles.Count; i++)
                {
                    if (isCancellationRequested())
                        break;
                    processed++;
                    string filePath = selectedFiles[i];
                    bool conversionSucceeded = false;
                    try
                    {


                        if (string.IsNullOrWhiteSpace(filePath))
                        {
                            showMessageBox($"[ERROR] Skipping null or empty file path at index {i}.");
                            fail++;
                            continue;
                        }
                        if (!System.IO.File.Exists(filePath))
                        {
                            showMessageBox($"[ERROR] File does not exist: {filePath}");
                            fail++;
                            continue;
                        }
                        string extCheck = System.IO.Path.GetExtension(filePath)?.ToLowerInvariant();
                        if (string.IsNullOrEmpty(extCheck))
                        {
                            showMessageBox($"[ERROR] Could not determine extension for: {filePath}");
                            fail++;
                            continue;
                        }
                        // ...existing code...
                    {
                        updateProgress(processed, selectedFiles.Count, i, $"Processing file {processed}/{selectedFiles.Count}: {System.IO.Path.GetFileName(selectedFiles[i])}");
                        string ext = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
                        string dir = !string.IsNullOrEmpty(selectedOutputFolder) ? selectedOutputFolder : System.IO.Path.GetDirectoryName(filePath);
                        string baseName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                        // Use sessionTempDir for all temp/intermediate files
                        if (ext == ".msg")
                        {
                            if (appendAttachments)
                            {
                                Storage.Message msg = null;
                                try
                                {
                                    try {
                                        msg = new Storage.Message(filePath);
                                    } catch (Exception msgEx) {
                                        showMessageBox($"[ERROR] Exception loading MSG file: {filePath} - {msgEx.Message}");
                                        fail++;
                                        continue;
                                    }

            if (msg == null)
            {
                showMessageBox($"[ERROR] Failed to load MSG file: {filePath}");
                fail++;
                continue;
            }
            if (appendAttachments && (msg.Attachments == null || msg.Attachments.Count == 0))
            {

                // Always convert the main message body to PDF
                string datePart_noatt = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
                string msgBaseName_noatt = System.IO.Path.GetFileNameWithoutExtension(filePath);
                string msgDir_noatt = baseTempDir;
                string pdfFilePath_noatt = System.IO.Path.Combine(msgDir_noatt, $"{msgBaseName_noatt} - {datePart_noatt}.pdf");
                string outputPdf_noatt = System.IO.Path.Combine(dir, $"{msgBaseName_noatt} - {datePart_noatt}.pdf");
                if (System.IO.File.Exists(pdfFilePath_noatt))
                    System.IO.File.Delete(pdfFilePath_noatt);
                var htmlResult_noatt = emailService.BuildEmailHtmlWithInlineImages(msg, false);
                if (emailService == null)
                {
                    showMessageBox($"[ERROR] emailService is null when building HTML for MSG file: {filePath}");
                    fail++;
                    continue;
                }
                if (string.IsNullOrEmpty(htmlResult_noatt.Html))
                {
                    showMessageBox($"[ERROR] Failed to build HTML for MSG file: {filePath}");
                    fail++;
                    continue;
                }
                string htmlWithHeader_noatt = htmlResult_noatt.Html;
                var tempHtmlPath_noatt = System.IO.Path.Combine(baseTempDir, Guid.NewGuid() + ".html");
                if (string.IsNullOrEmpty(tempHtmlPath_noatt))
                {
                    showMessageBox($"[ERROR] tempHtmlPath is null or empty for MSG file: {filePath}");
                    fail++;
                    continue;
                }
                if (string.IsNullOrEmpty(pdfFilePath_noatt))
                {
                    showMessageBox($"[ERROR] pdfFilePath is null or empty for MSG file: {filePath}");
                    fail++;
                    continue;
                }
                System.IO.File.WriteAllText(tempHtmlPath_noatt, htmlWithHeader_noatt, System.Text.Encoding.UTF8);
                sessionTempFiles.Add(tempHtmlPath_noatt);
                var inlineContentIds_noatt = emailService.GetInlineContentIds(htmlWithHeader_noatt) ?? new HashSet<string>();
                var psi_noatt = new System.Diagnostics.ProcessStartInfo();
                psi_noatt.FileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                psi_noatt.Arguments = $"--html2pdf \"{tempHtmlPath_noatt}\" \"{pdfFilePath_noatt}\"";
                psi_noatt.UseShellExecute = false;
                psi_noatt.CreateNoWindow = true;
                psi_noatt.RedirectStandardOutput = true;
                psi_noatt.RedirectStandardError = true;
                var proc_noatt = System.Diagnostics.Process.Start(psi_noatt);
                if (proc_noatt == null)
                {
                    showMessageBox($"[ERROR] Failed to start HtmlToPdfWorker process for MSG file: {filePath}");
                    fail++;
                    continue;
                }
                if (inlineContentIds_noatt == null)
                    showMessageBox($"[WARN] inlineContentIds is null for MSG file: {filePath}");
                proc_noatt.WaitForExit();
                System.IO.File.Delete(tempHtmlPath_noatt);
                if (proc_noatt.ExitCode != 0)
                {
                    showMessageBox($"[ERROR] HtmlToPdfWorker failed: {proc_noatt.StandardError.ReadToEnd()}");
                    fail++;
                    continue;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                // Always copy the final PDF to the output folder
                if (!string.Equals(pdfFilePath_noatt, outputPdf_noatt, StringComparison.OrdinalIgnoreCase) && System.IO.File.Exists(pdfFilePath_noatt))
                {
                    try
                    {
                        if (System.IO.File.Exists(outputPdf_noatt))
                            System.IO.File.Delete(outputPdf_noatt);
                        System.IO.File.Copy(pdfFilePath_noatt, outputPdf_noatt, true);

                        // Delete temp PDF after copying
                        try {
                            if (System.IO.File.Exists(pdfFilePath_noatt))
                                System.IO.File.Delete(pdfFilePath_noatt);
                            #if DEBUG
                            DebugLogger.Log($"[DEBUG] Deleted temp PDF: {pdfFilePath_noatt}");
                            #endif
                        } catch { }
                    }
                    catch (Exception moveEx)
                    {
                        #if DEBUG
                        DebugLogger.Log($"[ERROR] Failed to copy file: {moveEx.Message}");
                        #endif
                        outputPdf_noatt = pdfFilePath_noatt;
                    }
                }
                else if (System.IO.File.Exists(outputPdf_noatt))
                {
                    #if DEBUG
                    DebugLogger.Log($"[DEBUG] Output PDF already exists at {outputPdf_noatt}");
                    #endif
                }
                else if (System.IO.File.Exists(pdfFilePath_noatt))
                {
                    outputPdf_noatt = pdfFilePath_noatt;
                }
                generatedPdfs?.Add(outputPdf_noatt);
                success++;
                conversionSucceeded = true;
                updateFileProgress?.Invoke(100, 100);
                continue;
            }
            if (msg.Sender == null)
                showMessageBox($"[WARN] MSG file has null Sender: {filePath}");
            if (msg.SentOn == null)
                showMessageBox($"[WARN] MSG file has null SentOn: {filePath}");
            if (msg.BodyText == null)
                showMessageBox($"[WARN] MSG file has null BodyText: {filePath}");
            if (msg.Attachments == null)
                showMessageBox($"[WARN] MSG file has null Attachments: {filePath}");
            if (msg.Attachments != null && msg.Attachments.Count == 0)
                showMessageBox($"[INFO] MSG file has zero Attachments: {filePath}");
            int fileProgress = 0;
            int totalCount = 0;
            try
            {
                #if DEBUG
                DebugLogger.Log($"[DEBUG] Calling attachmentService.CountAllProcessableItems(msg)");
                #endif
                totalCount = attachmentService.CountAllProcessableItems(msg);
                #if DEBUG
                DebugLogger.Log($"[DEBUG] attachmentService.CountAllProcessableItems(msg) returned: {totalCount}");
                #endif
            }
            catch (Exception attCountEx)
            {
                #if DEBUG
                DebugLogger.Log($"[ERROR] Exception in attachmentService.CountAllProcessableItems: {attCountEx.Message}");
                #endif
                showMessageBox($"[ERROR] Could not count processable items: {attCountEx.Message}");
                totalCount = 1;
            }
            updateFileProgress?.Invoke(0, Math.Max(totalCount, 1));
                                    string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
                                    string msgBaseName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                                    string msgDir = baseTempDir;
                                    string uniquePdfName = GenerateUniquePdfFileName(filePath, msgDir, selectedFiles);
                                    string pdfBaseName = System.IO.Path.GetFileNameWithoutExtension(uniquePdfName);
                                    string pdfFilePath = System.IO.Path.Combine(msgDir, $"{pdfBaseName} - {datePart}.pdf");
                                    string outputPdf = GenerateUniquePdfFileName(filePath, dir, selectedFiles);
                                    if (System.IO.File.Exists(pdfFilePath))
                                        System.IO.File.Delete(pdfFilePath);
                                    var htmlResult = emailService.BuildEmailHtmlWithInlineImages(msg, false);
                                    if (emailService == null)
                                    {
                                        showMessageBox($"[ERROR] emailService is null when building HTML for MSG file: {filePath}");
                                        fail++;
                                        continue;
                                    }
                                    if (string.IsNullOrEmpty(htmlResult.Html))
                                    {
                                        showMessageBox($"[ERROR] Failed to build HTML for MSG file: {filePath}");
                                        fail++;
                                        continue;
                                    }
                                    string htmlWithHeader = htmlResult.Html;
                                    var tempHtmlPath = System.IO.Path.Combine(baseTempDir, Guid.NewGuid() + ".html");
                                    if (string.IsNullOrEmpty(tempHtmlPath))
                                    {
                                        showMessageBox($"[ERROR] tempHtmlPath is null or empty for MSG file: {filePath}");
                                        fail++;
                                        continue;
                                    }
                                    if (string.IsNullOrEmpty(pdfFilePath))
                                    {
                                        showMessageBox($"[ERROR] pdfFilePath is null or empty for MSG file: {filePath}");
                                        fail++;
                                        continue;
                                    }
                                    System.IO.File.WriteAllText(tempHtmlPath, htmlWithHeader, System.Text.Encoding.UTF8);
                                    sessionTempFiles.Add(tempHtmlPath);
                                    var inlineContentIds = emailService.GetInlineContentIds(htmlWithHeader) ?? new HashSet<string>();
                                    var psi = new System.Diagnostics.ProcessStartInfo();
                                    psi.FileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                                    psi.Arguments = $"--html2pdf \"{tempHtmlPath}\" \"{pdfFilePath}\"";
                                    psi.UseShellExecute = false;
                                    psi.CreateNoWindow = true;
                                    psi.RedirectStandardOutput = true;
                                    psi.RedirectStandardError = true;
                                    var proc = System.Diagnostics.Process.Start(psi);
                                    if (proc == null)
                                    {
                                        showMessageBox($"[ERROR] Failed to start HtmlToPdfWorker process for MSG file: {filePath}");
                                        fail++;
                                        continue;
                                    }
                                    if (inlineContentIds == null)
                                        showMessageBox($"[WARN] inlineContentIds is null for MSG file: {filePath}");
                                    proc.WaitForExit();
                                    System.IO.File.Delete(tempHtmlPath);
                                    if (proc.ExitCode != 0)
                                    {
                                        showMessageBox($"[ERROR] HtmlToPdfWorker failed: {proc.StandardError.ReadToEnd()}");
                                        fail++;
                                        continue;
                                    }
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                // --- Attachment/merge logic ---
                                var typedAttachments = new List<Storage.Attachment>();
                                var nestedMessages = new List<Storage.Message>();
                                if (msg.Attachments != null && msg.Attachments.Count > 0)
                                {
                                    // First, collect all non-inline attachments
                                    var allAttachments = new List<Storage.Attachment>();
                                    foreach (var att in msg.Attachments)
                                    {
                                        if (att == null)
                                        {
                                            showMessageBox($"[ERROR] Null attachment found in MSG file: {filePath}");
                                            continue;
                                        }
                                        if (att is Storage.Attachment a)
                                        {
                                            string ext2 = System.IO.Path.GetExtension(a.FileName ?? "").ToLowerInvariant();
                                            bool isImage = ext2 == ".jpg" || ext2 == ".jpeg" || ext2 == ".png" || ext2 == ".gif" || ext2 == ".bmp";
                                            if (isImage)
                                            {
                                                if (a.IsInline == true || (!string.IsNullOrEmpty(a.ContentId) && inlineContentIds.Contains(a.ContentId.Trim('<', '>', '"', '\'', ' '))))
                                                    continue;
                                                if (attachmentService == null)
                                                {
                                                    showMessageBox($"[ERROR] attachmentService is null when checking signature image for: {a.FileName}");
                                                    continue;
                                                }
                                                if (attachmentService.IsLikelySignatureImage(a))
                                                    continue;
                                            }
                                            allAttachments.Add(a);
                                        }
                                        else if (att is Storage.Message nestedMsg)
                                        {
                                            if (nestedMsg == null)
                                            {
                                                showMessageBox($"[ERROR] Null nested message found in MSG file: {filePath}");
                                                continue;
                                            }
                                            nestedMessages.Add(nestedMsg);
                                        }
                                    }

                                    // DEDUPLICATION: Group by base filename and prefer Office files over PDFs
                                    var attachmentsToSkip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                    var duplicateGroups = allAttachments
                                        .GroupBy(a => Path.GetFileNameWithoutExtension(a.FileName ?? ""), StringComparer.OrdinalIgnoreCase)
                                        .Where(g => g.Count() > 1)
                                        .ToList();

                                    foreach (var group in duplicateGroups)
                                    {
                                        var groupList = group.ToList();
                                        #if DEBUG
                                        DebugLogger.Log($"[DEDUP] Found {groupList.Count} files with base name '{group.Key}':");
                                        #endif
                                        
                                        foreach (var att in groupList)
                                        {
                                        #if DEBUG
                                        DebugLogger.Log($"[DEDUP]   {att.FileName}");
                                        #endif
                                        }

                                        var officeFiles = groupList.Where(a => {
                                            var fileExt = Path.GetExtension(a.FileName ?? "").ToLowerInvariant();
                                            return fileExt == ".doc" || fileExt == ".docx" || fileExt == ".xls" || fileExt == ".xlsx";
                                        }).ToList();

                                        var pdfFiles = groupList.Where(a => {
                                            var fileExt = Path.GetExtension(a.FileName ?? "").ToLowerInvariant();
                                            return fileExt == ".pdf";
                                        }).ToList();

                                        // If we have both Office and PDF files, prefer Office files (they may contain embedded objects)
                                        if (officeFiles.Count > 0 && pdfFiles.Count > 0)
                                        {
                                            // Keep the first Office file, skip all PDF files
                                            var keepOfficeFile = officeFiles.First();
                                            #if DEBUG
                                            DebugLogger.Log($"[DEDUP] Keeping Office file: {keepOfficeFile.FileName}");
                                            #endif
                                            
                                            foreach (var pdfFile in pdfFiles)
                                            {
                                                #if DEBUG
                                                DebugLogger.Log($"[DEDUP] Skipping PDF duplicate: {pdfFile.FileName}");
                                                #endif
                                                attachmentsToSkip.Add(pdfFile.FileName ?? "");
                                            }
                                            
                                            // If there are multiple Office files, keep only the first one
                                            for (int j = 1; j < officeFiles.Count; j++)
                                            {
                                                #if DEBUG
                                                DebugLogger.Log($"[DEDUP] Skipping duplicate Office file: {officeFiles[j].FileName}");
                                                #endif
                                                attachmentsToSkip.Add(officeFiles[j].FileName ?? "");
                                            }
                                        }
                                        else if (officeFiles.Count > 1)
                                        {
                                            // Multiple Office files with same base name - keep first one
                                            var keepOfficeFile = officeFiles.First();
                                            #if DEBUG
                                            DebugLogger.Log($"[DEDUP] Keeping first Office file: {keepOfficeFile.FileName}");
                                            #endif
                                            
                                            for (int j = 1; j < officeFiles.Count; j++)
                                            {
                                                #if DEBUG
                                                DebugLogger.Log($"[DEDUP] Skipping duplicate Office file: {officeFiles[j].FileName}");
                                                #endif
                                                attachmentsToSkip.Add(officeFiles[j].FileName ?? "");
                                            }
                                        }
                                        else if (pdfFiles.Count > 1)
                                        {
                                            // Multiple PDF files with same base name - keep first one
                                            var keepPdfFile = pdfFiles.First();
                                            #if DEBUG
                                            DebugLogger.Log($"[DEDUP] Keeping first PDF file: {keepPdfFile.FileName}");
                                            #endif
                                            
                                            for (int j = 1; j < pdfFiles.Count; j++)
                                            {
                                                #if DEBUG
                                                DebugLogger.Log($"[DEDUP] Skipping duplicate PDF file: {pdfFiles[j].FileName}");
                                                #endif
                                                attachmentsToSkip.Add(pdfFiles[j].FileName ?? "");
                                            }
                                        }
                                    }

                                    // Build final list, skipping duplicates
                                    foreach (var a in allAttachments)
                                    {
                                        if (attachmentsToSkip.Contains(a.FileName ?? ""))
                                        {
                                            #if DEBUG
                                            DebugLogger.Log($"[DEDUP] SKIPPING (as planned): {a.FileName}");
                                            #endif
                                            continue;
                                        }

                                        #if DEBUG
                                        DebugLogger.Log($"[DEDUP] Including attachment for processing: {a.FileName}");
                                        #endif
                                        typedAttachments.Add(a);
                                    }
                                }
                                var allPdfFiles = new List<string> { pdfFilePath };
                                var allTempFiles = new List<string>();
                                string tempDir = sessionTempDir;
                                for (int attIndex = 0; attIndex < typedAttachments.Count; attIndex++)
                                {
                                    var att = typedAttachments[attIndex];
                                    string attName = att.FileName ?? "attachment";
                                    string attPath = System.IO.Path.Combine(tempDir, attName);
                                    string headerText = $"Attachment {attIndex + 1}/{typedAttachments.Count} - {attName}";
                                    try
                                    {
                                        System.IO.File.WriteAllBytes(attPath, att.Data);
                                        allTempFiles.Add(attPath);
                                        sessionTempFiles.Add(attPath);
                                        var attachmentParentChain = new List<string> { msg.Subject ?? System.IO.Path.GetFileName(filePath) };
                                if (msg.Subject == null)
                                    showMessageBox($"[WARN] MSG file has null Subject: {filePath}");
                                        // Pass progressTick to ensure every processed file increments progress
                                        string finalAttachmentPdf = attachmentService.ProcessSingleAttachmentWithHierarchy(att, attPath, tempDir, headerText, allTempFiles, allPdfFiles, attachmentParentChain, attName, false, () => updateFileProgress?.Invoke(++fileProgress, Math.Max(totalCount, 1)));
                                        if (attachmentService == null)
                                        {
                                            showMessageBox($"[ERROR] attachmentService is null when processing attachment: {attName}");
                                            continue;
                                        }
                                    if (emailService == null)
                                    {
                                        showMessageBox($"[ERROR] emailService is null when building HTML for MSG file: {filePath}");
                                        fail++;
                                        continue;
                                    }
                                        if (finalAttachmentPdf != null)
                                            allPdfFiles.Add(finalAttachmentPdf);
                                    }
                                    catch (Exception ex)
                                    {
                                        string errorPdf = System.IO.Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                                        var errorParentChain = new List<string> { msg.Subject ?? System.IO.Path.GetFileName(filePath) };
                                        string enhancedErrorText = MsgToPdfConverter.Utils.TreeHeaderHelper.BuildTreeHeader(errorParentChain, attName) + "\n\n" + headerText + $"\n(Error: {ex.Message})";
                                        PdfService.AddHeaderPdf(errorPdf, enhancedErrorText);
                                        allPdfFiles.Add(errorPdf);
                                        allTempFiles.Add(errorPdf);
                                        sessionTempFiles.Add(errorPdf);
                                    }
                                }
                                for (int nestedIndex = 0; nestedIndex < nestedMessages.Count; nestedIndex++)
                                {
                                    var nestedMsg = nestedMessages[nestedIndex];
                                    string nestedSubject = nestedMsg.Subject ?? $"nested_msg_depth_1";
                                    string nestedHeaderText = $"Attachment (Depth 1): {nestedIndex + 1}/{nestedMessages.Count} - Nested Email: {nestedSubject}";
                                    var initialParentChain = new List<string> { msg.Subject ?? System.IO.Path.GetFileName(filePath) };
                                    // When calling ProcessMsgAttachmentsRecursively, pass a lambda to increment progress:
                                    attachmentService.ProcessMsgAttachmentsRecursively(nestedMsg, allPdfFiles, allTempFiles, tempDir, false, 1, 5, nestedHeaderText, initialParentChain, () => updateFileProgress?.Invoke(++fileProgress, Math.Max(totalCount, 1)));
                                }
                                string mergedPdf = System.IO.Path.Combine(tempDir, System.IO.Path.GetFileNameWithoutExtension(pdfFilePath) + "_merged.pdf");
                                // Restore sent date in output PDF filename for merged PDFs
                                string datePart_att = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
                                string pdfBaseName_att = System.IO.Path.GetFileNameWithoutExtension(uniquePdfName);
                                string mergedPdfWithDate = System.IO.Path.Combine(tempDir, $"{pdfBaseName_att} - {datePart_att}_merged.pdf");
                                PdfAppendTest.AppendPdfs(allPdfFiles, mergedPdfWithDate);
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                // Build list of files to protect: all user-selected source files and all output files
                                var filesToProtect = new HashSet<string>(selectedFiles.Select(f => Path.GetFullPath(f).TrimEnd(Path.DirectorySeparatorChar).ToLowerInvariant()));
                                filesToProtect.Add(Path.GetFullPath(pdfFilePath).TrimEnd(Path.DirectorySeparatorChar).ToLowerInvariant());
                                filesToProtect.Add(Path.GetFullPath(mergedPdfWithDate).TrimEnd(Path.DirectorySeparatorChar).ToLowerInvariant());
                                if (generatedPdfs != null)
                                {
                                    foreach (var genPdf in generatedPdfs)
                                    {
                                        filesToProtect.Add(Path.GetFullPath(genPdf).TrimEnd(Path.DirectorySeparatorChar).ToLowerInvariant());
                                    }
                                }

                                foreach (var prot in filesToProtect) Console.WriteLine($"  {prot}");
                                // DEBUG: Print temp and protected files before cleanup
                                AttachmentService.DebugPrintTempAndProtectedFiles(allTempFiles, filesToProtect);
                                foreach (var f in allTempFiles)
                                {
                                    try
                                    {
                                        var normF = Path.GetFullPath(f).TrimEnd(Path.DirectorySeparatorChar).ToLowerInvariant();
                                        if (System.IO.File.Exists(f) && !filesToProtect.Contains(normF))
                                        {
                                            FileService.RobustDeleteFile(f);
                                        }
                                        else if (System.IO.Directory.Exists(f))
                                        {
                                            System.IO.Directory.Delete(f, true);
                                        }
                                    }
                                    catch { }
                                }
                                if (System.IO.File.Exists(mergedPdfWithDate))
                                {
                                    if (System.IO.File.Exists(pdfFilePath))
                                        System.IO.File.Delete(pdfFilePath);
                                    System.IO.File.Move(mergedPdfWithDate, pdfFilePath);
                                }
                                // Always copy the final PDF to the output folder, using the sent date in the filename
                                // Decide output filename: always include sent date, only use _merged if >1 PDF
                                // Always use [BaseName] - [SentDate].pdf for output
                                string outputFileName = $"{pdfBaseName_att} - {datePart_att}.pdf";
                                string outputPdfFinal = System.IO.Path.Combine(dir, outputFileName);
                                if (!string.Equals(pdfFilePath, outputPdfFinal, StringComparison.OrdinalIgnoreCase) && System.IO.File.Exists(pdfFilePath))
                                {
                                    try
                                    {
                                        if (System.IO.File.Exists(outputPdfFinal))
                                            System.IO.File.Delete(outputPdfFinal);
                                        System.IO.File.Copy(pdfFilePath, outputPdfFinal, true);

                                        // Delete temp PDF after copying
                                        try {
                                            if (System.IO.File.Exists(pdfFilePath))
                                                System.IO.File.Delete(pdfFilePath);

                                        } catch { }
                                    }
                                    catch (Exception)
                                    {

                                        outputPdfFinal = pdfFilePath;
                                    }
                                }
                                else if (System.IO.File.Exists(outputPdfFinal))
                                {
                                    #if DEBUG
                                    DebugLogger.Log($"[DEBUG] Output PDF already exists at {outputPdfFinal}");
                                    #endif
                                }
                                else if (System.IO.File.Exists(pdfFilePath))
                                {
                                    outputPdfFinal = pdfFilePath;
                                }
                                generatedPdfs?.Add(outputPdfFinal);
                                success++;
                                conversionSucceeded = true;
                                // Mark file progress as complete
                                updateFileProgress?.Invoke(totalCount, totalCount);
                            }
                            catch (Exception ex)
                            {
                                showMessageBox($"Error processing {System.IO.Path.GetFileName(filePath)}: {ex.Message}");
                                fail++;
                            }
                            finally
                            {
                                if (msg != null && msg is IDisposable disposableMsg)
                                    disposableMsg.Dispose();
                                msg = null;
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                if (deleteFilesAfterConversion && conversionSucceeded && !combineAllPdfs) // <--- updated condition
                                {
                                    #if DEBUG
                                    DebugLogger.Log($"[DELETE] Should delete: {filePath}, deleteFilesAfterConversion={deleteFilesAfterConversion}, combineAllPdfs={combineAllPdfs}, appendAttachments={appendAttachments}");
                                    #endif
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        #if DEBUG
                                        DebugLogger.Log($"[DELETE] Attempting to delete: {filePath}");
                                        #endif
                                        try { FileService.MoveFileToRecycleBin(filePath); }
                                        catch (Exception ex)
                                        {
                                            #if DEBUG
                                            DebugLogger.Log($"[DELETE] Failed: {ex.Message}");
                                            #endif
                                        }
                                    }
                                    else
                                    {
                                        #if DEBUG
                                        DebugLogger.Log($"[DELETE] File not found: {filePath}");
                                        #endif
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Only convert the email body (no attachments)
                            updateFileProgress?.Invoke(0, 100);
                            
                            // Add intermediate progress updates for better UX
                            updateFileProgress?.Invoke(25, 100); // Loading MSG file
                            var result = ConvertSingleMsgFileWithProgress(filePath, dir, appendAttachments, extractOriginalOnly, emailService, attachmentService, generatedPdfs, selectedFiles, updateFileProgress);
                            if (result) { 
                                success++; 
                                conversionSucceeded = true; 
                                updateFileProgress?.Invoke(100, 100);
                            } else {
                                fail++;
                            }
                        }
                    }
                    else
                    {
                        // Use the same hierarchical logic as for attachments
                        // Reset file progress for non-MSG files and set up progress tracking
                        int fileProgress = 0;
                        int totalCount;
                        string extLower = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
                        if (extLower == ".doc" || extLower == ".docx" || extLower == ".xls" || extLower == ".xlsx")
                        {
                            totalCount = attachmentService.CountAllProcessableItemsFromFile(filePath);
                        }
                        else
                        {
                            totalCount = attachmentService.CountAllProcessableItemsFromFile(filePath);
                        }
                        updateFileProgress?.Invoke(0, Math.Max(totalCount, 1));

                        string outputPdf = GenerateUniquePdfFileName(filePath, dir, selectedFiles);
                        // Use only the main temp folder for all attachment processing
                        string tempDir = baseTempDir;
                        var tempFiles = new List<string>();
                        var allPdfFiles = new List<string>();
                        var allTempFiles = new List<string>();
                        string headerText = $"File: {System.IO.Path.GetFileName(filePath)}";
                        var parentChain = new List<string> { System.IO.Path.GetFileName(filePath) };
                        string processedPdf = null;
                        // Progress tick lambda for ZIP/7z: increment fileProgress and update UI
                        Action progressTick = () => updateFileProgress?.Invoke(++fileProgress, Math.Max(totalCount, 1));
                        try
                        {
                            processedPdf = attachmentService.ProcessSingleAttachmentWithHierarchy(
                                null, filePath, tempDir, headerText, allTempFiles, allPdfFiles, parentChain, baseName, extractOriginalOnly,
                                progressTick);
                            // Track all temp files created in allTempFiles
                            foreach (var tempF in allTempFiles)
                                sessionTempFiles.Add(tempF);
                            if (!string.IsNullOrEmpty(processedPdf))
                                allPdfFiles.Add(processedPdf);
                        }
                        catch (Exception ex)
                        {
                            showMessageBox($"Error processing {System.IO.Path.GetFileName(filePath)}: {ex.Message}");
                                #if DEBUG
                                DebugLogger.Log($"[ERROR] ProcessSingleAttachmentWithHierarchy failed: {ex}");
                                #endif
                            processedPdf = null;
                        }
                        // Robust existence check with retries for all generated PDFs
                        for (int pdfIdx = 0; pdfIdx < allPdfFiles.Count; pdfIdx++)
                        {
                            string pdfPath = allPdfFiles[pdfIdx];
                            int retryCount = 0;
                            while (!string.IsNullOrEmpty(pdfPath) && !System.IO.File.Exists(pdfPath) && retryCount < 5)
                            {
                                System.Threading.Thread.Sleep(200);
                                retryCount++;
                            }
                        }
                        // Log page counts and robustly filter zero-page PDFs before merging
                        var validPdfFiles = new List<string>();
                        foreach (var pdfPath in allPdfFiles)
                        {
                            try
                            {
                                using (var pdfDoc = PdfSharp.Pdf.IO.PdfReader.Open(pdfPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
                                {
                                    #if DEBUG
                                    DebugLogger.Log($"[DEBUG] PDF: {pdfPath}, PageCount: {pdfDoc.PageCount}");
                                    #endif
                                    if (pdfDoc.PageCount > 0)
                                    {
                                        validPdfFiles.Add(pdfPath);
                                    }
                                    else
                                    {
                                        #if DEBUG
                                        DebugLogger.Log($"[ERROR] PDF has zero pages: {pdfPath}");
                                        #endif
                                        showMessageBox($"Error: PDF generated from '{System.IO.Path.GetFileName(filePath)}' has no pages and will be skipped.");
                                    }
                                }
                            }
                            catch (Exception pdfEx)
                            {
                                #if DEBUG
                                DebugLogger.Log($"[ERROR] Failed to open PDF for page count check: {pdfEx.Message}");
                                #endif
                                showMessageBox($"Error: PDF generated from '{System.IO.Path.GetFileName(filePath)}' could not be opened and will be skipped.");
                            }
                        }
                        // Merge if more than one valid PDF, else just move/rename
                        string finalPdf = null;
                        if (validPdfFiles.Count > 1)
                        {
                            string mergedPdf = System.IO.Path.Combine(tempDir, baseName + "_merged.pdf");
                            PdfAppendTest.AppendPdfs(validPdfFiles, mergedPdf);
                            finalPdf = mergedPdf;
                        }
                        else if (validPdfFiles.Count == 1)
                        {
                            finalPdf = validPdfFiles[0];
                        }
                        else
                        {
                            #if DEBUG
                            DebugLogger.Log($"[ERROR] No valid PDFs to merge for {System.IO.Path.GetFileName(filePath)}");
                            #endif
                        }
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] finalPdf: {finalPdf}, outputPdf: {outputPdf}");
                        #endif
                        // Always move/copy the final PDF to the output folder if it exists
                        if (!string.IsNullOrEmpty(finalPdf) && System.IO.File.Exists(finalPdf))
                        {
                            try
                            {
                                if (System.IO.File.Exists(outputPdf))
                                    System.IO.File.Delete(outputPdf);
                                System.IO.File.Copy(finalPdf, outputPdf, true);
                                #if DEBUG
                                DebugLogger.Log($"[DEBUG] Copied {finalPdf} to {outputPdf}");
                                #endif
                            }
                            catch (Exception moveEx)
                            {
                            #if DEBUG
                            DebugLogger.Log($"[ERROR] Failed to copy file: {moveEx.Message}");
                            #endif
                                outputPdf = finalPdf;
                            }
                        }
                        else if (System.IO.File.Exists(outputPdf))
                        {
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] Output PDF already exists at {outputPdf}");
                        #endif
                        }
                        else if (!string.IsNullOrEmpty(finalPdf) && System.IO.File.Exists(finalPdf))
                        {
                            // If outputPdf does not exist but finalPdf does, treat as output
                            outputPdf = finalPdf;
                        }
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] Checking existence of output PDF: {outputPdf}");
                        #endif
                        if (System.IO.File.Exists(outputPdf))
                        {
                            generatedPdfs?.Add(outputPdf);
                            success++;
                            conversionSucceeded = true;
                            // Mark file progress as complete
                            updateFileProgress?.Invoke(totalCount, totalCount);
                        }
                        else
                        {
                        #if DEBUG
                        DebugLogger.Log($"[ERROR] No PDF generated for {System.IO.Path.GetFileName(filePath)}");
                        #endif
                            fail++;
                        }
                        // Cleanup temp files, but never delete the output PDF
                        // Do NOT delete temp files here; cleanup will happen at the end of the batch after all files are processed.
                    }
                }
            }
                catch (Exception ex)
                {
                    showMessageBox($"Error processing {System.IO.Path.GetFileName(filePath)}: {ex.Message}");
                    fail++;
                }
                finally
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    // Centralized deletion logic for all file types
                    if (deleteFilesAfterConversion && conversionSucceeded && !combineAllPdfs) // <--- updated condition
                    {
                        bool isPdf = string.Equals(Path.GetExtension(filePath), ".pdf", StringComparison.OrdinalIgnoreCase);
                        bool sameFolder = string.Equals(Path.GetDirectoryName(filePath)?.TrimEnd(Path.DirectorySeparatorChar),
                                                        (selectedOutputFolder ?? Path.GetDirectoryName(filePath))?.TrimEnd(Path.DirectorySeparatorChar),
                                                        StringComparison.OrdinalIgnoreCase);
                        if (isPdf && sameFolder)
                        {
                    #if DEBUG
                    DebugLogger.Log($"[DELETE] Skipping deletion of source PDF in output folder: {filePath}");
                    #endif
                        }
                        else if (System.IO.File.Exists(filePath))
                        {
                    #if DEBUG
                    DebugLogger.Log($"[DELETE] Attempting to delete: {filePath}");
                    #endif
                    try { FileService.MoveFileToRecycleBin(filePath); } catch (Exception ex) { 
                        #if DEBUG
                        DebugLogger.Log($"[DELETE] Failed: {ex.Message}");
                        #endif
                    }
                        }
                        else
                        {
                    #if DEBUG
                    DebugLogger.Log($"[DELETE] File not found: {filePath}");
                    #endif
                        }
                    }
                }
                }
            }
            finally
            {
                // Delete all files and subfolders in baseTempDir after conversion
                try
                {
                    if (System.IO.Directory.Exists(baseTempDir))
                    {
                        foreach (var file in System.IO.Directory.GetFiles(baseTempDir, "*", SearchOption.AllDirectories))
                        {
                            try { System.IO.File.Delete(file); } catch { }
                        }
                        foreach (var dir in System.IO.Directory.GetDirectories(baseTempDir, "*", SearchOption.AllDirectories).OrderByDescending(d => d.Length))
                        {
                            try { System.IO.Directory.Delete(dir, true); } catch { }
                        }
                        try { System.IO.Directory.Delete(baseTempDir, true); } catch { }
                    }
                }
                catch { }
            }
            #if DEBUG
            DebugLogger.Log("[DEBUG] ConversionService.ConvertFilesWithAttachments: Finished and returning.");
            #endif
            // Removed GC.Collect and GC.WaitForPendingFinalizers here to prevent freeze
            return (success, fail, processed, isCancellationRequested());
        }

        private bool ConvertSingleMsgFile(string msgFilePath, string outputDir, bool appendAttachments, bool extractOriginalOnly, EmailConverterService emailService, AttachmentService attachmentService, List<string> generatedPdfs, List<string> selectedFiles)
        {
            Storage.Message msg = new Storage.Message(msgFilePath);
            string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
            
            // Generate unique PDF filename to avoid conflicts when files have the same base name but different extensions
            string uniquePdfName = GenerateUniquePdfFileName(msgFilePath, outputDir, selectedFiles);
            string baseName = System.IO.Path.GetFileNameWithoutExtension(uniquePdfName);
            string pdfFilePath = System.IO.Path.Combine(outputDir, $"{baseName} - {datePart}.pdf");
            if (System.IO.File.Exists(pdfFilePath))
                System.IO.File.Delete(pdfFilePath);
            var htmlResult = emailService.BuildEmailHtmlWithInlineImages(msg, false);
            string htmlWithHeader = htmlResult.Html;
            var tempHtmlPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".html");
            System.IO.File.WriteAllText(tempHtmlPath, htmlWithHeader, System.Text.Encoding.UTF8);
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName,
                Arguments = $"--html2pdf \"{tempHtmlPath}\" \"{pdfFilePath}\"",
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            var proc = System.Diagnostics.Process.Start(psi);
            proc.WaitForExit();
            System.IO.File.Delete(tempHtmlPath);
            if (proc.ExitCode != 0)
                throw new Exception($"HtmlToPdfWorker failed");
            generatedPdfs?.Add(pdfFilePath);
            return true;
        }

        private bool ConvertSingleMsgFileWithProgress(string msgFilePath, string outputDir, bool appendAttachments, bool extractOriginalOnly, EmailConverterService emailService, AttachmentService attachmentService, List<string> generatedPdfs, List<string> selectedFiles, Action<int, int> updateFileProgress)
        {
            try
            {
                updateFileProgress?.Invoke(10, 100); // Starting MSG processing
                Storage.Message msg = new Storage.Message(msgFilePath);
                updateFileProgress?.Invoke(30, 100); // MSG file loaded
                string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
                string uniquePdfName = GenerateUniquePdfFileName(msgFilePath, outputDir, selectedFiles);
                string baseName = System.IO.Path.GetFileNameWithoutExtension(uniquePdfName);
                string pdfFilePath = System.IO.Path.Combine(outputDir, $"{baseName} - {datePart}.pdf");
                if (System.IO.File.Exists(pdfFilePath))
                    System.IO.File.Delete(pdfFilePath);

                updateFileProgress?.Invoke(50, 100); // Converting to HTML
                var htmlResult = emailService.BuildEmailHtmlWithInlineImages(msg, false);
                string htmlWithHeader = htmlResult.Html;
                var tempHtmlPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".html");
                System.IO.File.WriteAllText(tempHtmlPath, htmlWithHeader, System.Text.Encoding.UTF8);

                // Check for referenced temp files (attachments/images) and log missing ones
                if (htmlResult.TempFiles != null)
                {
                    foreach (var tempFile in htmlResult.TempFiles)
                    {
                        if (!System.IO.File.Exists(tempFile))
                        {
                            #if DEBUG
                            DebugLogger.Log($"Warning: Referenced attachment or image missing: {tempFile}. Skipping this file.");
                            #endif
                            // Optionally, remove the reference from HTML if needed
                        }
                    }
                }

                updateFileProgress?.Invoke(70, 100); // Starting PDF conversion
                var psi = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName,
                    Arguments = $"--html2pdf \"{tempHtmlPath}\" \"{pdfFilePath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                var proc = System.Diagnostics.Process.Start(psi);

                updateFileProgress?.Invoke(85, 100); // PDF conversion in progress
                proc.WaitForExit();
                System.IO.File.Delete(tempHtmlPath);

                if (proc.ExitCode != 0)
                    throw new Exception($"HtmlToPdfWorker failed");

                updateFileProgress?.Invoke(95, 100); // PDF conversion completed
                generatedPdfs?.Add(pdfFilePath);
                return true;
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"Error in ConvertSingleMsgFileWithProgress: {ex.Message}");
                #endif
                return false;
            }
        }

        /// <summary>
        /// Generates a PDF filename that avoids conflicts when multiple files have the same base name but different extensions.
        /// If there are conflicts (e.g., a.doc and a.xlsx), it will generate a.doc.pdf and a.xlsx.pdf instead of a.pdf for both.
        /// </summary>
        private string GenerateUniquePdfFileName(string filePath, string outputDir, List<string> allSelectedFiles)
        {
            string fileName = Path.GetFileName(filePath);
            string baseNameWithoutExt = Path.GetFileNameWithoutExtension(filePath);
            string originalExt = Path.GetExtension(filePath);
            
            // Check if there are other files in the list with the same base name but different extensions
            bool hasConflict = allSelectedFiles.Any(otherFile => 
                !string.Equals(otherFile, filePath, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(Path.GetFileNameWithoutExtension(otherFile), baseNameWithoutExt, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(Path.GetExtension(otherFile), originalExt, StringComparison.OrdinalIgnoreCase)
            );
            
            if (hasConflict)
            {
                // Include the original extension in the PDF name to avoid conflicts
                // Remove the dot from the extension for cleaner naming: a.doc.pdf instead of a..doc.pdf
                string cleanOriginalExt = originalExt.TrimStart('.');
                return Path.Combine(outputDir, $"{baseNameWithoutExt}.{cleanOriginalExt}.pdf");
            }
            else
            {
                // No conflict, use the standard naming
                return Path.Combine(outputDir, $"{baseNameWithoutExt}.pdf");
            }
        }
    }
}
