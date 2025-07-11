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
            for (int i = 0; i < selectedFiles.Count; i++)
            {
                if (isCancellationRequested())
                    break;
                processed++;
                string filePath = selectedFiles[i];
                bool conversionSucceeded = false;
                try
                {
                    updateProgress(processed, selectedFiles.Count, i, $"Processing file {processed}/{selectedFiles.Count}: {System.IO.Path.GetFileName(selectedFiles[i])}");
                    string ext = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
                    string dir = !string.IsNullOrEmpty(selectedOutputFolder) ? selectedOutputFolder : System.IO.Path.GetDirectoryName(filePath);
                    string baseName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                    if (ext == ".msg")
                    {
                        if (appendAttachments)
                        {
                            Storage.Message msg = null;
                            try
                            {
                                msg = new Storage.Message(filePath);
                                // --- Set up progress max for this file ---
                                int fileProgress = 0;
                                int totalCount = attachmentService.CountAllProcessableItems(msg);
                                updateFileProgress?.Invoke(0, Math.Max(totalCount, 1));
                                // ---
                                string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
                                string msgBaseName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                                string msgDir = !string.IsNullOrEmpty(selectedOutputFolder) ? selectedOutputFolder : System.IO.Path.GetDirectoryName(filePath);
                                
                                // Generate unique PDF filename to avoid conflicts when files have same base name but different extensions
                                string uniquePdfName = GenerateUniquePdfFileName(filePath, msgDir, selectedFiles);
                                string pdfBaseName = System.IO.Path.GetFileNameWithoutExtension(uniquePdfName);
                                string pdfFilePath = System.IO.Path.Combine(msgDir, $"{pdfBaseName} - {datePart}.pdf");
                                if (System.IO.File.Exists(pdfFilePath))
                                    System.IO.File.Delete(pdfFilePath);
                                var htmlResult = emailService.BuildEmailHtmlWithInlineImages(msg, false);
                                string htmlWithHeader = htmlResult.Html;
                                var tempHtmlPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".html");
                                System.IO.File.WriteAllText(tempHtmlPath, htmlWithHeader, System.Text.Encoding.UTF8);
                                var inlineContentIds = emailService.GetInlineContentIds(htmlWithHeader);
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
                                    throw new Exception($"HtmlToPdfWorker failed: {proc.StandardError.ReadToEnd()}");
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
                                        if (att is Storage.Attachment a)
                                        {
                                            string ext2 = System.IO.Path.GetExtension(a.FileName ?? "").ToLowerInvariant();
                                            bool isImage = ext2 == ".jpg" || ext2 == ".jpeg" || ext2 == ".png" || ext2 == ".gif" || ext2 == ".bmp";
                                            if (isImage)
                                            {
                                                if (a.IsInline == true || (!string.IsNullOrEmpty(a.ContentId) && inlineContentIds.Contains(a.ContentId.Trim('<', '>', '"', '\'', ' '))))
                                                    continue;
                                                if (attachmentService.IsLikelySignatureImage(a))
                                                    continue;
                                            }
                                            allAttachments.Add(a);
                                        }
                                        else if (att is Storage.Message nestedMsg)
                                        {
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
                                        Console.WriteLine($"[DEDUP] Found {groupList.Count} files with base name '{group.Key}':");
                                        
                                        foreach (var att in groupList)
                                        {
                                            Console.WriteLine($"[DEDUP]   {att.FileName}");
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
                                            Console.WriteLine($"[DEDUP] Keeping Office file: {keepOfficeFile.FileName}");
                                            
                                            foreach (var pdfFile in pdfFiles)
                                            {
                                                Console.WriteLine($"[DEDUP] Skipping PDF duplicate: {pdfFile.FileName}");
                                                attachmentsToSkip.Add(pdfFile.FileName ?? "");
                                            }
                                            
                                            // If there are multiple Office files, keep only the first one
                                            for (int j = 1; j < officeFiles.Count; j++)
                                            {
                                                Console.WriteLine($"[DEDUP] Skipping duplicate Office file: {officeFiles[j].FileName}");
                                                attachmentsToSkip.Add(officeFiles[j].FileName ?? "");
                                            }
                                        }
                                        else if (officeFiles.Count > 1)
                                        {
                                            // Multiple Office files with same base name - keep first one
                                            var keepOfficeFile = officeFiles.First();
                                            Console.WriteLine($"[DEDUP] Keeping first Office file: {keepOfficeFile.FileName}");
                                            
                                            for (int j = 1; j < officeFiles.Count; j++)
                                            {
                                                Console.WriteLine($"[DEDUP] Skipping duplicate Office file: {officeFiles[j].FileName}");
                                                attachmentsToSkip.Add(officeFiles[j].FileName ?? "");
                                            }
                                        }
                                        else if (pdfFiles.Count > 1)
                                        {
                                            // Multiple PDF files with same base name - keep first one
                                            var keepPdfFile = pdfFiles.First();
                                            Console.WriteLine($"[DEDUP] Keeping first PDF file: {keepPdfFile.FileName}");
                                            
                                            for (int j = 1; j < pdfFiles.Count; j++)
                                            {
                                                Console.WriteLine($"[DEDUP] Skipping duplicate PDF file: {pdfFiles[j].FileName}");
                                                attachmentsToSkip.Add(pdfFiles[j].FileName ?? "");
                                            }
                                        }
                                    }

                                    // Build final list, skipping duplicates
                                    foreach (var a in allAttachments)
                                    {
                                        if (attachmentsToSkip.Contains(a.FileName ?? ""))
                                        {
                                            Console.WriteLine($"[DEDUP] SKIPPING (as planned): {a.FileName}");
                                            continue;
                                        }

                                        Console.WriteLine($"[DEDUP] Including attachment for processing: {a.FileName}");
                                        typedAttachments.Add(a);
                                    }
                                }
                                var allPdfFiles = new List<string> { pdfFilePath };
                                var allTempFiles = new List<string>();
                                string tempDir = System.IO.Path.GetDirectoryName(pdfFilePath);
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
                                        var attachmentParentChain = new List<string> { msg.Subject ?? System.IO.Path.GetFileName(filePath) };
                                        // Pass progressTick to ensure every processed file increments progress
                                        string finalAttachmentPdf = attachmentService.ProcessSingleAttachmentWithHierarchy(att, attPath, tempDir, headerText, allTempFiles, attachmentParentChain, attName, false, () => updateFileProgress?.Invoke(++fileProgress, Math.Max(totalCount, 1)));
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
                                PdfAppendTest.AppendPdfs(allPdfFiles, mergedPdf);
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                // Build list of files to protect: all user-selected source files and all output files
                                var filesToProtect = new HashSet<string>(selectedFiles.Select(f => Path.GetFullPath(f).TrimEnd(Path.DirectorySeparatorChar).ToLowerInvariant()));
                                filesToProtect.Add(Path.GetFullPath(pdfFilePath).TrimEnd(Path.DirectorySeparatorChar).ToLowerInvariant());
                                filesToProtect.Add(Path.GetFullPath(mergedPdf).TrimEnd(Path.DirectorySeparatorChar).ToLowerInvariant());
                                if (generatedPdfs != null)
                                {
                                    foreach (var genPdf in generatedPdfs)
                                    {
                                        filesToProtect.Add(Path.GetFullPath(genPdf).TrimEnd(Path.DirectorySeparatorChar).ToLowerInvariant());
                                    }
                                }
                                Console.WriteLine("[DEBUG] Protected files:");
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
                                if (System.IO.File.Exists(mergedPdf))
                                {
                                    if (System.IO.File.Exists(pdfFilePath))
                                        System.IO.File.Delete(pdfFilePath);
                                    System.IO.File.Move(mergedPdf, pdfFilePath);
                                }
                                generatedPdfs?.Add(pdfFilePath);
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
                                Console.WriteLine($"[DELETE] Should delete: {filePath}, deleteMsgAfterConversion={deleteFilesAfterConversion}");
                                if (deleteFilesAfterConversion && conversionSucceeded && !combineAllPdfs) // <--- updated condition
                                {
                                    if (System.IO.File.Exists(filePath))
                                    {
                                        Console.WriteLine($"[DELETE] Attempting to delete: {filePath}");
                                        try { FileService.MoveFileToRecycleBin(filePath); } catch (Exception ex) { Console.WriteLine($"[DELETE] Failed: {ex.Message}"); }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[DELETE] File not found: {filePath}");
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
                        int totalCount = attachmentService.CountAllProcessableItemsFromFile(filePath);
                        updateFileProgress?.Invoke(0, Math.Max(totalCount, 1));
                        
                        string outputPdf = GenerateUniquePdfFileName(filePath, dir, selectedFiles);
                        var tempFiles = new List<string>();
                        var allPdfFiles = new List<string>();
                        var allTempFiles = new List<string>();
                        string tempDir = dir;
                        string headerText = $"File: {System.IO.Path.GetFileName(filePath)}";
                        var parentChain = new List<string> { System.IO.Path.GetFileName(filePath) };
                        string processedPdf = null;
                        try
                        {
                            processedPdf = attachmentService.ProcessSingleAttachmentWithHierarchy(
                                null, filePath, tempDir, headerText, allTempFiles, parentChain, baseName, extractOriginalOnly,
                                () => updateFileProgress?.Invoke(++fileProgress, Math.Max(totalCount, 1)));
                            if (!string.IsNullOrEmpty(processedPdf))
                                allPdfFiles.Add(processedPdf);
                        }
                        catch (Exception ex)
                        {
                            showMessageBox($"Error processing {System.IO.Path.GetFileName(filePath)}: {ex.Message}");
                            Console.WriteLine($"[ERROR] ProcessSingleAttachmentWithHierarchy failed: {ex}");
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
                        // Remove any zero-page PDFs
                        allPdfFiles.RemoveAll(pdfPath => {
                            try
                            {
                                using (var pdfDoc = PdfSharp.Pdf.IO.PdfReader.Open(pdfPath, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import))
                                {
                                    if (pdfDoc.PageCount == 0)
                                    {
                                        Console.WriteLine($"[ERROR] PDF has zero pages: {pdfPath}");
                                        showMessageBox($"Error: PDF generated from '{System.IO.Path.GetFileName(filePath)}' has no pages and will be skipped.");
                                        return true;
                                    }
                                }
                            }
                            catch (Exception pdfEx)
                            {
                                Console.WriteLine($"[ERROR] Failed to open PDF for page count check: {pdfEx.Message}");
                                return true;
                            }
                            return false;
                        });
                        // Merge if more than one PDF, else just move/rename
                        string finalPdf = outputPdf;
                        if (allPdfFiles.Count > 1)
                        {
                            string mergedPdf = System.IO.Path.Combine(tempDir, baseName + "_merged.pdf");
                            PdfAppendTest.AppendPdfs(allPdfFiles, mergedPdf);
                            finalPdf = mergedPdf;
                        }
                        else if (allPdfFiles.Count == 1)
                        {
                            finalPdf = allPdfFiles[0];
                        }
                        Console.WriteLine($"[DEBUG] finalPdf: {finalPdf}, outputPdf: {outputPdf}");
                        if (!string.Equals(finalPdf, outputPdf, StringComparison.OrdinalIgnoreCase) && System.IO.File.Exists(finalPdf))
                        {
                            try
                            {
                                if (System.IO.File.Exists(outputPdf))
                                    System.IO.File.Delete(outputPdf);
                                if (deleteFilesAfterConversion)
                                {
                                    System.IO.File.Move(finalPdf, outputPdf);
                                    Console.WriteLine($"[DEBUG] Moved {finalPdf} to {outputPdf}");
                                }
                                else
                                {
                                    System.IO.File.Copy(finalPdf, outputPdf, true);
                                    Console.WriteLine($"[DEBUG] Copied {finalPdf} to {outputPdf}");
                                }
                            }
                            catch (Exception moveEx)
                            {
                                Console.WriteLine($"[ERROR] Failed to move/copy file: {moveEx.Message}");
                                outputPdf = finalPdf;
                            }
                        }
                        else if (System.IO.File.Exists(outputPdf))
                        {
                            Console.WriteLine($"[DEBUG] Output PDF already exists at {outputPdf}");
                        }
                        else if (System.IO.File.Exists(finalPdf))
                        {
                            // If outputPdf does not exist but finalPdf does, treat as output
                            outputPdf = finalPdf;
                        }
                        Console.WriteLine($"[DEBUG] Checking existence of output PDF: {outputPdf}");
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
                            Console.WriteLine($"[ERROR] No PDF generated for {System.IO.Path.GetFileName(filePath)}");
                            fail++;
                        }
                        // Cleanup temp files, but never delete the output PDF
                        foreach (var tempFile in allTempFiles)
                        {
                            try {
                                if (System.IO.File.Exists(tempFile) && !string.Equals(tempFile, outputPdf, StringComparison.OrdinalIgnoreCase))
                                {
                                    System.IO.File.Delete(tempFile);
                                    Console.WriteLine($"[DEBUG] Deleted temp file: {tempFile}");
                                }
                            } catch { }
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
                            Console.WriteLine($"[DELETE] Skipping deletion of source PDF in output folder: {filePath}");
                        }
                        else if (System.IO.File.Exists(filePath))
                        {
                            Console.WriteLine($"[DELETE] Attempting to delete: {filePath}");
                            try { FileService.MoveFileToRecycleBin(filePath); } catch (Exception ex) { Console.WriteLine($"[DELETE] Failed: {ex.Message}"); }
                        }
                        else
                        {
                            Console.WriteLine($"[DELETE] File not found: {filePath}");
                        }
                    }
                }
            }
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
                
                // Generate unique PDF filename to avoid conflicts when files have the same base name but different extensions
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
                Console.WriteLine($"Error in ConvertSingleMsgFileWithProgress: {ex.Message}");
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
