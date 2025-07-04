using MsgReader.Outlook;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Collections.Generic;
using System.IO;
using System;
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
                string pdfFileName = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(msgFilePath) + ".pdf");
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
            EmailConverterService emailService,
            AttachmentService attachmentService,
            Action<int, int, int, string> updateProgress, // (processed, total, progress, statusText)
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
                            // Use robust attachment/merge logic for .msg files
                            Storage.Message msg = null;
                            try
                            {
                                msg = new Storage.Message(filePath);
                                string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
                                string msgBaseName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                                string msgDir = !string.IsNullOrEmpty(selectedOutputFolder) ? selectedOutputFolder : System.IO.Path.GetDirectoryName(filePath);
                                string pdfFilePath = System.IO.Path.Combine(msgDir, $"{msgBaseName} - {datePart}.pdf");
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
                                            typedAttachments.Add(a);
                                        }
                                        else if (att is Storage.Message nestedMsg)
                                        {
                                            nestedMessages.Add(nestedMsg);
                                        }
                                    }
                                }
                                var allPdfFiles = new List<string> { pdfFilePath };
                                var allTempFiles = new List<string>();
                                string tempDir = System.IO.Path.GetDirectoryName(pdfFilePath);
                                int totalAttachments = typedAttachments.Count;
                                for (int attIndex = 0; attIndex < typedAttachments.Count; attIndex++)
                                {
                                    var att = typedAttachments[attIndex];
                                    string attName = att.FileName ?? "attachment";
                                    string attPath = System.IO.Path.Combine(tempDir, attName);
                                    string headerText = $"Attachment {attIndex + 1}/{totalAttachments} - {attName}";
                                    try
                                    {
                                        System.IO.File.WriteAllBytes(attPath, att.Data);
                                        allTempFiles.Add(attPath);
                                        var attachmentParentChain = new List<string> { msg.Subject ?? System.IO.Path.GetFileName(filePath) };
                                        string finalAttachmentPdf = attachmentService.ProcessSingleAttachmentWithHierarchy(att, attPath, tempDir, headerText, allTempFiles, attachmentParentChain, attName, false);
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
                                    attachmentService.ProcessMsgAttachmentsRecursively(nestedMsg, allPdfFiles, allTempFiles, tempDir, false, 1, 5, nestedHeaderText, initialParentChain);
                                }
                                string mergedPdf = System.IO.Path.Combine(tempDir, System.IO.Path.GetFileNameWithoutExtension(pdfFilePath) + "_merged.pdf");
                                PdfAppendTest.AppendPdfs(allPdfFiles, mergedPdf);
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                foreach (var f in allTempFiles)
                                {
                                    try
                                    {
                                        if (System.IO.File.Exists(f) && !string.Equals(f, mergedPdf, StringComparison.OrdinalIgnoreCase) && !string.Equals(f, pdfFilePath, StringComparison.OrdinalIgnoreCase))
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
                                if (deleteFilesAfterConversion && conversionSucceeded)
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
                            var result = ConvertSingleMsgFile(filePath, dir, appendAttachments, extractOriginalOnly, emailService, attachmentService, generatedPdfs);
                            if (result) { success++; conversionSucceeded = true; } else fail++;
                        }
                    }
                    else
                    {
                        // Use the same hierarchical logic as for attachments
                        string outputPdf = System.IO.Path.Combine(dir, $"{baseName}.pdf");
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
                                null, filePath, tempDir, headerText, allTempFiles, parentChain, baseName, extractOriginalOnly);
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
                                System.IO.File.Move(finalPdf, outputPdf);
                                Console.WriteLine($"[DEBUG] Moved {finalPdf} to {outputPdf}");
                            }
                            catch (Exception moveEx)
                            {
                                Console.WriteLine($"[ERROR] Failed to move file: {moveEx.Message}");
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
                    if (deleteFilesAfterConversion && conversionSucceeded)
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
            return (success, fail, processed, isCancellationRequested());
        }

        private bool ConvertSingleMsgFile(string msgFilePath, string outputDir, bool appendAttachments, bool extractOriginalOnly, EmailConverterService emailService, AttachmentService attachmentService, List<string> generatedPdfs)
        {
            Storage.Message msg = new Storage.Message(msgFilePath);
            string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
            string baseName = System.IO.Path.GetFileNameWithoutExtension(msgFilePath);
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
    }
}
