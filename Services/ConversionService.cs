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

        public (int Success, int Fail, int Processed, bool Cancelled) ConvertMsgFilesWithAttachments(
            List<string> selectedFiles,
            string selectedOutputFolder,
            bool appendAttachments,
            bool extractOriginalOnly,
            bool deleteMsgAfterConversion,
            EmailConverterService emailService,
            AttachmentService attachmentService,
            Action<int, int, int, string> updateProgress, // (processed, total, progress, statusText)
            Func<bool> isCancellationRequested,
            Action<string> showMessageBox // (message)
        )
        {
            int success = 0, fail = 0, processed = 0;
            for (int i = 0; i < selectedFiles.Count; i++)
            {
                if (isCancellationRequested())
                    break;
                processed++;
                string msgFilePath = selectedFiles[i];
                Storage.Message msg = null;
                try
                {
                    updateProgress(processed, selectedFiles.Count, i, $"Processing file {processed}/{selectedFiles.Count}: {System.IO.Path.GetFileName(selectedFiles[i])}");
                    msg = new Storage.Message(msgFilePath);
                    string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
                    string baseName = System.IO.Path.GetFileNameWithoutExtension(msgFilePath);
                    string dir = !string.IsNullOrEmpty(selectedOutputFolder) ? selectedOutputFolder : System.IO.Path.GetDirectoryName(msgFilePath);
                    string pdfFilePath = System.IO.Path.Combine(dir, $"{baseName} - {datePart}.pdf");
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
                    if (appendAttachments && msg.Attachments != null && msg.Attachments.Count > 0)
                    {
                        Console.WriteLine($"[DEBUG] Found {msg.Attachments.Count} total attachments in {baseName}");
                        var typedAttachments = new List<Storage.Attachment>();
                        var nestedMessages = new List<Storage.Message>();
                        foreach (var att in msg.Attachments)
                        {
                            if (att is Storage.Attachment a)
                            {
                                Console.WriteLine($"[DEBUG] Attachment: {a.FileName}, IsInline: {a.IsInline}, ContentId: '{a.ContentId}', Size: {a.Data?.Length ?? 0} bytes");
                                string ext = System.IO.Path.GetExtension(a.FileName ?? "").ToLowerInvariant();
                                bool isImage = ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".gif" || ext == ".bmp";
                                // Only skip inline or signature images, not other file types
                                if (isImage)
                                {
                                    if (a.IsInline == true || (!string.IsNullOrEmpty(a.ContentId) && inlineContentIds.Contains(a.ContentId.Trim('<', '>', '"', '\'', ' '))))
                                    {
                                        Console.WriteLine($"[DEBUG] Skipping inline attachment (IsInline: {a.IsInline} or referenced in email body): {a.FileName}");
                                        continue;
                                    }
                                    if (IsLikelySignatureImage(a))
                                    {
                                        Console.WriteLine($"[DEBUG] Skipping likely signature/decorative image: {a.FileName}");
                                        continue;
                                    }
                                }
                                Console.WriteLine($"[DEBUG] Adding attachment for processing: {a.FileName}");
                                typedAttachments.Add(a);
                            }
                            else if (att is Storage.Message nestedMsg)
                            {
                                Console.WriteLine($"[DEBUG] Found nested message: {nestedMsg.Subject ?? "[No Subject]"}");
                                nestedMessages.Add(nestedMsg);
                            }
                        }
                        Console.WriteLine($"[DEBUG] After filtering: {typedAttachments.Count} regular attachments, {nestedMessages.Count} nested messages");
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
                                var attachmentParentChain = new List<string> { msg.Subject ?? System.IO.Path.GetFileName(msgFilePath) };
                                string finalAttachmentPdf = attachmentService.ProcessSingleAttachmentWithHierarchy(att, attPath, tempDir, headerText, allTempFiles, attachmentParentChain, attName, false);
                                if (finalAttachmentPdf != null)
                                    allPdfFiles.Add(finalAttachmentPdf);
                            }
                            catch (Exception ex)
                            {
                                string errorPdf = System.IO.Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                                var errorParentChain = new List<string> { msg.Subject ?? System.IO.Path.GetFileName(msgFilePath) };
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
                            var initialParentChain = new List<string> { msg.Subject ?? System.IO.Path.GetFileName(msgFilePath) };
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
                    }
                    success++;
                }
                catch
                {
                    fail++;
                }
                finally
                {
                    if (msg != null && msg is IDisposable disposableMsg)
                        disposableMsg.Dispose();
                    msg = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    Console.WriteLine($"[DELETE] Should delete: {msgFilePath}, deleteMsgAfterConversion={deleteMsgAfterConversion}");
                    if (deleteMsgAfterConversion)
                    {
                        if (System.IO.File.Exists(msgFilePath))
                        {
                            Console.WriteLine($"[DELETE] Attempting to delete: {msgFilePath}");
                            try { FileService.MoveFileToRecycleBin(msgFilePath); } catch (Exception ex) { Console.WriteLine($"[DELETE] Failed: {ex.Message}"); }
                        }
                        else
                        {
                            Console.WriteLine($"[DELETE] File not found: {msgFilePath}");
                        }
                    }
                }
            }
            return (success, fail, processed, isCancellationRequested());
        }

        /// <summary>
        /// Determines if an attachment is likely a signature image or decorative element that should be skipped
        /// </summary>
        private bool IsLikelySignatureImage(Storage.Attachment attachment)
        {
            try
            {
                string fileName = attachment.FileName ?? "";
                string ext = Path.GetExtension(fileName).ToLowerInvariant();

                // Only check image files
                if (ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" && ext != ".bmp")
                {
                    return false; // Not an image, so not a signature image
                }

                // Check file size - signature images are typically small (less than 50KB)
                int fileSizeKB = (attachment.Data?.Length ?? 0) / 1024;
                bool isSmallImage = fileSizeKB < 50;

                // Check for common signature image patterns in filename
                string lowerFileName = fileName.ToLowerInvariant();
                bool hasSignaturePattern = lowerFileName.Contains("image") ||
                                         lowerFileName.Contains("signature") ||
                                         lowerFileName.Contains("logo") ||
                                         lowerFileName.Contains("banner") ||
                                         lowerFileName.StartsWith("oledata.mso");

                // If it's a small image with signature patterns, likely a signature
                if (isSmallImage && hasSignaturePattern)
                {
                    Console.WriteLine($"[FILTER] Detected signature image: {fileName} ({fileSizeKB}KB)");
                    return true;
                }

                // If it's marked as inline AND small, likely decorative/signature
                if (attachment.IsInline == true && isSmallImage)
                {
                    Console.WriteLine($"[FILTER] Detected small inline image: {fileName} ({fileSizeKB}KB)");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[FILTER] Error checking signature image {attachment.FileName}: {ex.Message}");
                return false; // If in doubt, don't filter out
            }
        }
    }
}