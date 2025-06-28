using System;
using System.Collections.Generic;
using System.IO;
using MsgReader.Outlook;
using System.Diagnostics;

namespace MsgToPdfConverter.Services
{
    public class AttachmentService
    {
        private readonly Action<string, string, string> _addHeaderPdf;
        private readonly Func<string, string, bool> _tryConvertOfficeToPdf;
        private readonly Action<List<string>, string> _appendPdfs;
        private readonly EmailConverterService _emailService;

        public AttachmentService(Action<string, string, string> addHeaderPdf, Func<string, string, bool> tryConvertOfficeToPdf, Action<List<string>, string> appendPdfs, EmailConverterService emailService)
        {
            _addHeaderPdf = addHeaderPdf;
            _tryConvertOfficeToPdf = tryConvertOfficeToPdf;
            _appendPdfs = appendPdfs;
            _emailService = emailService;
        }

        public void ProcessMsgAttachmentsRecursively(Storage.Message msg, List<string> allPdfFiles, List<string> allTempFiles, string tempDir, bool extractOriginalOnly, int depth = 0, int maxDepth = 5, string headerText = null)
        {
            // Prevent infinite recursion
            if (depth > maxDepth)
            {
                Console.WriteLine($"[MSG] Stopping recursion at depth {depth} - max depth reached");
                return;
            }

            Console.WriteLine($"[MSG] Processing message recursively at depth {depth}: {msg.Subject ?? "No Subject"}");

            // First, create PDF for the nested MSG body content (if we're processing a nested message)
            if (depth > 0)
            {
                string msgSubject = msg.Subject ?? $"nested_msg_depth_{depth}";
                string usedHeaderText = headerText ?? $"Nested Email (Depth {depth}): {msgSubject}";
                try
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Creating PDF for nested message body: {msgSubject}");
                    // Use new inline image logic for nested emails
                    var htmlResult = _emailService.BuildEmailHtmlWithInlineImages(msg, extractOriginalOnly);
                    string nestedHtml = htmlResult.Html;
                    List<string> tempFiles = htmlResult.TempFiles;
                    allTempFiles.AddRange(tempFiles); // Track temp image files for cleanup
                    string nestedPdf = Path.Combine(tempDir, $"depth{depth}_{Guid.NewGuid()}_nested_msg.pdf");
                    string nestedHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".html");
                    File.WriteAllText(nestedHtmlPath, nestedHtml, System.Text.Encoding.UTF8);
                    var startInfo = new ProcessStartInfo
                    {
                        FileName = System.Reflection.Assembly.GetExecutingAssembly().Location,
                        Arguments = $"--html2pdf \"{nestedHtmlPath}\" \"{nestedPdf}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    using (var process = Process.Start(startInfo))
                    {
                        process.WaitForExit();
                        if (process.ExitCode == 0)
                        {
                            // Always use the provided headerText for the header
                            string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                            _addHeaderPdf(headerPdf, usedHeaderText, null);
                            string finalNestedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_nested_merged.pdf");
                            _appendPdfs(new List<string> { headerPdf, nestedPdf }, finalNestedPdf);
                            allPdfFiles.Add(finalNestedPdf);
                            allTempFiles.Add(headerPdf);
                            allTempFiles.Add(nestedPdf);
                            allTempFiles.Add(finalNestedPdf);
                            Console.WriteLine($"[MSG] Depth {depth} - Successfully created PDF for nested MSG body: {msgSubject}");
                        }
                        else
                        {
                            Console.WriteLine($"[MSG] Depth {depth} - Failed to convert nested MSG body to PDF: {msgSubject}");
                        }
                    }
                    File.Delete(nestedHtmlPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Error processing nested MSG body {msgSubject}: {ex.Message}");
                    string errorPdf = Path.Combine(tempDir, Guid.NewGuid() + "_msg_error.pdf");
                    _addHeaderPdf(errorPdf, usedHeaderText + $"\n(Error: {ex.Message})", null);
                    allPdfFiles.Add(errorPdf);
                    allTempFiles.Add(errorPdf);
                }
            }

            // Now process attachments if they exist
            if (msg.Attachments == null || msg.Attachments.Count == 0)
            {
                Console.WriteLine($"[MSG] Depth {depth} - No attachments to process");
                return;
            }

            Console.WriteLine($"[MSG] Processing attachments at depth {depth}, found {msg.Attachments.Count} attachments");

            var inlineContentIds = _emailService.GetInlineContentIds(msg.BodyHtml ?? "");
            var typedAttachments = new List<Storage.Attachment>();
            var nestedMessages = new List<Storage.Message>();

            // Separate attachments and nested messages
            foreach (var att in msg.Attachments)
            {
                if (att is Storage.Attachment a)
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Examining attachment: {a.FileName} (IsInline: {a.IsInline}, ContentId: {a.ContentId})");

                    if ((a.IsInline == true) || (!string.IsNullOrEmpty(a.ContentId) && inlineContentIds.Contains(a.ContentId.Trim('<', '>', '\"', '\'', ' '))))
                    {
                        Console.WriteLine($"[MSG] Depth {depth} - Skipping inline attachment: {a.FileName}");
                        continue;
                    }

                    typedAttachments.Add(a);
                }
                else if (att is Storage.Message nestedMsg)
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Found nested MSG: {nestedMsg.Subject ?? "No Subject"}");
                    nestedMessages.Add(nestedMsg);
                }
            }

            // Process regular attachments
            int totalAttachments = typedAttachments.Count;
            for (int attIndex = 0; attIndex < typedAttachments.Count; attIndex++)
            {
                var att = typedAttachments[attIndex];
                string attName = att.FileName ?? "attachment";
                string attPath = Path.Combine(tempDir, $"depth{depth}_{attName}");
                string attachmentHeaderText = $"Attachment (Depth {depth}): {attIndex + 1}/{totalAttachments} - {attName}";

                try
                {
                    File.WriteAllBytes(attPath, att.Data);
                    allTempFiles.Add(attPath);
                    string finalAttachmentPdf = ProcessSingleAttachment(att, attPath, tempDir, attachmentHeaderText, allTempFiles);

                    if (finalAttachmentPdf != null)
                        allPdfFiles.Add(finalAttachmentPdf);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Error processing attachment {attName}: {ex.Message}");
                    string errorPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                    _addHeaderPdf(errorPdf, attachmentHeaderText + $"\n(Error: {ex.Message})", null);
                    allPdfFiles.Add(errorPdf);
                    allTempFiles.Add(errorPdf);
                }
            }

            // Process nested MSG files recursively (this will handle both their body content and attachments)
            for (int msgIndex = 0; msgIndex < nestedMessages.Count; msgIndex++)
            {
                var nestedMsg = nestedMessages[msgIndex];
                string nestedSubject = nestedMsg.Subject ?? $"nested_msg_depth_{depth + 1}";
                string nestedHeaderText = $"Attachment (Depth {depth + 1}): {msgIndex + 1}/{nestedMessages.Count} - Nested Email: {nestedSubject}";
                Console.WriteLine($"[MSG] Depth {depth} - Recursively processing nested message {msgIndex + 1}/{nestedMessages.Count}: {nestedSubject}");
                ProcessMsgAttachmentsRecursively(nestedMsg, allPdfFiles, allTempFiles, tempDir, extractOriginalOnly, depth + 1, maxDepth, nestedHeaderText);
            }
        }

        public string ProcessSingleAttachment(Storage.Attachment att, string attPath, string tempDir, string headerText, List<string> allTempFiles)
        {
            string attName = att.FileName ?? "attachment";
            string ext = Path.GetExtension(attName).ToLowerInvariant();
            string attPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName) + ".pdf");
            string finalAttachmentPdf = null;

            try
            {
                if (ext == ".pdf")
                {
                    string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                    _addHeaderPdf(headerPdf, headerText, null);
                    finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                    _appendPdfs(new List<string> { headerPdf, attPath }, finalAttachmentPdf);
                    allTempFiles.Add(headerPdf);
                    allTempFiles.Add(finalAttachmentPdf);
                }
                else if (ext == ".jpg" || ext == ".jpeg")
                {
                    using (var writer = new iText.Kernel.Pdf.PdfWriter(attPdf))
                    using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                    using (var docImg = new iText.Layout.Document(pdf))
                    {
                        var p = new iText.Layout.Element.Paragraph(headerText)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetFontSize(16);
                        docImg.Add(p);
                        var imgData = iText.IO.Image.ImageDataFactory.Create(attPath);
                        var image = new iText.Layout.Element.Image(imgData);
                        docImg.Add(image);
                    }
                    finalAttachmentPdf = attPdf;
                    allTempFiles.Add(attPdf);
                }
                else if (ext == ".doc" || ext == ".docx" || ext == ".xls" || ext == ".xlsx")
                {
                    if (_tryConvertOfficeToPdf(attPath, attPdf))
                    {
                        string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                        _addHeaderPdf(headerPdf, headerText, null);
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                        _appendPdfs(new List<string> { headerPdf, attPdf }, finalAttachmentPdf);
                        allTempFiles.Add(headerPdf);
                        allTempFiles.Add(attPdf);
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                        _addHeaderPdf(finalAttachmentPdf, headerText + "\n(Conversion failed)", null);
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else if (ext == ".zip")
                {
                    finalAttachmentPdf = ProcessZipAttachment(attPath, tempDir, headerText, allTempFiles);
                }
                else
                {
                    finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                    _addHeaderPdf(finalAttachmentPdf, headerText + "\n(Unsupported type)", null);
                    allTempFiles.Add(finalAttachmentPdf);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ATTACH] Error processing attachment {attName}: {ex.Message}");
                finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                _addHeaderPdf(finalAttachmentPdf, headerText + $"\n(Error: {ex.Message})", null);
                allTempFiles.Add(finalAttachmentPdf);
            }

            return finalAttachmentPdf;
        }

        public string ProcessZipAttachment(string attPath, string tempDir, string headerText, List<string> allTempFiles)
        {
            string extractDir = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attPath));
            System.IO.Compression.ZipFile.ExtractToDirectory(attPath, extractDir);
            allTempFiles.Add(extractDir);

            var zipFiles = Directory.GetFiles(extractDir, "*.*", System.IO.SearchOption.AllDirectories);
            var zipPdfFiles = new List<string>();
            int zipFileIndex = 0;

            foreach (var zf in zipFiles)
            {
                zipFileIndex++;
                string zfPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(zf) + ".pdf");
                string zfExt = Path.GetExtension(zf).ToLowerInvariant();
                string zipHeader = $"{headerText} (ZIP {zipFileIndex}/{zipFiles.Length})";
                string finalZipPdf = null;

                if (zfExt == ".pdf")
                {
                    string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                    _addHeaderPdf(headerPdf, zipHeader, null);
                    finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                    _appendPdfs(new List<string> { headerPdf, zf }, finalZipPdf);
                    allTempFiles.Add(headerPdf);
                    allTempFiles.Add(finalZipPdf);
                }
                else if (zfExt == ".doc" || zfExt == ".docx" || zfExt == ".xls" || zfExt == ".xlsx")
                {
                    if (_tryConvertOfficeToPdf(zf, zfPdf))
                    {
                        string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                        _addHeaderPdf(headerPdf, zipHeader, null);
                        finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                        _appendPdfs(new List<string> { headerPdf, zfPdf }, finalZipPdf);
                        allTempFiles.Add(headerPdf);
                        allTempFiles.Add(zfPdf);
                        allTempFiles.Add(finalZipPdf);
                    }
                    else
                    {
                        _addHeaderPdf(finalZipPdf, zipHeader + "\n(Conversion failed)", null);
                        allTempFiles.Add(finalZipPdf);
                    }
                }
                else
                {
                    finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                    _addHeaderPdf(finalZipPdf, zipHeader + "\n(Unsupported type)", null);
                    allTempFiles.Add(finalZipPdf);
                }
                if (finalZipPdf != null)
                    zipPdfFiles.Add(finalZipPdf);
            }

            // Merge all files from the ZIP into a single PDF
            if (zipPdfFiles.Count == 0)
            {
                // No processable files found in ZIP
                string placeholderPdf = Path.Combine(tempDir, Guid.NewGuid() + "_empty_zip.pdf");
                _addHeaderPdf(placeholderPdf, headerText + "\n(Empty or no processable files in ZIP)", null);
                allTempFiles.Add(placeholderPdf);
                return placeholderPdf;
            }
            else if (zipPdfFiles.Count == 1)
            {
                // Only one file, return it directly
                return zipPdfFiles[0];
            }
            else
            {
                // Multiple files - merge them all into one comprehensive PDF
                string mergedZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_merged.pdf");
                Console.WriteLine($"[ZIP] Merging {zipPdfFiles.Count} files from ZIP into single PDF: {mergedZipPdf}");
                _appendPdfs(zipPdfFiles, mergedZipPdf);
                allTempFiles.Add(mergedZipPdf);
                return mergedZipPdf;
            }
        }
    }
}
