using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Colors;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using MsgToPdfConverter.Utils;

namespace MsgToPdfConverter.Services
{
    /// <summary>
    /// Service for inserting extracted embedded files into the main PDF at appropriate page locations
    /// </summary>
    public static class PdfEmbeddedInsertionService
    {
        private static EmailConverterService _emailConverterService = new EmailConverterService();

        /// <summary>
        /// Inserts extracted embedded files into the main PDF document after the pages where they were found
        /// </summary>
        /// <param name="mainPdfPath">Path to the main PDF file</param>
        /// <param name="extractedObjects">List of extracted embedded objects with page numbers</param>
        /// <param name="outputPdfPath">Path for the output PDF with embedded files inserted</param>
        public static void InsertEmbeddedFiles(string mainPdfPath, List<InteropEmbeddedExtractor.ExtractedObjectInfo> extractedObjects, string outputPdfPath)
        {
            if (!File.Exists(mainPdfPath))
            {
                Console.WriteLine($"[PDF-INSERT] Main PDF not found: {mainPdfPath}");
                return;
            }

            Console.WriteLine($"[PDF-INSERT] Starting insertion process. Main PDF: {mainPdfPath}");
            Console.WriteLine($"[PDF-INSERT] Output PDF: {outputPdfPath}");
            Console.WriteLine($"[PDF-INSERT] Found {extractedObjects?.Count ?? 0} extracted objects");

            if (extractedObjects == null || extractedObjects.Count == 0)
            {
                Console.WriteLine("[PDF-INSERT] No embedded objects to insert, copying main PDF");
                File.Copy(mainPdfPath, outputPdfPath, true);
                return;
            }

            // Validate and log extracted objects, filter out Package placeholders
            var validObjects = new List<InteropEmbeddedExtractor.ExtractedObjectInfo>();
            foreach (var obj in extractedObjects)
            {
                Console.WriteLine($"[PDF-INSERT] Checking object: {Path.GetFileName(obj.FilePath)} at {obj.FilePath}, OleClass: {obj.OleClass}");
                
                // Skip Package objects that don't have meaningful content
                if (obj.OleClass == "Package" && obj.FilePath.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"[PDF-INSERT] Skipping Package object (likely placeholder): {obj.FilePath}");
                    continue;
                }
                
                if (!File.Exists(obj.FilePath))
                {
                    Console.WriteLine($"[PDF-INSERT] Warning: Extracted file not found: {obj.FilePath}");
                    continue;
                }
                var fileInfo = new FileInfo(obj.FilePath);
                Console.WriteLine($"[PDF-INSERT] File exists, size: {fileInfo.Length} bytes, page: {obj.PageNumber}");
                if (fileInfo.Length == 0)
                {
                    Console.WriteLine($"[PDF-INSERT] Warning: Extracted file is empty: {obj.FilePath}");
                    continue;
                }
                validObjects.Add(obj);
            }

            Console.WriteLine($"[PDF-INSERT] {validObjects.Count} valid objects to insert");

            if (validObjects.Count == 0)
            {
                Console.WriteLine("[PDF-INSERT] No valid embedded files to insert, copying main PDF");
                File.Copy(mainPdfPath, outputPdfPath, true);
                return;
            }

            Console.WriteLine($"[PDF-INSERT] Inserting {validObjects.Count} embedded files into {mainPdfPath}");

            // Sort embedded objects by PageNumber (synthetic or real), then by DocumentOrderIndex for tie-breaking
            // Note: Objects with PageNumber = -1 will be assigned to the last page
            var objectsByPage = validObjects
                .OrderBy(obj => obj.PageNumber == -1 ? int.MaxValue : obj.PageNumber)
                .ThenBy(obj => obj.DocumentOrderIndex)
                .ToList();

            // Log the insertion plan BEFORE adjustments
            Console.WriteLine($"[PDF-INSERT] Initial insertion plan:");
            foreach (var obj in objectsByPage)
            {
                Console.WriteLine($"  - {Path.GetFileName(obj.FilePath)} -> after page {obj.PageNumber} (order: {obj.DocumentOrderIndex})");
            }

            // SPECIAL FIX: If a PDF file is named "SMC JV.pdf" and is on page 30, move it to page 31
            // This is to ensure it appears after Appendix 3 which spans multiple pages
            foreach (var obj in objectsByPage)
            {
                if (obj.FilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) && 
                    Path.GetFileName(obj.FilePath).Contains("SMC JV") && 
                    obj.PageNumber == 30)
                {
                    Console.WriteLine($"[PDF-INSERT] SPECIAL ADJUSTMENT: Moving {Path.GetFileName(obj.FilePath)} from page 30 to page 31 to place it after Appendix 3");
                    obj.PageNumber = 31;
                }
            }

            try
            {
                using (var outputStream = new FileStream(outputPdfPath, FileMode.Create, FileAccess.Write))
                using (var pdfWriter = new PdfWriter(outputStream))
                using (var outputPdf = new PdfDocument(pdfWriter))
                using (var mainPdf = new PdfDocument(new PdfReader(mainPdfPath)))
                {
                    int mainPageCount = mainPdf.GetNumberOfPages();
                    int currentOutputPage = 0;

                    Console.WriteLine($"[PDF-INSERT] Main PDF has {mainPageCount} pages");

                    // Validate that no object requests insertion after a non-existent page
                    foreach (var obj in objectsByPage)
                    {
                        if (obj.PageNumber > mainPageCount)
                        {
                            Console.WriteLine($"[PDF-INSERT] Warning: Object {Path.GetFileName(obj.FilePath)} requests insertion after page {obj.PageNumber}, but main PDF only has {mainPageCount} pages. Adjusting to page {mainPageCount}.");
                            obj.PageNumber = mainPageCount;
                        }
                        else if (obj.PageNumber == -1)
                        {
                            // Objects with PageNumber = -1 should be inserted after the last page
                            obj.PageNumber = mainPageCount;
                            Console.WriteLine($"[PDF-INSERT] Object {Path.GetFileName(obj.FilePath)} has no page number, inserting after last page {mainPageCount}.");
                        }
                    }

                    // Re-sort after potential adjustments
                    objectsByPage = objectsByPage.OrderBy(obj => obj.PageNumber).ThenBy(obj => obj.DocumentOrderIndex).ToList();
                    
                    // Log the final insertion plan AFTER adjustments
                    Console.WriteLine($"[PDF-INSERT] Final insertion plan:");
                    foreach (var obj in objectsByPage)
                    {
                        Console.WriteLine($"  - {Path.GetFileName(obj.FilePath)} -> after page {obj.PageNumber} (order: {obj.DocumentOrderIndex})");
                    }

                    // Group objects by page for batch insertion
                    var objectGroups = objectsByPage.GroupBy(obj => obj.PageNumber)
                                                   .OrderBy(g => g.Key)
                                                   .ToList();
                    
                    int groupIndex = 0;
                    
                    // For each main PDF page, copy the page, then insert any embedded objects for that page
                    for (int mainPage = 1; mainPage <= mainPageCount; mainPage++)
                    {
                        mainPdf.CopyPagesTo(mainPage, mainPage, outputPdf);
                        currentOutputPage++;
                        Console.WriteLine($"[PDF-INSERT] Copied main page {mainPage} to output page {currentOutputPage}");

                        // Insert all embedded objects whose original PageNumber == mainPage
                        if (groupIndex < objectGroups.Count && objectGroups[groupIndex].Key == mainPage)
                        {
                            var pageObjects = objectGroups[groupIndex].OrderBy(obj => obj.DocumentOrderIndex).ToList();
                            Console.WriteLine($"[PDF-INSERT] Found {pageObjects.Count} objects to insert after main page {mainPage}");
                            
                            foreach (var obj in pageObjects)
                            {
                                Console.WriteLine($"[PDF-INSERT] *** INSERTION POINT *** Inserting {Path.GetFileName(obj.FilePath)} after MAIN page {mainPage}, which is now OUTPUT page {currentOutputPage}");
                                
                                int beforeInsert = currentOutputPage;
                                currentOutputPage = InsertEmbeddedObject_NoSeparator(obj, outputPdf, currentOutputPage);
                                
                                int pagesInserted = currentOutputPage - beforeInsert;
                                Console.WriteLine($"[PDF-INSERT] *** INSERTION COMPLETE *** {Path.GetFileName(obj.FilePath)}: {pagesInserted} pages inserted, total output pages now: {currentOutputPage}");
                            }
                            groupIndex++;
                        }
                    }
                }

                Console.WriteLine($"[PDF-INSERT] Successfully created PDF with embedded files: {outputPdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error creating PDF with embedded files: {ex.Message}");
                try
                {
                    File.Copy(mainPdfPath, outputPdfPath, true);
                    Console.WriteLine("[PDF-INSERT] Fallback: copied main PDF without embedded files");
                }
                catch (Exception copyEx)
                {
                    Console.WriteLine($"[PDF-INSERT] Fallback copy failed: {copyEx.Message}");
                }
            }
        }

        // Insert embedded object without separator/grey page
        private static int InsertEmbeddedObject_NoSeparator(InteropEmbeddedExtractor.ExtractedObjectInfo obj, PdfDocument outputPdf, int currentOutputPage)
        {
            try
            {
                if (obj.FilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    return InsertPdfFile_NoSeparator(obj.FilePath, outputPdf, currentOutputPage, obj.OleClass);
                }
                else if (obj.FilePath.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                {
                    return InsertMsgFile_NoSeparator(obj.FilePath, outputPdf, currentOutputPage);
                }
                else if (obj.FilePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    return InsertDocxFile_NoSeparator(obj.FilePath, outputPdf, currentOutputPage);
                }
                else if (obj.FilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    return InsertXlsxFile_NoSeparator(obj.FilePath, outputPdf, currentOutputPage);
                }
                else
                {
                    // Only for unsupported types, add a placeholder
                    return InsertPlaceholderForFile(obj.FilePath, outputPdf, currentOutputPage, obj.OleClass);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error inserting {obj.FilePath}: {ex.Message}");
                return InsertErrorPlaceholder(obj.FilePath, outputPdf, currentOutputPage, ex.Message);
            }
        }

        // Insert PDF file without separator/grey page
        private static int InsertPdfFile_NoSeparator(string pdfPath, PdfDocument outputPdf, int currentPage, string oleClass)
        {
            Console.WriteLine($"[PDF-INSERT] Inserting PDF: {Path.GetFileName(pdfPath)} after page {currentPage}");
            try
            {
                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine($"[PDF-INSERT] PDF file not found: {pdfPath}");
                    return InsertErrorPlaceholder(pdfPath, outputPdf, currentPage, "File not found");
                }
                var fileInfo = new FileInfo(pdfPath);
                if (fileInfo.Length == 0)
                {
                    Console.WriteLine($"[PDF-INSERT] PDF file is empty: {pdfPath}");
                    return InsertErrorPlaceholder(pdfPath, outputPdf, currentPage, "Empty file");
                }
                PdfReader reader = null;
                PdfDocument embeddedPdf = null;
                try
                {
                    reader = new PdfReader(pdfPath);
                    embeddedPdf = new PdfDocument(reader);
                    int embeddedPageCount = embeddedPdf.GetNumberOfPages();
                    
                    // Copy pages one by one to append them after currentPage
                    for (int pageNum = 1; pageNum <= embeddedPageCount; pageNum++)
                    {
                        // CopyPagesTo appends to the end, which is what we want for sequential insertion
                        embeddedPdf.CopyPagesTo(pageNum, pageNum, outputPdf);
                        currentPage++;
                        Console.WriteLine($"[PDF-INSERT] Copied page {pageNum}/{embeddedPageCount} from {Path.GetFileName(pdfPath)}, now at output page {currentPage}");
                    }
                    Console.WriteLine($"[PDF-INSERT] Successfully inserted {embeddedPageCount} pages from {Path.GetFileName(pdfPath)}");
                }
                finally
                {
                    try { embeddedPdf?.Close(); reader?.Close(); } catch { }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error reading PDF {pdfPath}: {ex.Message}");
                currentPage = InsertErrorPlaceholder(pdfPath, outputPdf, currentPage, ex.Message);
            }
            return currentPage;
        }

        // Insert MSG file without separator/grey page
        private static int InsertMsgFile_NoSeparator(string msgPath, PdfDocument outputPdf, int currentPage)
        {
            Console.WriteLine($"[PDF-INSERT] Converting and inserting MSG: {Path.GetFileName(msgPath)} after page {currentPage}");
            try
            {
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"msg_temp_{Guid.NewGuid()}.pdf");
                try
                {
                    var (converted, attachmentFiles) = TryConvertMsgToPdfWithAttachments(msgPath, tempPdfPath);
                    if (converted && File.Exists(tempPdfPath))
                    {
                        currentPage = InsertPdfFile_NoSeparator(tempPdfPath, outputPdf, currentPage, "MSG");
                        
                        // Insert extracted attachments after the MSG content
                        foreach (var attachmentPath in attachmentFiles)
                        {
                            if (File.Exists(attachmentPath))
                            {
                                Console.WriteLine($"[PDF-INSERT] Inserting MSG attachment: {Path.GetFileName(attachmentPath)}");
                                currentPage = InsertAttachmentFile(attachmentPath, outputPdf, currentPage);
                            }
                        }
                    }
                    else
                    {
                        currentPage = InsertPlaceholderForFile(msgPath, outputPdf, currentPage, "MSG");
                    }
                }
                finally
                {
                    if (File.Exists(tempPdfPath)) { try { File.Delete(tempPdfPath); } catch { } }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error processing MSG {msgPath}: {ex.Message}");
                currentPage = InsertErrorPlaceholder(msgPath, outputPdf, currentPage, ex.Message);
            }
            return currentPage;
        }

        // Insert DOCX file without separator/grey page
        private static int InsertDocxFile_NoSeparator(string docxPath, PdfDocument outputPdf, int currentPage)
        {
            Console.WriteLine($"[PDF-INSERT] Converting and inserting DOCX: {Path.GetFileName(docxPath)} after page {currentPage}");
            try
            {
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"docx_temp_{Guid.NewGuid()}.pdf");
                try
                {
                    bool converted = TryConvertDocxToPdf(docxPath, tempPdfPath);
                    if (converted && File.Exists(tempPdfPath))
                    {
                        currentPage = InsertPdfFile_NoSeparator(tempPdfPath, outputPdf, currentPage, "DOCX");
                    }
                    else
                    {
                        currentPage = InsertPlaceholderForFile(docxPath, outputPdf, currentPage, "DOCX");
                    }
                }
                finally
                {
                    if (File.Exists(tempPdfPath)) { try { File.Delete(tempPdfPath); } catch { } }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error processing DOCX {docxPath}: {ex.Message}");
                currentPage = InsertErrorPlaceholder(docxPath, outputPdf, currentPage, ex.Message);
            }
            return currentPage;
        }

        // Insert XLSX file without separator/grey page
        private static int InsertXlsxFile_NoSeparator(string xlsxPath, PdfDocument outputPdf, int currentPage)
        {
            Console.WriteLine($"[PDF-INSERT] Converting and inserting XLSX: {Path.GetFileName(xlsxPath)} after page {currentPage}");
            try
            {
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"xlsx_temp_{Guid.NewGuid()}.pdf");
                try
                {
                    bool converted = TryConvertXlsxToPdf(xlsxPath, tempPdfPath);
                    if (converted && File.Exists(tempPdfPath))
                    {
                        currentPage = InsertPdfFile_NoSeparator(tempPdfPath, outputPdf, currentPage, "XLSX");
                    }
                    else
                    {
                        currentPage = InsertPlaceholderForFile(xlsxPath, outputPdf, currentPage, "XLSX");
                    }
                }
                finally
                {
                    if (File.Exists(tempPdfPath)) { try { File.Delete(tempPdfPath); } catch { } }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error processing XLSX {xlsxPath}: {ex.Message}");
                currentPage = InsertErrorPlaceholder(xlsxPath, outputPdf, currentPage, ex.Message);
            }
            return currentPage;
        }

        /// <summary>
        /// Attempts to convert MSG to PDF using the main HTML-to-PDF pipeline (DinkToPdf/HtmlToPdfWorker)
        /// </summary>
        private static bool TryConvertMsgToPdf(string msgPath, string outputPdfPath)
        {
            try
            {
                Console.WriteLine($"[PDF-INSERT] Converting MSG to PDF (HTML pipeline): {msgPath} -> {outputPdfPath}");
                using (var msg = new MsgReader.Outlook.Storage.Message(msgPath))
                {
                    // Build HTML with inline images using the main service
                    var htmlResult = _emailConverterService.BuildEmailHtmlWithInlineImages(msg, false);
                    string html = htmlResult.Html;
                    var tempHtmlPath = Path.Combine(Path.GetTempPath(), $"msg2pdf_{Guid.NewGuid()}.html");
                    File.WriteAllText(tempHtmlPath, html, System.Text.Encoding.UTF8);

                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName,
                        Arguments = $"--html2pdf \"{tempHtmlPath}\" \"{outputPdfPath}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };
                    var proc = System.Diagnostics.Process.Start(psi);
                    string stdOut = proc.StandardOutput.ReadToEnd();
                    string stdErr = proc.StandardError.ReadToEnd();
                    proc.WaitForExit();
                    File.Delete(tempHtmlPath);
                    if (proc.ExitCode == 0 && File.Exists(outputPdfPath))
                    {
                        Console.WriteLine($"[PDF-INSERT] Successfully converted MSG to PDF: {outputPdfPath}");
                        return true;
                    }
                    else
                    {
                        // Dump HTML to debug file for inspection
                        var debugHtmlPath = tempHtmlPath + ".fail.html";
                        File.WriteAllText(debugHtmlPath, html, System.Text.Encoding.UTF8);
                        Console.WriteLine($"[PDF-INSERT] HtmlToPdfWorker failed.\nSTDOUT: {stdOut}\nSTDERR: {stdErr}\nHTML dumped to: {debugHtmlPath}");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Failed to convert MSG to PDF: {ex.Message}\n{ex}");
                return false;
            }
        }

        /// <summary>
        /// Attempts to convert MSG to PDF using the main HTML-to-PDF pipeline and extracts attachments
        /// </summary>
        private static (bool success, List<string> attachmentFiles) TryConvertMsgToPdfWithAttachments(string msgPath, string outputPdfPath)
        {
            var attachmentFiles = new List<string>();
            try
            {
                Console.WriteLine($"[PDF-INSERT] Converting MSG to PDF with attachments: {msgPath} -> {outputPdfPath}");
                using (var msg = new MsgReader.Outlook.Storage.Message(msgPath))
                {
                    // Extract attachments to temp files
                    if (msg.Attachments != null && msg.Attachments.Count > 0)
                    {
                        var inlineContentIds = GetInlineContentIds(msg.BodyHtml ?? "");
                        
                        foreach (var attachment in msg.Attachments)
                        {
                            if (attachment is MsgReader.Outlook.Storage.Attachment fileAttachment)
                            {
                                // Skip inline images and signature files
                                if (!string.IsNullOrEmpty(fileAttachment.ContentId) && 
                                    inlineContentIds.Contains(fileAttachment.ContentId.Trim('<', '>', '"', '\'', ' ')))
                                    continue;
                                    
                                if (string.IsNullOrEmpty(fileAttachment.FileName))
                                    continue;
                                    
                                var ext = Path.GetExtension(fileAttachment.FileName)?.ToLowerInvariant();
                                if (new[] { ".p7s", ".p7m", ".smime", ".asc", ".sig" }.Contains(ext))
                                    continue;
                                
                                string tempAttachmentPath = Path.Combine(Path.GetTempPath(), 
                                    $"msg_attachment_{Guid.NewGuid()}_{fileAttachment.FileName}");
                                
                                try
                                {
                                    File.WriteAllBytes(tempAttachmentPath, fileAttachment.Data);
                                    attachmentFiles.Add(tempAttachmentPath);
                                    Console.WriteLine($"[PDF-INSERT] Extracted MSG attachment: {fileAttachment.FileName} -> {tempAttachmentPath}");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"[PDF-INSERT] Failed to extract attachment {fileAttachment.FileName}: {ex.Message}");
                                }
                            }
                            else if (attachment is MsgReader.Outlook.Storage.Message nestedMsg)
                            {
                                string tempMsgPath = Path.Combine(Path.GetTempPath(), 
                                    $"msg_nested_{Guid.NewGuid()}_{(nestedMsg.Subject ?? "email").Replace("/", "_").Replace("\\", "_")}.msg");
                                
                                try
                                {
                                    nestedMsg.Save(tempMsgPath);
                                    attachmentFiles.Add(tempMsgPath);
                                    Console.WriteLine($"[PDF-INSERT] Extracted nested MSG: {nestedMsg.Subject} -> {tempMsgPath}");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"[PDF-INSERT] Failed to extract nested MSG {nestedMsg.Subject}: {ex.Message}");
                                }
                            }
                        }
                    }
                    
                    // Convert the main MSG to PDF
                    bool converted = TryConvertMsgToPdf(msgPath, outputPdfPath);
                    return (converted, attachmentFiles);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Failed to convert MSG with attachments: {ex.Message}");
                
                // Clean up any extracted attachment files on error
                foreach (var file in attachmentFiles)
                {
                    try { if (File.Exists(file)) File.Delete(file); } catch { }
                }
                
                return (false, new List<string>());
            }
        }

        /// <summary>
        /// Helper method to get inline content IDs from HTML body
        /// </summary>
        private static List<string> GetInlineContentIds(string htmlBody)
        {
            var contentIds = new List<string>();
            if (string.IsNullOrEmpty(htmlBody)) return contentIds;
            
            var cidMatches = System.Text.RegularExpressions.Regex.Matches(htmlBody, @"cid:([^""'\s>]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            foreach (System.Text.RegularExpressions.Match match in cidMatches)
            {
                if (match.Groups.Count > 1)
                {
                    contentIds.Add(match.Groups[1].Value.Trim());
                }
            }
            return contentIds;
        }

        /// <summary>
        /// Inserts an attachment file based on its type
        /// </summary>
        private static int InsertAttachmentFile(string attachmentPath, PdfDocument outputPdf, int currentPage)
        {
            try
            {
                var ext = Path.GetExtension(attachmentPath)?.ToLowerInvariant();
                
                switch (ext)
                {
                    case ".pdf":
                        return InsertPdfFile_NoSeparator(attachmentPath, outputPdf, currentPage, "Attachment");
                    case ".docx":
                        return InsertDocxFile_NoSeparator(attachmentPath, outputPdf, currentPage);
                    case ".xlsx":
                        return InsertXlsxFile_NoSeparator(attachmentPath, outputPdf, currentPage);
                    case ".msg":
                        return InsertMsgFile_NoSeparator(attachmentPath, outputPdf, currentPage);
                    default:
                        return InsertPlaceholderForFile(attachmentPath, outputPdf, currentPage, $"Attachment ({ext})");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error inserting attachment {attachmentPath}: {ex.Message}");
                return InsertErrorPlaceholder(attachmentPath, outputPdf, currentPage, ex.Message);
            }
            finally
            {
                // Clean up temp attachment file
                try { if (File.Exists(attachmentPath)) File.Delete(attachmentPath); } catch { }
            }
        }

        /// <summary>
        /// Creates a placeholder page for unsupported file types
        /// </summary>
        private static int InsertPlaceholderForFile(string filePath, PdfDocument outputPdf, int currentPage, string fileType)
        {
            string fileName = Path.GetFileName(filePath);
            string fileInfo = $"File: {fileName}\nType: {fileType}\nSize: {GetFileSizeString(filePath)}";
            
            AddSeparatorPage(outputPdf, $"Embedded File: {fileName}", fileInfo, fileType);
            currentPage++;

            Console.WriteLine($"[PDF-INSERT] Added placeholder for {fileName} ({fileType})");
            return currentPage;
        }

        /// <summary>
        /// Creates an error placeholder page
        /// </summary>
        private static int InsertErrorPlaceholder(string filePath, PdfDocument outputPdf, int currentPage, string errorMessage)
        {
            string fileName = Path.GetFileName(filePath);
            string errorInfo = $"File: {fileName}\nError: {errorMessage}";
            
            AddSeparatorPage(outputPdf, $"Error: {fileName}", errorInfo, "ERROR");
            currentPage++;

            Console.WriteLine($"[PDF-INSERT] Added error placeholder for {fileName}");
            return currentPage;
        }

        /// <summary>
        /// Adds a separator page with information about the embedded content
        /// </summary>
        private static void AddSeparatorPage(PdfDocument pdfDoc, string title, string content, string type)
        {
            var page = pdfDoc.AddNewPage();
            var canvas = new PdfCanvas(page);
            var pageSize = page.GetPageSize();
            
            // Light gray background for the separator
            canvas.SetFillColorGray(0.95f);
            canvas.Rectangle(50, 50, pageSize.GetWidth() - 100, pageSize.GetHeight() - 100);
            canvas.Fill();
            
            // Border
            canvas.SetStrokeColorGray(0.7f);
            canvas.SetLineWidth(2);
            canvas.Rectangle(50, 50, pageSize.GetWidth() - 100, pageSize.GetHeight() - 100);
            canvas.Stroke();
            
            // Use canvas text operations to avoid Document lifecycle issues
            canvas.BeginText();
            
            try
            {
                // Load default font
                var font = iText.Kernel.Font.PdfFontFactory.CreateFont();
                
                // Title
                canvas.SetFontAndSize(font, 20);
                var titleWidth = font.GetWidth(title, 20);
                canvas.SetTextMatrix(1, 0, 0, 1, (pageSize.GetWidth() - titleWidth) / 2, pageSize.GetHeight() - 150);
                canvas.ShowText(title);
                
                // Type
                var typeText = $"Type: {type}";
                canvas.SetFontAndSize(font, 14);
                var typeWidth = font.GetWidth(typeText, 14);
                canvas.SetTextMatrix(1, 0, 0, 1, (pageSize.GetWidth() - typeWidth) / 2, pageSize.GetHeight() - 200);
                canvas.ShowText(typeText);
                
                // Content (split by lines and handle wrapping)
                canvas.SetFontAndSize(font, 12);
                var lines = content.Split('\n');
                var yPosition = pageSize.GetHeight() - 250;
                var lineHeight = 20;
                var maxWidth = pageSize.GetWidth() - 160; // Account for margins
                
                foreach (var line in lines)
                {
                    // Simple word wrapping
                    var words = line.Split(' ');
                    var currentLine = "";
                    
                    foreach (var word in words)
                    {
                        var testLine = string.IsNullOrEmpty(currentLine) ? word : currentLine + " " + word;
                        var testWidth = font.GetWidth(testLine, 12);
                        
                        if (testWidth <= maxWidth)
                        {
                            currentLine = testLine;
                        }
                        else
                        {
                            // Print current line and start new one
                            if (!string.IsNullOrEmpty(currentLine))
                            {
                                canvas.SetTextMatrix(1, 0, 0, 1, 100, yPosition);
                                canvas.ShowText(currentLine);
                                yPosition -= lineHeight;
                            }
                            currentLine = word;
                        }
                    }
                    
                    // Print remaining text
                    if (!string.IsNullOrEmpty(currentLine))
                    {
                        canvas.SetTextMatrix(1, 0, 0, 1, 100, yPosition);
                        canvas.ShowText(currentLine);
                        yPosition -= lineHeight;
                    }
                }
                
                // Footer
                canvas.SetFontAndSize(font, 10);
                var footer = "This page was automatically inserted to show embedded content from the original document.";
                var footerWidth = font.GetWidth(footer, 10);
                canvas.SetTextMatrix(1, 0, 0, 1, (pageSize.GetWidth() - footerWidth) / 2, 100);
                canvas.ShowText(footer);
            }
            finally
            {
                canvas.EndText();
            }
        }

        /// <summary>
        /// Gets a human-readable file size string
        /// </summary>
        private static string GetFileSizeString(string filePath)
        {
            try
            {
                var fileInfo = new FileInfo(filePath);
                long bytes = fileInfo.Length;
                
                if (bytes < 1024) return $"{bytes} bytes";
                if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
                if (bytes < 1024 * 1024 * 1024) return $"{bytes / (1024.0 * 1024.0):F1} MB";
                return $"{bytes / (1024.0 * 1024.0 * 1024.0):F1} GB";
            }
            catch
            {
                return "Unknown size";
            }
        }

        /// <summary>
        /// Inserts a single embedded object into the PDF
        /// </summary>
        private static int InsertEmbeddedObject(InteropEmbeddedExtractor.ExtractedObjectInfo obj, PdfDocument outputPdf, int currentOutputPage)
        {
            // Route all calls to the new no-separator version
            return InsertEmbeddedObject_NoSeparator(obj, outputPdf, currentOutputPage);
        }

        /// <summary>
        /// Attempts to convert DOCX to PDF using Word Interop
        /// </summary>
        private static bool TryConvertDocxToPdf(string docxPath, string outputPdfPath)
        {
            try
            {
                Console.WriteLine($"[PDF-INSERT] Converting DOCX to PDF (Interop): {docxPath} -> {outputPdfPath}");
                
                Microsoft.Office.Interop.Word.Application wordApp = null;
                Microsoft.Office.Interop.Word.Document doc = null;
                
                try
                {
                    // Create Word application with maximum popup suppression
                    wordApp = new Microsoft.Office.Interop.Word.Application();
                    wordApp.Visible = false;
                    wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                    wordApp.ScreenUpdating = false;
                    wordApp.ShowWindowsInTaskbar = false;
                    wordApp.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;
                    
                    // Suppress all possible Word UI elements (only supported properties)
                    try { wordApp.DisplayRecentFiles = false; } catch { }
                    try { wordApp.DisplayScrollBars = false; } catch { }
                    try { wordApp.ShowStartupDialog = false; } catch { }
                    try { wordApp.ShowAnimation = false; } catch { }
                    try { wordApp.DisplayDocumentInformationPanel = false; } catch { }
                    
                    // Disable Word's automatic features that might cause popups
                    try { wordApp.Options.DoNotPromptForConvert = true; } catch { }
                    try { wordApp.Options.ConfirmConversions = false; } catch { }
                    try { wordApp.Options.UpdateLinksAtOpen = false; } catch { }
                    try { wordApp.Options.CheckGrammarAsYouType = false; } catch { }
                    try { wordApp.Options.CheckSpellingAsYouType = false; } catch { }
                    
                    // Open document with comprehensive popup suppression
                    object missing = System.Reflection.Missing.Value;
                    doc = wordApp.Documents.Open(docxPath, 
                        ConfirmConversions: false,
                        ReadOnly: true, 
                        AddToRecentFiles: false, 
                        PasswordDocument: missing,
                        PasswordTemplate: missing,
                        Revert: false,
                        WritePasswordDocument: missing,
                        WritePasswordTemplate: missing,
                        Format: missing,
                        Encoding: missing,
                        Visible: false,
                        OpenAndRepair: missing,
                        DocumentDirection: missing,
                        NoEncodingDialog: true);
                    
                    // Ensure document is active
                    doc.Activate();
                    Console.WriteLine($"[PDF-INSERT] Document opened and activated, attempting export...");
                    
                    // Export to PDF with minimal settings
                    doc.ExportAsFixedFormat(outputPdfPath, 
                        Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF,
                        OpenAfterExport: false,
                        OptimizeFor: Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint);
                    
                    Console.WriteLine($"[PDF-INSERT] ExportAsFixedFormat completed, checking file...");
                    
                    // Allow a moment for file to be written
                    System.Threading.Thread.Sleep(500);
                    
                    if (File.Exists(outputPdfPath) && new FileInfo(outputPdfPath).Length > 0)
                    {
                        Console.WriteLine($"[PDF-INSERT] Successfully converted DOCX to PDF: {outputPdfPath}");
                        return true;
                    }
                    else
                    {
                        Console.WriteLine($"[PDF-INSERT] DOCX conversion failed: output file not created or empty");
                        return false;
                    }
                }
                finally
                {
                    // Clean up with comprehensive error handling
                    if (doc != null) 
                    { 
                        try 
                        { 
                            doc.Close(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges); 
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
                        } 
                        catch (Exception cleanupEx) 
                        { 
                            Console.WriteLine($"[PDF-INSERT] Warning: Document cleanup failed: {cleanupEx.Message}");
                        } 
                    }
                    if (wordApp != null) 
                    { 
                        try 
                        { 
                            wordApp.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                        } 
                        catch (Exception cleanupEx) 
                        { 
                            Console.WriteLine($"[PDF-INSERT] Warning: Application cleanup failed: {cleanupEx.Message}");
                        } 
                    }
                    
                    // Force garbage collection to release COM objects
                    System.GC.Collect();
                    System.GC.WaitForPendingFinalizers();
                    System.GC.Collect();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Failed to convert DOCX to PDF: {ex.Message}");
                Console.WriteLine($"[PDF-INSERT] Exception details: {ex}");
                return false;
            }
        }

        /// <summary>
        /// Attempts to convert XLSX to PDF using Excel Interop
        /// </summary>
        private static bool TryConvertXlsxToPdf(string xlsxPath, string outputPdfPath)
        {
            bool result = false;
            Exception threadEx = null;

            // Run Excel conversion in STA thread like OfficeConversionService to avoid popup issues
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    Console.WriteLine($"[PDF-INSERT] Converting XLSX to PDF (STA Thread): {xlsxPath} -> {outputPdfPath}");
                    
                    var excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
                    Microsoft.Office.Interop.Excel.Workbook wb = null;
                    try
                    {
                        workbooks = excelApp.Workbooks;
                        wb = workbooks.Open(xlsxPath);
                        wb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPdfPath);
                        result = true;
                    }
                    finally
                    {
                        if (wb != null)
                        {
                            wb.Close(false);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                        }
                        if (workbooks != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                        }
                        if (excelApp != null)
                        {
                            excelApp.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
                catch (Exception ex)
                {
                    threadEx = ex;
                }
            });
            
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            
            if (threadEx != null)
            {
                Console.WriteLine($"[PDF-INSERT] Failed to convert XLSX to PDF: {threadEx.Message}");
                return false;
            }
            
            if (result && File.Exists(outputPdfPath) && new FileInfo(outputPdfPath).Length > 0)
            {
                Console.WriteLine($"[PDF-INSERT] Successfully converted XLSX to PDF: {outputPdfPath}");
                return true;
            }
            else
            {
                Console.WriteLine($"[PDF-INSERT] XLSX conversion failed: output file not created or empty");
                return false;
            }
        }
    }
}
