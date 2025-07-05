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

            // Validate and log extracted objects
            var validObjects = new List<InteropEmbeddedExtractor.ExtractedObjectInfo>();
            foreach (var obj in extractedObjects)
            {
                Console.WriteLine($"[PDF-INSERT] Checking object: {Path.GetFileName(obj.FilePath)} at {obj.FilePath}");
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
            var objectsByPage = validObjects
                .Where(obj => obj.PageNumber > 0)
                .OrderBy(obj => obj.PageNumber)
                .ThenBy(obj => obj.DocumentOrderIndex)
                .ToList();

            // Log the insertion plan
            Console.WriteLine($"[PDF-INSERT] Insertion plan:");
            foreach (var obj in objectsByPage)
            {
                Console.WriteLine($"  - {Path.GetFileName(obj.FilePath)} -> after page {obj.PageNumber} (order: {obj.DocumentOrderIndex})");
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
                    int nextObjIdx = 0;

                    Console.WriteLine($"[PDF-INSERT] Main PDF has {mainPageCount} pages");

                    // Validate that no object requests insertion after a non-existent page
                    foreach (var obj in objectsByPage)
                    {
                        if (obj.PageNumber > mainPageCount)
                        {
                            Console.WriteLine($"[PDF-INSERT] Warning: Object {Path.GetFileName(obj.FilePath)} requests insertion after page {obj.PageNumber}, but main PDF only has {mainPageCount} pages. Adjusting to page {mainPageCount}.");
                            obj.PageNumber = mainPageCount;
                        }
                    }

                    // Re-sort after potential adjustments
                    objectsByPage = objectsByPage.OrderBy(obj => obj.PageNumber).ThenBy(obj => obj.DocumentOrderIndex).ToList();

                    // For each main PDF page, copy the page, then insert any embedded objects whose PageNumber == current main page
                    for (int mainPage = 1; mainPage <= mainPageCount; mainPage++)
                    {
                        mainPdf.CopyPagesTo(mainPage, mainPage, outputPdf);
                        currentOutputPage++;
                        Console.WriteLine($"[PDF-INSERT] Copied main page {mainPage} to output page {currentOutputPage}");

                        // Insert all embedded objects whose PageNumber == mainPage (immediately after this page)
                        while (nextObjIdx < objectsByPage.Count && objectsByPage[nextObjIdx].PageNumber == mainPage)
                        {
                            var obj = objectsByPage[nextObjIdx];
                            // Insert PDF or MSG directly, do NOT add separator/grey page
                            currentOutputPage = InsertEmbeddedObject_NoSeparator(obj, outputPdf, currentOutputPage);
                            nextObjIdx++;
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
                    for (int pageNum = 1; pageNum <= embeddedPageCount; pageNum++)
                    {
                        embeddedPdf.CopyPagesTo(pageNum, pageNum, outputPdf);
                        currentPage++;
                        Console.WriteLine($"[PDF-INSERT] Copied page {pageNum}/{embeddedPageCount} from {Path.GetFileName(pdfPath)}");
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
                    bool converted = TryConvertMsgToPdf(msgPath, tempPdfPath);
                    if (converted && File.Exists(tempPdfPath))
                    {
                        currentPage = InsertPdfFile_NoSeparator(tempPdfPath, outputPdf, currentPage, "MSG");
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
                    wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false, DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone };
                    doc = wordApp.Documents.Open(docxPath, ReadOnly: true, Visible: false);
                    
                    doc.SaveAs2(outputPdfPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                    
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
                    if (doc != null) { try { doc.Close(false); } catch { } }
                    if (wordApp != null) { try { wordApp.Quit(false); } catch { } }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Failed to convert DOCX to PDF: {ex.Message}");
                return false;
            }
        }

        // Simple HTML tag stripper for fallback
        private static string StripHtml(string html)
        {
            if (string.IsNullOrEmpty(html)) return string.Empty;
            var array = new char[html.Length];
            int arrayIndex = 0;
            bool inside = false;
            foreach (char let in html)
            {
                if (let == '<') { inside = true; continue; }
                if (let == '>') { inside = false; continue; }
                if (!inside) array[arrayIndex++] = let;
            }
            return new string(array, 0, arrayIndex);
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
    }
}
