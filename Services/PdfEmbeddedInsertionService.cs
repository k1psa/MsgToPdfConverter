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

            // Group extracted objects by page number and sort by page
            var objectsByPage = validObjects
                .Where(obj => File.Exists(obj.FilePath))
                .GroupBy(obj => obj.PageNumber > 0 ? obj.PageNumber : int.MaxValue) // Insert unknown page objects at the end
                .OrderBy(g => g.Key)
                .ToList();

            // Also separate objects that should be inserted at the end (page -1 or 0)
            var objectsAtEnd = validObjects
                .Where(obj => obj.PageNumber <= 0 && File.Exists(obj.FilePath))
                .ToList();

            try
            {
                using (var outputStream = new FileStream(outputPdfPath, FileMode.Create, FileAccess.Write))
                using (var pdfWriter = new PdfWriter(outputStream))
                using (var outputPdf = new PdfDocument(pdfWriter))
                {
                    int mainPageCount;
                    int currentOutputPage = 0;

                    // First, copy all pages from the main PDF
                    using (var mainPdf = new PdfDocument(new PdfReader(mainPdfPath)))
                    {
                        mainPageCount = mainPdf.GetNumberOfPages();

                        for (int mainPage = 1; mainPage <= mainPageCount; mainPage++)
                        {
                            // Copy the current page from main PDF
                            mainPdf.CopyPagesTo(mainPage, mainPage, outputPdf);
                            currentOutputPage++;

                            Console.WriteLine($"[PDF-INSERT] Copied main page {mainPage} to output page {currentOutputPage}");

                            // Check if there are embedded objects for this page (excluding end-of-document objects)
                            var objectsForThisPage = objectsByPage.FirstOrDefault(g => g.Key == mainPage);
                            if (objectsForThisPage != null)
                            {
                                foreach (var obj in objectsForThisPage)
                                {
                                    currentOutputPage = InsertEmbeddedObject(obj, outputPdf, currentOutputPage);
                                }
                            }
                        }
                    } // mainPdf is disposed here, but outputPdf remains open

                    // Insert all objects that couldn't be placed at specific pages (page <= 0) at the end
                    if (objectsAtEnd.Count > 0)
                    {
                        Console.WriteLine($"[PDF-INSERT] Inserting {objectsAtEnd.Count} objects at the end of the document");
                        foreach (var obj in objectsAtEnd)
                        {
                            currentOutputPage = InsertEmbeddedObject(obj, outputPdf, currentOutputPage);
                        }
                    }
                }

                Console.WriteLine($"[PDF-INSERT] Successfully created PDF with embedded files: {outputPdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error creating PDF with embedded files: {ex.Message}");
                
                // Fallback: just copy the main PDF
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

        /// <summary>
        /// Inserts a PDF file into the output document
        /// </summary>
        private static int InsertPdfFile(string pdfPath, PdfDocument outputPdf, int currentPage, string oleClass)
        {
            Console.WriteLine($"[PDF-INSERT] Inserting PDF: {Path.GetFileName(pdfPath)} after page {currentPage}");

            try
            {
                // First, validate the PDF file
                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine($"[PDF-INSERT] PDF file not found: {pdfPath}");
                    return InsertErrorPlaceholder(pdfPath, outputPdf, currentPage, "File not found");
                }

                // Check file size
                var fileInfo = new FileInfo(pdfPath);
                if (fileInfo.Length == 0)
                {
                    Console.WriteLine($"[PDF-INSERT] PDF file is empty: {pdfPath}");
                    return InsertErrorPlaceholder(pdfPath, outputPdf, currentPage, "Empty file");
                }

                Console.WriteLine($"[PDF-INSERT] Reading PDF file ({fileInfo.Length} bytes): {pdfPath}");

                // Try to create a reader with more robust error handling
                PdfReader reader = null;
                PdfDocument embeddedPdf = null;
                
                try
                {
                    reader = new PdfReader(pdfPath);
                    embeddedPdf = new PdfDocument(reader);
                    
                    int embeddedPageCount = embeddedPdf.GetNumberOfPages();
                    Console.WriteLine($"[PDF-INSERT] PDF has {embeddedPageCount} pages");
                    
                    // Add a separator page with information about the embedded file
                    AddSeparatorPage(outputPdf, $"Embedded PDF: {Path.GetFileName(pdfPath)}", $"Original location: Page {currentPage}", oleClass);
                    currentPage++;

                    // Copy all pages from the embedded PDF one by one for better error handling
                    for (int pageNum = 1; pageNum <= embeddedPageCount; pageNum++)
                    {
                        try
                        {
                            embeddedPdf.CopyPagesTo(pageNum, pageNum, outputPdf);
                            Console.WriteLine($"[PDF-INSERT] Copied page {pageNum}/{embeddedPageCount} from {Path.GetFileName(pdfPath)}");
                        }
                        catch (Exception pageEx)
                        {
                            Console.WriteLine($"[PDF-INSERT] Error copying page {pageNum}: {pageEx.Message}");
                            // Continue with next page
                        }
                    }
                    
                    currentPage += embeddedPageCount;
                    Console.WriteLine($"[PDF-INSERT] Successfully inserted {embeddedPageCount} pages from {Path.GetFileName(pdfPath)}");
                }
                finally
                {
                    // Explicit cleanup
                    try
                    {
                        embeddedPdf?.Close();
                        reader?.Close();
                    }
                    catch (Exception disposeEx)
                    {
                        Console.WriteLine($"[PDF-INSERT] Error disposing PDF resources: {disposeEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error reading PDF {pdfPath}: {ex.Message}");
                Console.WriteLine($"[PDF-INSERT] Exception details: {ex}");
                currentPage = InsertErrorPlaceholder(pdfPath, outputPdf, currentPage, ex.Message);
            }

            return currentPage;
        }

        /// <summary>
        /// Converts and inserts an MSG file
        /// </summary>
        private static int InsertMsgFile(string msgPath, PdfDocument outputPdf, int currentPage)
        {
            Console.WriteLine($"[PDF-INSERT] Converting and inserting MSG: {Path.GetFileName(msgPath)} after page {currentPage}");

            try
            {
                // Create a temporary PDF for the MSG content
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"msg_temp_{Guid.NewGuid()}.pdf");
                
                try
                {
                    // Try to convert MSG to PDF using existing services
                    bool converted = TryConvertMsgToPdf(msgPath, tempPdfPath);
                    
                    if (converted && File.Exists(tempPdfPath))
                    {
                        currentPage = InsertPdfFile(tempPdfPath, outputPdf, currentPage, "MSG");
                    }
                    else
                    {
                        // Fallback: create a placeholder for the MSG file
                        currentPage = InsertPlaceholderForFile(msgPath, outputPdf, currentPage, "MSG");
                    }
                }
                finally
                {
                    // Clean up temp file
                    if (File.Exists(tempPdfPath))
                    {
                        try { File.Delete(tempPdfPath); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error processing MSG {msgPath}: {ex.Message}");
                currentPage = InsertErrorPlaceholder(msgPath, outputPdf, currentPage, ex.Message);
            }

            return currentPage;
        }

        /// <summary>
        /// Attempts to convert MSG to PDF using existing conversion services
        /// </summary>
        private static bool TryConvertMsgToPdf(string msgPath, string outputPdfPath)
        {
            try
            {
                // This would use the existing MSG conversion logic
                // For now, we'll create a simple placeholder
                // In a full implementation, this would call the existing MSG to PDF conversion
                return false; // Placeholder - implement actual MSG conversion
            }
            catch
            {
                return false;
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
            try
            {
                if (obj.FilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    // Insert PDF file
                    return InsertPdfFile(obj.FilePath, outputPdf, currentOutputPage, obj.OleClass);
                }
                else if (obj.FilePath.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                {
                    // Convert MSG to PDF first, then insert
                    return InsertMsgFile(obj.FilePath, outputPdf, currentOutputPage);
                }
                else
                {
                    // Create a placeholder PDF for other file types
                    return InsertPlaceholderForFile(obj.FilePath, outputPdf, currentOutputPage, obj.OleClass);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[PDF-INSERT] Error inserting {obj.FilePath}: {ex.Message}");
                // Insert error placeholder
                return InsertErrorPlaceholder(obj.FilePath, outputPdf, currentOutputPage, ex.Message);
            }
        }
    }
}
