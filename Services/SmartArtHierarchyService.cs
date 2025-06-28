using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using MsgToPdfConverter.Utils;

namespace MsgToPdfConverter.Services
{
    public static class SmartArtHierarchyService
    {
        /// <summary>
        /// Creates a SmartArt hierarchy diagram using Word Interop and exports as image
        /// </summary>
        public static bool CreateHierarchyDiagram(List<string> parentChain, string currentItem, string outputImagePath)
        {
            Microsoft.Office.Interop.Word.Application wordApp = null;
            Document doc = null;
            
            try
            {
                // Create Word application
                wordApp = new Microsoft.Office.Interop.Word.Application();
                wordApp.Visible = false;
                
                // Create a new document
                doc = wordApp.Documents.Add();
                
                // Insert a basic organizational chart using shapes
                var shapes = doc.Shapes;
                
                // Calculate positions for hierarchical tree layout
                float startX = 50;
                float startY = 50;
                float boxWidth = 250; // Much wider for longer text
                float boxHeight = 60; // Taller for better readability
                float verticalSpacing = 80;
                float horizontalSpacing = 280; // More spacing between boxes
                
                // Build complete hierarchy structure
                var allItems = new List<string>();
                if (parentChain != null) 
                {
                    foreach (var item in parentChain)
                    {
                        allItems.Add(item + ".msg"); // Add .msg for emails in chain
                    }
                }
                // Add proper extension for current item
                string currentWithExt = currentItem;
                if (!Path.HasExtension(currentItem))
                {
                    currentWithExt = AddFileExtension(currentItem, false);
                }
                allItems.Add(currentWithExt);
                
                // Create root email box at top
                if (allItems.Count > 0)
                {
                    float rootX = startX + (horizontalSpacing * 2); // Center the root
                    var rootShape = shapes.AddShape(1, rootX, startY, boxWidth, boxHeight);
                    rootShape.TextFrame.TextRange.Text = TruncateText(allItems[0], 45); // Much longer text
                    rootShape.TextFrame.TextRange.Font.Size = 12; // Larger font
                    rootShape.TextFrame.TextRange.Font.Bold = 1;
                    rootShape.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    rootShape.Fill.ForeColor.RGB = 15724527; // Light blue for email
                    rootShape.Line.ForeColor.RGB = 0; // Black border
                    rootShape.Line.Weight = 2.0f;
                    
                    // Create attachment boxes horizontally below root
                    if (allItems.Count > 1)
                    {
                        int attachmentCount = allItems.Count - 1;
                        float totalWidth = attachmentCount * horizontalSpacing;
                        float attachmentStartX = startX + (horizontalSpacing * 2) - (totalWidth / 2) + (horizontalSpacing / 2);
                        
                        for (int i = 1; i < allItems.Count; i++)
                        {
                            float attX = attachmentStartX + ((i - 1) * horizontalSpacing);
                            float attY = startY + verticalSpacing;
                            
                            var attShape = shapes.AddShape(1, attX, attY, boxWidth, boxHeight);
                            attShape.TextFrame.TextRange.Text = TruncateText(allItems[i], 40); // Much longer text
                            attShape.TextFrame.TextRange.Font.Size = 11; // Larger font
                            attShape.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                            
                            // Highlight current attachment in red, others in light gray
                            if (i == allItems.Count - 1)
                            {
                                // Current attachment - RED highlight
                                attShape.Fill.ForeColor.RGB = 255; // Red
                                attShape.TextFrame.TextRange.Font.Bold = 1;
                                attShape.Line.ForeColor.RGB = 128; // Dark red border
                                attShape.Line.Weight = 3.0f;
                            }
                            else
                            {
                                // Other attachments - light gray
                                attShape.Fill.ForeColor.RGB = 12632256; // Light gray
                                attShape.Line.ForeColor.RGB = 8421504; // Gray border
                                attShape.Line.Weight = 1.5f;
                            }
                            
                            // Connect to root with line
                            var line = shapes.AddLine(
                                rootX + (boxWidth / 2), startY + boxHeight,
                                attX + (boxWidth / 2), attY);
                            line.Line.ForeColor.RGB = 0; // Black
                            line.Line.Weight = 1.5f;
                            
                            // If this attachment has nested items, show them below
                            if (i < allItems.Count - 1)
                            {
                                // Add nested indicator
                                for (int j = i + 1; j < allItems.Count; j++)
                                {
                                    float nestedX = attX;
                                    float nestedY = attY + (verticalSpacing * (j - i));
                                    
                                    var nestedShape = shapes.AddShape(1, nestedX, nestedY, boxWidth, boxHeight);
                                    nestedShape.TextFrame.TextRange.Text = TruncateText(allItems[j], 40); // Much longer text
                                    nestedShape.TextFrame.TextRange.Font.Size = 10; // Readable font size
                                    nestedShape.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    
                                    if (j == allItems.Count - 1)
                                    {
                                        // Current item - RED
                                        nestedShape.Fill.ForeColor.RGB = 255; // Red
                                        nestedShape.TextFrame.TextRange.Font.Bold = 1;
                                        nestedShape.Line.ForeColor.RGB = 128; // Dark red
                                        nestedShape.Line.Weight = 3.0f;
                                    }
                                    else
                                    {
                                        // Parent nested items - light gray
                                        nestedShape.Fill.ForeColor.RGB = 12632256; // Light gray
                                        nestedShape.Line.ForeColor.RGB = 8421504; // Gray
                                        nestedShape.Line.Weight = 1.5f;
                                    }
                                    
                                    // Connect to parent
                                    var nestedLine = shapes.AddLine(
                                        attX + (boxWidth / 2), attY + boxHeight,
                                        nestedX + (boxWidth / 2), nestedY);
                                    nestedLine.Line.ForeColor.RGB = 0; // Black
                                    nestedLine.Line.Weight = 1.5f;
                                }
                                break; // Only show nested structure once
                            }
                        }
                    }
                }
                
                // Select all shapes to determine bounds
                float minX = float.MaxValue, minY = float.MaxValue;
                float maxX = float.MinValue, maxY = float.MinValue;
                
                foreach (Shape shape in shapes)
                {
                    minX = Math.Min(minX, shape.Left);
                    minY = Math.Min(minY, shape.Top);
                    maxX = Math.Max(maxX, shape.Left + shape.Width);
                    maxY = Math.Max(maxY, shape.Top + shape.Height);
                }
                
                // Add some padding
                float padding = 20;
                minX -= padding;
                minY -= padding;
                maxX += padding;
                maxY += padding;
                
                // Create a range that covers all shapes
                Range shapeRange = doc.Range();
                shapeRange.Select();
                
                // Copy selection as picture (using default format)
                wordApp.Selection.CopyAsPicture();
                
                // Create new document to paste and export the image
                Document imageDoc = wordApp.Documents.Add();
                
                // Paste the image
                imageDoc.Range().Paste();
                
                // Export as PDF instead of PNG (Word doesn't support PNG export directly)
                try
                {
                    // Export the document as PDF with high quality settings
                    imageDoc.ExportAsFixedFormat(
                        OutputFileName: outputImagePath,
                        ExportFormat: WdExportFormat.wdExportFormatPDF,
                        BitmapMissingFonts: true,
                        DocStructureTags: false,
                        CreateBookmarks: WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                        OptimizeFor: WdExportOptimizeFor.wdExportOptimizeForPrint,
                        Range: WdExportRange.wdExportAllDocument);
                    
                    Console.WriteLine($"[SMARTART] Successfully exported hierarchy as PDF: {outputImagePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[SMARTART] PDF export failed: {ex.Message}");
                    // Fallback: save as Word document
                    try
                    {
                        string tempDocPath = outputImagePath.Replace(".pdf", "_temp.docx");
                        imageDoc.SaveAs2(tempDocPath, FileFormat: WdSaveFormat.wdFormatDocumentDefault);
                        
                        // Try to export the saved document as PDF
                        imageDoc.Close(SaveChanges: false);
                        Document savedDoc = wordApp.Documents.Open(tempDocPath);
                        savedDoc.ExportAsFixedFormat(
                            OutputFileName: outputImagePath,
                            ExportFormat: WdExportFormat.wdExportFormatPDF);
                        savedDoc.Close(SaveChanges: false);
                        
                        // Clean up temp doc
                        try { File.Delete(tempDocPath); } catch { }
                        Console.WriteLine($"[SMARTART] Successfully exported hierarchy via fallback method: {outputImagePath}");
                    }
                    catch (Exception fallbackEx)
                    {
                        Console.WriteLine($"[SMARTART] Fallback PDF export also failed: {fallbackEx.Message}");
                        // Return false to use text fallback
                        imageDoc.Close(SaveChanges: false);
                        doc.Close(SaveChanges: false);
                        wordApp.Quit(SaveChanges: false);
                        return false;
                    }
                }
                
                imageDoc.Close(SaveChanges: false);
                
                return File.Exists(outputImagePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SMARTART] Error creating hierarchy diagram: {ex.Message}");
                return false;
            }
            finally
            {
                // Clean up
                try
                {
                    if (doc != null)
                    {
                        doc.Close(SaveChanges: false);
                        Marshal.ReleaseComObject(doc);
                    }
                    
                    if (wordApp != null)
                    {
                        wordApp.Quit(SaveChanges: false);
                        Marshal.ReleaseComObject(wordApp);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[SMARTART] Error during cleanup: {ex.Message}");
                }
                
                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        
        /// <summary>
        /// Creates a hierarchy header PDF with visual diagram
        /// </summary>
        public static bool CreateHierarchyHeaderPdf(List<string> parentChain, string currentItem, string headerText, string outputPdfPath)
        {
            try
            {
                // Try to create the visual hierarchy diagram
                string tempDiagramPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + "_hierarchy.pdf");
                
                if (CreateHierarchyDiagram(parentChain, currentItem, tempDiagramPath))
                {
                    // Create combined PDF with header text and hierarchy diagram
                    using (var writer = new iText.Kernel.Pdf.PdfWriter(outputPdfPath))
                    using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                    using (var document = new iText.Layout.Document(pdf))
                    {
                        // Add main header
                        var headerParagraph = new iText.Layout.Element.Paragraph(headerText)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetFontSize(14)
                            .SetBold();
                        document.Add(headerParagraph);
                        
                        // Add spacing
                        document.Add(new iText.Layout.Element.Paragraph("\n"));
                        
                        // Add hierarchy diagram by merging the PDF
                        if (File.Exists(tempDiagramPath))
                        {
                            try
                            {
                                // Read the hierarchy diagram PDF and add it as content
                                using (var hierarchyReader = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(tempDiagramPath)))
                                {
                                    if (hierarchyReader.GetNumberOfPages() > 0)
                                    {
                                        var hierarchyPage = hierarchyReader.GetPage(1);
                                        var formXObject = hierarchyPage.CopyAsFormXObject(pdf);
                                        
                                        // Add the hierarchy diagram as an image element
                                        var hierarchyImage = new iText.Layout.Element.Image(formXObject);
                                        hierarchyImage.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
                                        hierarchyImage.SetMaxWidth(500);
                                        hierarchyImage.SetMaxHeight(200);
                                        document.Add(hierarchyImage);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[SMARTART] Error adding hierarchy diagram to PDF: {ex.Message}");
                                // Fall back to text hierarchy
                                string treeHeader = TreeHeaderHelper.BuildTreeHeader(parentChain, currentItem);
                                var treeParagraph = new iText.Layout.Element.Paragraph("Hierarchy:\n" + treeHeader)
                                    .SetFontSize(10)
                                    .SetFontFamily("Courier");
                                document.Add(treeParagraph);
                            }
                            finally
                            {
                                try { File.Delete(tempDiagramPath); } catch { }
                            }
                        }
                        
                        // Add spacing after hierarchy
                        document.Add(new iText.Layout.Element.Paragraph("\n"));
                    }
                    
                    return true;
                }
                else
                {
                    // Fall back to text-only hierarchy
                    string treeHeader = TreeHeaderHelper.BuildTreeHeader(parentChain, currentItem);
                    string combinedHeader = $"{headerText}\n\nHierarchy:\n{treeHeader}";
                    PdfService.AddHeaderPdf(outputPdfPath, combinedHeader);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[SMARTART] Error creating hierarchy header PDF: {ex.Message}");
                // Fall back to basic header
                try
                {
                    PdfService.AddHeaderPdf(outputPdfPath, headerText);
                    return false;
                }
                catch (Exception fallbackEx)
                {
                    Console.WriteLine($"[SMARTART] Error creating fallback header: {fallbackEx.Message}");
                    return false;
                }
            }
        }
        
        /// <summary>
        /// Adds appropriate file extension to display text
        /// </summary>
        private static string AddFileExtension(string fileName, bool isEmail)
        {
            if (string.IsNullOrEmpty(fileName))
                return "Unknown";

            // If it's an email, add .msg extension
            if (isEmail)
            {
                return fileName + ".msg";
            }

            // For attachments, ensure they have an extension
            if (!Path.HasExtension(fileName))
            {
                // Try to guess extension based on name patterns
                string lowerName = fileName.ToLower();
                if (lowerName.Contains("word") || lowerName.Contains("doc"))
                    return fileName + ".docx";
                else if (lowerName.Contains("excel") || lowerName.Contains("sheet"))
                    return fileName + ".xlsx";
                else if (lowerName.Contains("pdf"))
                    return fileName + ".pdf";
                else if (lowerName.Contains("zip") || lowerName.Contains("archive"))
                    return fileName + ".zip";
                else if (lowerName.Contains("image") || lowerName.Contains("picture"))
                    return fileName + ".png";
                else
                    return fileName + ".file"; // Generic extension
            }

            return fileName; // Already has extension
        }

        /// <summary>
        /// Truncates text to a specified length for better display
        /// </summary>
        private static string TruncateText(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text))
                return "Unknown";
                
            if (text.Length <= maxLength)
                return text;
                
            return text.Substring(0, maxLength - 3) + "...";
        }
    }
}
