using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace MsgToPdfConverter.Services
{
    public class WordSmartArtHierarchyService
    {
        public string CreateHierarchyPdf(List<string> hierarchyChain, string currentAttachment, string outputFolder)
        {
            Application wordApp = null;
            Document doc = null;
            string tempDocPath = null;
            string outputPdfPath = null;

            try
            {
                if (hierarchyChain == null || hierarchyChain.Count == 0)
                    return null;

                // Process the complete hierarchy chain to determine file types and extensions
                var processedChain = new List<string>();
                int currentIndex = -1; // Track which item should be highlighted
                
                for (int i = 0; i < hierarchyChain.Count; i++)
                {
                    string item = hierarchyChain[i];
                    bool isEmail;
                    
                    if (i < hierarchyChain.Count - 1)
                    {
                        // All items except the last are definitely emails
                        isEmail = true;
                    }
                    else
                    {
                        // For the last item, check if it looks like an email subject
                        isEmail = IsLikelyEmailSubject(item);
                    }
                    
                    string processedItem = AddFileExtension(item, isEmail);
                    processedChain.Add(processedItem);
                    
                    // Check if this item matches the current attachment we should highlight
                    if (string.Equals(item, currentAttachment, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(processedItem, currentAttachment, StringComparison.OrdinalIgnoreCase))
                    {
                        currentIndex = i;
                    }
                }
                
                // If we couldn't find a match, highlight the last item as fallback
                if (currentIndex == -1)
                {
                    currentIndex = processedChain.Count - 1;
                }

                Console.WriteLine($"[WORD-HIERARCHY] Creating SmartArt for chain: {string.Join(" -> ", processedChain)}");
                Console.WriteLine($"[WORD-HIERARCHY] Highlighting item at index {currentIndex}: {processedChain[currentIndex]}");

                // Create Word application
                wordApp = new Application();
                wordApp.Visible = false;
                wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;

                // Create a new document
                doc = wordApp.Documents.Add();

                // Add title
                Paragraph titleParagraph = doc.Paragraphs.Add();
                titleParagraph.Range.Text = "Email Attachment Hierarchy";
                titleParagraph.Range.Font.Name = "Arial";
                titleParagraph.Range.Font.Size = 16;
                titleParagraph.Range.Font.Bold = 1;
                titleParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                titleParagraph.Range.InsertParagraphAfter();

                // Add some space
                Paragraph spaceParagraph = doc.Paragraphs.Add();
                spaceParagraph.Range.InsertParagraphAfter();

                // Insert SmartArt hierarchy using Shapes with reflection
                Range smartArtRange = doc.Paragraphs[doc.Paragraphs.Count].Range;
                
                try
                {
                    Console.WriteLine("[WORD-HIERARCHY] Attempting to create SmartArt...");
                    
                    // Try to add SmartArt using reflection to avoid compile-time dependencies
                    object smartArtShape = doc.Shapes.GetType().InvokeMember(
                        "AddSmartArt",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null,
                        doc.Shapes,
                        new object[] { 21, 100.0f, 100.0f, 400.0f, 300.0f, smartArtRange }
                    );

                    if (smartArtShape != null)
                    {
                        Console.WriteLine("[WORD-HIERARCHY] SmartArt shape created successfully");
                        
                        // Access SmartArt property using reflection
                        object smartArt = smartArtShape.GetType().InvokeMember(
                            "SmartArt",
                            System.Reflection.BindingFlags.GetProperty,
                            null,
                            smartArtShape,
                            null
                        );

                        if (smartArt != null)
                        {
                            Console.WriteLine("[WORD-HIERARCHY] SmartArt object accessed successfully");
                            
                            // Get AllNodes collection
                            object allNodes = smartArt.GetType().InvokeMember(
                                "AllNodes",
                                System.Reflection.BindingFlags.GetProperty,
                                null,
                                smartArt,
                                null
                            );

                            // Clear existing nodes
                            int nodeCount = (int)allNodes.GetType().InvokeMember(
                                "Count",
                                System.Reflection.BindingFlags.GetProperty,
                                null,
                                allNodes,
                                null
                            );

                            while (nodeCount > 0)
                            {
                                object firstNode = allNodes.GetType().InvokeMember(
                                    "Item",
                                    System.Reflection.BindingFlags.InvokeMethod,
                                    null,
                                    allNodes,
                                    new object[] { 1 }
                                );
                                
                                firstNode.GetType().InvokeMember(
                                    "Delete",
                                    System.Reflection.BindingFlags.InvokeMethod,
                                    null,
                                    firstNode,
                                    null
                                );
                                
                                nodeCount = (int)allNodes.GetType().InvokeMember(
                                    "Count",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null,
                                    allNodes,
                                    null
                                );
                            }

                            // Add nodes for hierarchy items
                            object parentNode = null;
                            for (int i = 0; i < processedChain.Count; i++)
                            {
                                string item = processedChain[i];
                                bool isCurrent = i == currentIndex;
                                
                                object node;
                                if (i == 0)
                                {
                                    // First node (root)
                                    node = allNodes.GetType().InvokeMember(
                                        "Add",
                                        System.Reflection.BindingFlags.InvokeMethod,
                                        null,
                                        allNodes,
                                        null
                                    );
                                    parentNode = node;
                                }
                                else
                                {
                                    // Child nodes
                                    node = parentNode.GetType().InvokeMember(
                                        "AddNode",
                                        System.Reflection.BindingFlags.InvokeMethod,
                                        null,
                                        parentNode,
                                        new object[] { 1 } // msoSmartArtNodeAfter = 1
                                    );
                                    parentNode = node;
                                }
                                
                                // Set node text using reflection
                                object textFrame2 = node.GetType().InvokeMember(
                                    "TextFrame2",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null,
                                    node,
                                    null
                                );
                                
                                object textRange = textFrame2.GetType().InvokeMember(
                                    "TextRange",
                                    System.Reflection.BindingFlags.GetProperty,
                                    null,
                                    textFrame2,
                                    null
                                );
                                
                                textRange.GetType().InvokeMember(
                                    "Text",
                                    System.Reflection.BindingFlags.SetProperty,
                                    null,
                                    textRange,
                                    new object[] { item }
                                );
                                
                                // Style the node (basic approach, colors may not work perfectly)
                                if (isCurrent)
                                {
                                    Console.WriteLine($"[WORD-HIERARCHY] Highlighting current item: {item}");
                                    // Make text bold for current item
                                    object font = textRange.GetType().InvokeMember(
                                        "Font",
                                        System.Reflection.BindingFlags.GetProperty,
                                        null,
                                        textRange,
                                        null
                                    );
                                    
                                    font.GetType().InvokeMember(
                                        "Bold",
                                        System.Reflection.BindingFlags.SetProperty,
                                        null,
                                        font,
                                        new object[] { 1 }
                                    );
                                }
                            }
                            
                            Console.WriteLine("[WORD-HIERARCHY] SmartArt hierarchy created successfully");
                        }
                        else
                        {
                            throw new Exception("Could not access SmartArt property");
                        }
                    }
                    else
                    {
                        throw new Exception("Could not create SmartArt shape");
                    }
                }
                catch (Exception smartArtEx)
                {
                    Console.WriteLine($"[WORD-HIERARCHY] SmartArt creation failed, using table fallback: {smartArtEx.Message}");
                    
                    // Fallback to table approach if SmartArt fails
                    CreateHierarchyTable(doc, processedChain, currentIndex);
                }

                // Save as temporary Word document
                tempDocPath = Path.Combine(outputFolder, $"hierarchy_{Guid.NewGuid()}.docx");
                doc.SaveAs2(tempDocPath);

                // Export to PDF
                outputPdfPath = Path.Combine(outputFolder, $"hierarchy_{Guid.NewGuid()}.pdf");
                doc.ExportAsFixedFormat(
                    OutputFileName: outputPdfPath,
                    ExportFormat: WdExportFormat.wdExportFormatPDF,
                    OpenAfterExport: false,
                    OptimizeFor: WdExportOptimizeFor.wdExportOptimizeForPrint,
                    BitmapMissingFonts: true,
                    DocStructureTags: false,
                    CreateBookmarks: WdExportCreateBookmarks.wdExportCreateNoBookmarks
                );

                Console.WriteLine($"[WORD-HIERARCHY] Successfully created PDF: {outputPdfPath}");
                return outputPdfPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[WORD-HIERARCHY] Error creating SmartArt hierarchy: {ex.Message}");
                return null;
            }
            finally
            {
                // Clean up
                try
                {
                    if (doc != null)
                    {
                        doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                        Marshal.ReleaseComObject(doc);
                    }
                    if (wordApp != null)
                    {
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                    }
                    
                    // Delete temporary Word document
                    if (!string.IsNullOrEmpty(tempDocPath) && File.Exists(tempDocPath))
                    {
                        File.Delete(tempDocPath);
                    }
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine($"[WORD-HIERARCHY] Cleanup warning: {cleanupEx.Message}");
                }
                
                // Force garbage collection to release COM objects
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private bool IsLikelyEmailSubject(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            // If it already has a file extension, it's definitely an attachment
            if (Path.HasExtension(text))
                return false;
                
            string lowerText = text.ToLower();
            
            // Check for common email subject patterns
            if (lowerText.StartsWith("re:") || lowerText.StartsWith("fw:") || lowerText.StartsWith("fwd:"))
                return true;
                
            // Check for common email-related words
            if (lowerText.Contains("request") || lowerText.Contains("approval") || 
                lowerText.Contains("meeting") || lowerText.Contains("notification") ||
                lowerText.Contains("response") || lowerText.Contains("inquiry") ||
                lowerText.Contains("follow") || lowerText.Contains("update"))
                return true;
                
            // Check if it contains typical email subject indicators (colons, dashes in business context)
            if ((lowerText.Contains(" - ") || lowerText.Contains(": ")) && 
                (lowerText.Length > 20)) // Longer subjects are more likely emails
                return true;
                
            // If it looks like a short code or identifier, it might be an email
            if (System.Text.RegularExpressions.Regex.IsMatch(text, @"^[A-Z0-9\-_]+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) && text.Length < 15)
                return true;
                
            return true; // Default to email if no file extension (most hierarchy items are emails)
        }

        private string AddFileExtension(string fileName, bool isEmail)
        {
            if (string.IsNullOrEmpty(fileName))
                return "Unknown";

            // If it already has an extension, keep it as-is
            if (Path.HasExtension(fileName))
            {
                return fileName;
            }

            // If it's an email, add .msg extension
            if (isEmail)
            {
                return fileName + ".msg";
            }

            // For attachments without extensions, try to guess based on content
            string lowerName = fileName.ToLower();
            
            // Check for common document types
            if (lowerName.Contains("word") || lowerName.Contains("doc"))
                return fileName + ".docx";
            else if (lowerName.Contains("excel") || lowerName.Contains("sheet") || lowerName.Contains("xls"))
                return fileName + ".xlsx";
            else if (lowerName.Contains("powerpoint") || lowerName.Contains("ppt"))
                return fileName + ".pptx";
            else if (lowerName.Contains("pdf"))
                return fileName + ".pdf";
            else if (lowerName.Contains("zip") || lowerName.Contains("archive"))
                return fileName + ".zip";
            else if (lowerName.Contains("image") || lowerName.Contains("picture") || lowerName.Contains("photo"))
                return fileName + ".png";
            else if (lowerName.Contains("text") || lowerName.Contains("note"))
                return fileName + ".txt";
            else
                return fileName + ".file"; // Generic fallback
        }

        private void CreateHierarchyTable(Document doc, List<string> processedChain, int currentIndex)
        {
            // Create a table to represent the hierarchy
            Range tableRange = doc.Paragraphs[doc.Paragraphs.Count].Range;
            Table hierarchyTable = doc.Tables.Add(tableRange, processedChain.Count, 1);
            
            // Configure table appearance
            hierarchyTable.Borders.Enable = 1;
            hierarchyTable.Range.Font.Name = "Arial";
            hierarchyTable.Range.Font.Size = 12;
            
            // Add hierarchy items to table
            for (int i = 0; i < processedChain.Count; i++)
            {
                string item = processedChain[i];
                bool isCurrent = i == currentIndex;
                
                Cell cell = hierarchyTable.Cell(i + 1, 1);
                
                // Add indentation based on hierarchy level
                string indentedText = new string(' ', i * 4) + "└─ " + item;
                if (i == 0)
                    indentedText = item; // Root item has no indent
                else if (i > 0)
                    indentedText = new string(' ', i * 2) + "└─ " + item;
                
                cell.Range.Text = indentedText;
                
                // Highlight current item
                if (isCurrent)
                {
                    cell.Shading.BackgroundPatternColor = WdColor.wdColorRed;
                    cell.Range.Font.Color = WdColor.wdColorWhite;
                    cell.Range.Font.Bold = 1;
                }
                else
                {
                    cell.Shading.BackgroundPatternColor = WdColor.wdColorLightBlue;
                    cell.Range.Font.Color = WdColor.wdColorBlack;
                    cell.Range.Font.Bold = 0;
                }
                
                // Add some padding
                cell.TopPadding = 8;
                cell.BottomPadding = 8;
                cell.LeftPadding = 12;
                cell.RightPadding = 12;
            }
        }
    }
}
