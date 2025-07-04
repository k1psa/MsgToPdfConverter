using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;

namespace MsgToPdfConverter.Utils
{
    public class InteropEmbeddedExtractor
    {
        public class ExtractedObjectInfo
        {
            public string FilePath { get; set; }
            public int PageNumber { get; set; } // 1-based page number
            public string OleClass { get; set; }
        }

        /// <summary>
        /// Extracts embedded OLE objects from a .doc or .docx file using Word Interop, saving them to the specified output directory.
        /// Returns a list of extracted file info, including the page number where each object was found.
        /// </summary>
        public static List<ExtractedObjectInfo> ExtractEmbeddedObjects(string docxPath, string outputDir)
        {
            var results = new List<ExtractedObjectInfo>();
            Application wordApp = null;
            Document doc = null;
            int counter = 1;
            bool interopSuccess = false;
            try
            {
                Console.WriteLine($"[InteropExtractor] ExtractEmbeddedObjects called for: {docxPath}");
                wordApp = new Application { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
                doc = wordApp.Documents.Open(docxPath, ReadOnly: true, Visible: false);

                Console.WriteLine($"[InteropExtractor] InlineShapes count: {doc.InlineShapes.Count}");
                int found = 0;
                // InlineShapes (OLE objects, e.g. embedded PDFs, Excels, etc.)
                foreach (InlineShape ish in doc.InlineShapes)
                {
                    Console.WriteLine($"[InteropExtractor] InlineShape: Type={ish.Type}, OLE ProgID={ish.OLEFormat?.ProgID}");
                    if (ish.Type == WdInlineShapeType.wdInlineShapeEmbeddedOLEObject)
                    {
                        found++;
                        var ole = ish.OLEFormat;
                        string ext = GetExtensionFromProgID(ole.ProgID);
                        string outFile = Path.Combine(outputDir, $"Embedded_{counter}{ext}");
                        counter++;
                        try
                        {
                            if ((ole.ProgID != null && ole.ProgID.ToLowerInvariant() == "package"))
                            {
                                // Special handling for OLE Package: try SaveAs/SaveToFile via reflection
                                bool saved = false;
                                dynamic obj = ole.Object;
                                var type = obj?.GetType();
                                if (type != null)
                                {
                                    var saveAs = type.GetMethod("SaveAs");
                                    if (saveAs != null)
                                    {
                                        saveAs.Invoke(obj, new object[] { outFile });
                                        saved = true;
                                    }
                                    else
                                    {
                                        var saveToFile = type.GetMethod("SaveToFile");
                                        if (saveToFile != null)
                                        {
                                            saveToFile.Invoke(obj, new object[] { outFile });
                                            saved = true;
                                        }
                                    }
                                }
                                if (!saved)
                                {
                                    Console.WriteLine($"[InteropExtractor] No SaveAs/SaveToFile method for OLE Package object: {type}");
                                    throw new InvalidOperationException("Cannot extract OLE Package data: no SaveAs/SaveToFile");
                                }
                            }
                            else
                            {
                                // Try to save the embedded object if possible
                                SaveOleObjectToFile(ole, outFile);
                            }
                            int page = (int)ish.Range.get_Information(WdInformation.wdActiveEndPageNumber);
                            results.Add(new ExtractedObjectInfo { FilePath = outFile, PageNumber = page, OleClass = ole.ProgID });
                            Console.WriteLine($"[InteropExtractor] Extracted: {outFile} (page {page}, ProgID={ole.ProgID})");
                            interopSuccess = true;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[InteropExtractor] Extraction error: {ex.Message}");
                        }
                    }
                }
                Console.WriteLine($"[InteropExtractor] Embedded OLE InlineShapes found: {found}");

                // Shapes (floating OLE objects) - SKIPPED: Requires office.dll (MsoShapeType)
                // foreach (Shape shape in doc.Shapes)
                // {
                //     // 7 = msoEmbeddedOLEObject (no COM reference needed)
                //     int shapeType = Convert.ToInt32(shape.Type);
                //     if (shapeType == 7)
                //     {
                //         var ole = shape.OLEFormat;
                //         string ext = GetExtensionFromProgID(ole.ProgID);
                //         string outFile = Path.Combine(outputDir, $"Embedded_{counter}{ext}");
                //         counter++;
                //         try
                //         {
                //             SaveOleObjectToFile(ole, outFile);
                //             int page = (int)shape.Anchor.get_Information(WdInformation.wdActiveEndPageNumber);
                //             results.Add(new ExtractedObjectInfo { FilePath = outFile, PageNumber = page, OleClass = ole.ProgID });
                //         }
                //         catch (Exception)
                //         {
                //             // Log or handle extraction error
                //         }
                //     }
                // }
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close(false);
                }
                if (wordApp != null)
                {
                    wordApp.Quit(false);
                }
            }

            // Fallback: If no objects were extracted, try Open XML SDK extraction for .docx
            if (!interopSuccess && docxPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("[InteropExtractor] Interop failed or found no objects, using OpenXml fallback...");
                try
                {
                    using (var wordDoc = WordprocessingDocument.Open(docxPath, false))
                    {
                        var embeddedParts = wordDoc.MainDocumentPart.EmbeddedObjectParts.ToList();
                        int xmlCounter = 1;
                        foreach (var part in embeddedParts)
                        {
                            string partExt = ".bin";
                            string partFile = Path.Combine(outputDir, $"Embedded_OpenXml_{xmlCounter}{partExt}");
                            using (var fs = new FileStream(partFile, FileMode.Create, FileAccess.Write))
                            {
                                part.GetStream().CopyTo(fs);
                            }
                            Console.WriteLine($"[InteropExtractor] OpenXml extracted OLE: {partFile}");
                            results.Add(new ExtractedObjectInfo { FilePath = partFile, PageNumber = 0, OleClass = "Package" });
                            xmlCounter++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[InteropExtractor] OpenXml fallback extraction error: {ex.Message}");
                }
            }

            // After extracting .bin OLE packages, extract real files from them using OpenMcdf
            foreach (var obj in results.ToList())
            {
                if (obj.FilePath.EndsWith(".bin", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        var bytes = File.ReadAllBytes(obj.FilePath);
                        var pkg = MsgToPdfConverter.Utils.OlePackageExtractor.ExtractPackage(bytes);
                        if (pkg != null)
                        {
                            string realFilePath = Path.Combine(Path.GetDirectoryName(obj.FilePath), pkg.FileName);
                            File.WriteAllBytes(realFilePath, pkg.Data);
                            Console.WriteLine($"[InteropExtractor] OLE bin extracted: {realFilePath} (from {obj.FilePath})");
                            obj.FilePath = realFilePath;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[InteropExtractor] OLE bin extraction error: {ex.Message}");
                    }
                }
            }

            return results;
        }

        // Attempts to save the OLE object to a file, if possible
        private static void SaveOleObjectToFile(OLEFormat ole, string outFile)
        {
            // Only certain ProgIDs support direct saving; for others, try to save the object if it's a known type
            if (ole.ProgID != null && ole.ProgID.ToLowerInvariant().Contains("pdf"))
            {
                // Embedded PDF: try to save as file
                dynamic obj = ole.Object;
                if (obj != null && obj is MemoryStream)
                {
                    using (var fs = new FileStream(outFile, FileMode.Create, FileAccess.Write))
                    {
                        ((MemoryStream)obj).WriteTo(fs);
                    }
                }
                else
                {
                    // Fallback: try Package extraction (not always possible)
                    ole.Activate();
                }
            }
            else
            {
                // For Excel, Word, etc., try SaveCopyAs if available
                try
                {
                    dynamic obj = ole.Object;
                    if (obj != null && obj.GetType().GetMethod("SaveCopyAs") != null)
                    {
                        obj.SaveCopyAs(outFile);
                    }
                }
                catch { }
            }
        }

        private static string GetExtensionFromProgID(string progId)
        {
            // Map common OLE ProgIDs to file extensions
            if (string.IsNullOrEmpty(progId)) return ".bin";
            progId = progId.ToLowerInvariant();
            if (progId.Contains("pdf")) return ".pdf";
            if (progId.Contains("excel")) return ".xlsx";
            if (progId.Contains("word")) return ".docx";
            if (progId.Contains("package")) return ".bin";
            if (progId.Contains("powerpoint")) return ".pptx";
            return ".bin";
        }
    }
}
