using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace MsgToPdfConverter.Utils
{
    public class InteropEmbeddedExtractor
    {
        // Constants for safety limits
        private const int MAX_SHAPES_TO_PROCESS = 100;
        private const int MAX_PROCESSING_TIME_MINUTES = 3;
        
        public class ExtractedObjectInfo
        {
            public string FilePath { get; set; }
            public int PageNumber { get; set; } // 1-based page number
            public string OleClass { get; set; }
            public int DocumentOrderIndex { get; set; } // Order in document flow
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
            int docOrderIndex = 0;
            bool interopSuccess = false;
            try
            {
                Console.WriteLine($"[InteropExtractor] ExtractEmbeddedObjects called for: {docxPath}");
                wordApp = new Application { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
                doc = wordApp.Documents.Open(docxPath, ReadOnly: true, Visible: false);

                Console.WriteLine($"[InteropExtractor] InlineShapes count: {doc.InlineShapes.Count}");
                
                // Add limits to prevent excessive processing
                int maxShapesToProcess = Math.Min(doc.InlineShapes.Count, 50); // Limit to 50 shapes
                var startTime = DateTime.Now;
                var maxProcessingTime = TimeSpan.FromMinutes(2); // 2 minute timeout
                int found = 0;
                // InlineShapes (OLE objects, e.g. embedded PDFs, Excels, etc.)
                for (int i = 1; i <= maxShapesToProcess; i++)
                {
                    // Check timeout
                    if (DateTime.Now - startTime > maxProcessingTime)
                    {
                        Console.WriteLine($"[InteropExtractor] Timeout reached, stopping extraction after {i-1} shapes");
                        break;
                    }
                    
                    try
                    {
                        var ish = doc.InlineShapes[i];
                        Console.WriteLine($"[InteropExtractor] Processing InlineShape {i}/{maxShapesToProcess}: Type={ish.Type}, OLE ProgID={ish.OLEFormat?.ProgID}");
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
                                    // Special handling for OLE Package: use DoVerb to activate and try to save
                                    Console.WriteLine($"[InteropExtractor] Attempting to extract OLE Package object");
                                    try
                                    {
                                        // Try different approaches for Package objects
                                        bool saved = false;
                                        
                                        // Method 1: Try to get the object and use reflection carefully
                                        try
                                        {
                                            var obj = ole.Object;
                                            if (obj != null)
                                            {
                                                var type = obj.GetType();
                                                Console.WriteLine($"[InteropExtractor] OLE Package object type: {type.FullName}");
                                                
                                                // Try common save methods
                                                var methods = type.GetMethods().Where(m => 
                                                    m.Name.ToLower().Contains("save") || 
                                                    m.Name.ToLower().Contains("export")).ToArray();
                                                
                                                foreach (var method in methods)
                                                {
                                                    Console.WriteLine($"[InteropExtractor] Found method: {method.Name}");
                                                    try
                                                    {
                                                        if (method.Name == "SaveAs" && method.GetParameters().Length == 1)
                                                        {
                                                            method.Invoke(obj, new object[] { outFile });
                                                            saved = true;
                                                            break;
                                                        }
                                                        else if (method.Name == "SaveToFile" && method.GetParameters().Length == 1)
                                                        {
                                                            method.Invoke(obj, new object[] { outFile });
                                                            saved = true;
                                                            break;
                                                        }
                                                    }
                                                    catch (Exception methodEx)
                                                    {
                                                        Console.WriteLine($"[InteropExtractor] Method {method.Name} failed: {methodEx.Message}");
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception objEx)
                                        {
                                            Console.WriteLine($"[InteropExtractor] Could not access OLE object: {objEx.Message}");
                                        }
                                        
                                        // Method 2: Skip DoVerb activation as it can cause freezes and dialogs
                                        // DoVerb can cause Word to hang or show dialogs, so we skip it
                                        if (!saved)
                                        {
                                            Console.WriteLine($"[InteropExtractor] Skipping DoVerb activation to prevent freezing");
                                        }
                                        
                                        if (!saved)
                                        {
                                            Console.WriteLine($"[InteropExtractor] Could not extract Package object directly - will rely on fallback extraction");
                                            // Don't throw exception - let fallback handle it
                                            continue; // Skip this object for now
                                        }
                                    }
                                    catch (Exception packageEx)
                                    {
                                        Console.WriteLine($"[InteropExtractor] Package extraction failed: {packageEx.Message}");
                                        continue; // Skip this object
                                    }
                                }
                                else
                                {
                                    // Try to save the embedded object if possible
                                    SaveOleObjectToFile(ole, outFile);
                                }
                                // Try multiple robust methods to get page number
                                int page = 0;
                                try
                                {
                                    page = (int)ish.Range.get_Information(WdInformation.wdActiveEndPageNumber);
                                    if (page <= 0)
                                    {
                                        page = (int)ish.Range.get_Information(WdInformation.wdActiveEndAdjustedPageNumber);
                                    }
                                    if (page <= 0)
                                    {
                                        var range = ish.Range;
                                        range.Select();
                                        page = (int)range.get_Information(WdInformation.wdActiveEndPageNumber);
                                    }
                                }
                                catch (Exception pageEx)
                                {
                                    Console.WriteLine($"[InteropExtractor] Could not determine page number: {pageEx.Message}");
                                }
                                if (page <= 0) page = -1;
                                results.Add(new ExtractedObjectInfo { FilePath = outFile, PageNumber = page, OleClass = ole.ProgID, DocumentOrderIndex = docOrderIndex });
                                docOrderIndex++;
                                Console.WriteLine($"[InteropExtractor] Extracted: {outFile} (page {page}, ProgID={ole.ProgID}, Order={docOrderIndex-1})");
                                interopSuccess = true;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[InteropExtractor] Extraction error: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[InteropExtractor] Error processing InlineShape: {ex.Message}");
                    }
                }
                Console.WriteLine($"[InteropExtractor] Embedded OLE InlineShapes found: {found}");

                // Shapes (floating OLE objects) - with safety limits to prevent freezing
                int floatingFound = 0;
                int shapesCount = 0;
                const int MAX_SHAPES_TO_PROCESS = 100; // Limit to prevent freezing
                
                try
                {
                    shapesCount = doc.Shapes.Count;
                    Console.WriteLine($"[InteropExtractor] Total shapes in document: {shapesCount}");
                    
                    if (shapesCount > MAX_SHAPES_TO_PROCESS)
                    {
                        Console.WriteLine($"[InteropExtractor] WARNING: Document has {shapesCount} shapes. Processing only first {MAX_SHAPES_TO_PROCESS} to prevent freezing.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[InteropExtractor] Could not get shapes count: {ex.Message}");
                    shapesCount = 0;
                }
                
                if (shapesCount > 0)
                {
                    int processedShapes = 0;
                    foreach (Shape shape in doc.Shapes)
                    {
                        processedShapes++;
                        if (processedShapes > MAX_SHAPES_TO_PROCESS)
                        {
                            Console.WriteLine($"[InteropExtractor] Stopped processing shapes at limit {MAX_SHAPES_TO_PROCESS}");
                            break;
                        }
                        
                        try
                        {
                            // Add timeout for each shape processing to prevent hanging
                            var shapeTask = System.Threading.Tasks.Task.Run(() =>
                            {
                                // Check if this shape has an OLE format without referencing MsoShapeType
                                bool hasOleFormat = false;
                                try
                                {
                                    var oleFormat = shape.OLEFormat;
                                    hasOleFormat = (oleFormat != null);
                                }
                                catch
                                {
                                    hasOleFormat = false;
                                }
                                
                                if (hasOleFormat)
                                {
                                    var ole = shape.OLEFormat;
                                    string ext = GetExtensionFromProgID(ole.ProgID);
                                    string outFile = Path.Combine(outputDir, $"Embedded_Floating_{counter}{ext}");
                                    counter++;
                                    SaveOleObjectToFile(ole, outFile);
                                    int page = -1;
                                    try
                                    {
                                        page = (int)shape.Anchor.get_Information(WdInformation.wdActiveEndPageNumber);
                                    }
                                    catch (Exception pageEx)
                                    {
                                        Console.WriteLine($"[InteropExtractor] Could not determine page number for floating shape: {pageEx.Message}");
                                    }
                                    if (page <= 0) page = -1;
                                    results.Add(new ExtractedObjectInfo { FilePath = outFile, PageNumber = page, OleClass = ole.ProgID, DocumentOrderIndex = docOrderIndex });
                                    docOrderIndex++;
                                    Console.WriteLine($"[InteropExtractor] Extracted floating OLE: {outFile} (page {page}, ProgID={ole.ProgID}, Order={docOrderIndex-1})");
                                    floatingFound++;
                                }
                            });
                            
                            // Wait for task with timeout (5 seconds per shape)
                            if (!shapeTask.Wait(5000))
                            {
                                Console.WriteLine($"[InteropExtractor] Shape processing timed out, skipping shape {processedShapes}");
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[InteropExtractor] Error extracting floating OLE from shape {processedShapes}: {ex.Message}");
                        }
                    }
                }
                Console.WriteLine($"[InteropExtractor] Embedded OLE floating Shapes found: {floatingFound}");
            }
            finally
            {
                if (doc != null)
                {
                    try 
                    {
                        doc.Close(false);
                        doc = null;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[InteropExtractor] Error closing document: {ex.Message}");
                    }
                }
                if (wordApp != null)
                {
                    try
                    {
                        wordApp.Quit(false);
                        wordApp = null;
                        // Give Word time to fully close and release file locks
                        System.Threading.Thread.Sleep(2000);
                        // Force garbage collection to ensure COM objects are released
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[InteropExtractor] Error closing Word application: {ex.Message}");
                    }
                }
            }

            // Fallback: If no objects were extracted, try Open XML SDK extraction for .docx
            if (!interopSuccess && docxPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("[InteropExtractor] Interop failed or found no objects, using OpenXml fallback...");
                Console.WriteLine("[InteropExtractor] Attempting to determine document order of embedded objects via OpenXml.");

                // --- Improved: Parse document.xml for embedded object order, including nested r:id ---
                List<string> orderedRelIds = new List<string>();
                try
                {
                    using (var wordDoc = WordprocessingDocument.Open(docxPath, false))
                    {
                        var mainPart = wordDoc.MainDocumentPart;
                        if (mainPart != null)
                        {
                            var xdoc = System.Xml.Linq.XDocument.Load(mainPart.GetStream());
                            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                            XNamespace v = "urn:schemas-microsoft-com:vml";
                            XNamespace o = "urn:schemas-microsoft-com:office:office";
                            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

                            // Find all <w:object> elements in order
                            var objectElements = xdoc.Descendants(w + "object").ToList();
                            foreach (var objElem in objectElements)
                            {
                                // 1. Check for <w:oleObject r:id="..."> descendant
                                var oleElem = objElem.Descendants(w + "oleObject").FirstOrDefault();
                                if (oleElem != null)
                                {
                                    var relId = oleElem.Attribute(r + "id")?.Value;
                                    if (!string.IsNullOrEmpty(relId))
                                    {
                                        orderedRelIds.Add(relId);
                                        continue;
                                    }
                                }
                                // 2. Check for <o:OLEObject r:id="..."> descendant (sometimes used)
                                var oElem = objElem.Descendants(o + "OLEObject").FirstOrDefault();
                                if (oElem != null)
                                {
                                    var relId = oElem.Attribute(r + "id")?.Value;
                                    if (!string.IsNullOrEmpty(relId))
                                    {
                                        orderedRelIds.Add(relId);
                                        continue;
                                    }
                                }
                                // 3. Check for r:id directly on <w:object>
                                var directId = objElem.Attribute(r + "id")?.Value;
                                if (!string.IsNullOrEmpty(directId))
                                {
                                    orderedRelIds.Add(directId);
                                    continue;
                                }
                            }
                            // Also check for <w:oleObject> outside <w:object>
                            var looseOleElements = xdoc.Descendants(w + "oleObject").ToList();
                            foreach (var elem in looseOleElements)
                            {
                                var relId = elem.Attribute(r + "id")?.Value;
                                if (!string.IsNullOrEmpty(relId) && !orderedRelIds.Contains(relId))
                                    orderedRelIds.Add(relId);
                            }
                            // Also check for <w:altChunk> (for embedded files)
                            var altChunkElements = xdoc.Descendants(w + "altChunk").ToList();
                            foreach (var elem in altChunkElements)
                            {
                                var relId = elem.Attribute(r + "id")?.Value;
                                if (!string.IsNullOrEmpty(relId) && !orderedRelIds.Contains(relId))
                                    orderedRelIds.Add(relId);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[InteropExtractor] OpenXml document order parse error: {ex.Message}");
                }
                // --- End Improved ---

                // Retry mechanism for file lock issues
                int retries = 3;
                // Place usedRelIds declaration here so it is in scope for both loops
                var usedRelIds = new HashSet<string>();
                for (int i = 0; i < retries; i++)
                {
                    try
                    {
                        using (var wordDoc = WordprocessingDocument.Open(docxPath, false))
                        {
                            var embeddedParts = wordDoc.MainDocumentPart.EmbeddedObjectParts.ToList();
                            Console.WriteLine($"[InteropExtractor] OpenXml found {embeddedParts.Count} embedded object parts");

                            // Build relId -> part mapping
                            var relIdToPart = new Dictionary<string, EmbeddedObjectPart>();
                            foreach (var rel in wordDoc.MainDocumentPart.Parts)
                            {
                                if (rel.OpenXmlPart is EmbeddedObjectPart objPart)
                                {
                                    relIdToPart[rel.RelationshipId] = objPart;
                                }
                            }

                            int xmlCounter = 1;
                            // Insert in document order first
                            foreach (var relId in orderedRelIds)
                            {
                                if (relIdToPart.TryGetValue(relId, out var part))
                                {
                                    string partExt = ".bin";
                                    string partFile = Path.Combine(outputDir, $"Embedded_OpenXml_{xmlCounter}{partExt}");
                                    using (var fs = new FileStream(partFile, FileMode.Create, FileAccess.Write))
                                    {
                                        part.GetStream().CopyTo(fs);
                                    }
                                    Console.WriteLine($"[InteropExtractor] OpenXml extracted OLE: {partFile} (relId={relId}, Order={docOrderIndex})");
                                    results.Add(new ExtractedObjectInfo { FilePath = partFile, PageNumber = -1, OleClass = "Package", DocumentOrderIndex = docOrderIndex });
                                    docOrderIndex++;
                                    xmlCounter++;
                                    usedRelIds.Add(relId); // Mark as used here to prevent duplicates
                                }
                            }
                            // Add any remaining parts not referenced in document order (rare)
                            foreach (var part in embeddedParts)
                            {
                                var rel = wordDoc.MainDocumentPart.Parts.FirstOrDefault(p => p.OpenXmlPart == part);
                                if (rel != null && !usedRelIds.Contains(rel.RelationshipId))
                                {
                                    string partExt = ".bin";
                                    string partFile = Path.Combine(outputDir, $"Embedded_OpenXml_{xmlCounter}{partExt}");
                                    using (var fs = new FileStream(partFile, FileMode.Create, FileAccess.Write))
                                    {
                                        part.GetStream().CopyTo(fs);
                                    }
                                    Console.WriteLine($"[InteropExtractor] OpenXml extracted OLE (unreferenced): {partFile} (relId={rel.RelationshipId}, Order={docOrderIndex})");
                                    results.Add(new ExtractedObjectInfo { FilePath = partFile, PageNumber = -1, OleClass = "Package", DocumentOrderIndex = docOrderIndex });
                                    docOrderIndex++;
                                    xmlCounter++;
                                    usedRelIds.Add(rel.RelationshipId); // Mark as used so no duplicates
                                }
                            }
                        }
                        break; // Success, exit retry loop
                    }
                    catch (IOException ex) when (ex.Message.Contains("being used by another process"))
                    {
                        Console.WriteLine($"[InteropExtractor] File locked (attempt {i + 1}/{retries}): {ex.Message}");
                        if (i < retries - 1)
                        {
                            // Wait longer before retrying
                            System.Threading.Thread.Sleep(3000 * (i + 1));
                        }
                        else
                        {
                            Console.WriteLine($"[InteropExtractor] OpenXml fallback failed after {retries} attempts due to file lock");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[InteropExtractor] OpenXml fallback extraction error: {ex.Message}");
                        break; // Non-retryable error
                    }
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

            // --- Find ACTUAL page numbers using Word Interop if fallback was used ---
            if (results.Count > 0 && results.All(o => o.PageNumber == -1))
            {
                Console.WriteLine($"[InteropExtractor] Finding actual page numbers for {results.Count} objects using Word Interop");
                
                Application pageWordApp = null;
                Document pageDoc = null;
                try
                {
                    pageWordApp = new Application { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
                    pageDoc = pageWordApp.Documents.Open(docxPath, ReadOnly: true, Visible: false);
                    
                    // Get the actual page numbers from the InlineShapes we found earlier
                    Console.WriteLine($"[InteropExtractor] Document has {pageDoc.InlineShapes.Count} InlineShapes");
                    
                    for (int i = 0; i < results.Count && i < pageDoc.InlineShapes.Count; i++)
                    {
                        try
                        {
                            var shape = pageDoc.InlineShapes[i + 1]; // 1-based indexing
                            if (shape.Type == WdInlineShapeType.wdInlineShapeEmbeddedOLEObject)
                            {
                                int actualPage = -1;
                                
                                // Try multiple methods to get the page number
                                try
                                {
                                    actualPage = (int)shape.Range.get_Information(WdInformation.wdActiveEndPageNumber);
                                    if (actualPage <= 0)
                                    {
                                        actualPage = (int)shape.Range.get_Information(WdInformation.wdActiveEndAdjustedPageNumber);
                                    }
                                    if (actualPage <= 0)
                                    {
                                        // Try using the range start
                                        var startRange = pageDoc.Range(shape.Range.Start, shape.Range.Start);
                                        actualPage = (int)startRange.get_Information(WdInformation.wdActiveEndPageNumber);
                                    }
                                }
                                catch (Exception pageEx)
                                {
                                    Console.WriteLine($"[InteropExtractor] Could not get page for shape {i+1}: {pageEx.Message}");
                                }
                                
                                if (actualPage > 0)
                                {
                                    results[i].PageNumber = actualPage;
                                    Console.WriteLine($"[InteropExtractor] Object {i+1} found on actual page {actualPage}");
                                }
                                else
                                {
                                    // Fallback: use simple sequential assignment
                                    results[i].PageNumber = i + 1;
                                    Console.WriteLine($"[InteropExtractor] Object {i+1} assigned to fallback page {i+1}");
                                }
                            }
                        }
                        catch (Exception shapeEx)
                        {
                            Console.WriteLine($"[InteropExtractor] Error processing shape {i+1}: {shapeEx.Message}");
                            results[i].PageNumber = i + 1; // Fallback
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[InteropExtractor] Page detection failed: {ex.Message}. Using simple assignment.");
                    
                    // Final fallback: simple sequential assignment
                    for (int i = 0; i < results.Count; i++)
                    {
                        results[i].PageNumber = i + 1;
                        Console.WriteLine($"[InteropExtractor] Object {i+1} assigned to fallback page {i+1}");
                    }
                }
                finally
                {
                    if (pageDoc != null)
                    {
                        try { pageDoc.Close(false); } catch { }
                    }
                    if (pageWordApp != null)
                    {
                        try { pageWordApp.Quit(false); } catch { }
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
        
        /// <summary>
        /// Gets the page number for an InlineShape using multiple robust methods
        /// </summary>
        private static int GetPageNumberForInlineShape(InlineShape inlineShape, Document doc)
        {
            int page = 0;
            
            try
            {
                Console.WriteLine($"[InteropExtractor] Attempting to get page number for InlineShape...");
                
                // Method 1: Direct range information (most reliable)
                var range = inlineShape.Range;
                page = (int)range.get_Information(WdInformation.wdActiveEndPageNumber);
                Console.WriteLine($"[InteropExtractor] Method 1 (Range.Information): page = {page}");
                
                if (page <= 0)
                {
                    // Method 2: Try with adjusted page number
                    page = (int)range.get_Information(WdInformation.wdActiveEndAdjustedPageNumber);
                    Console.WriteLine($"[InteropExtractor] Method 2 (AdjustedPageNumber): page = {page}");
                }
                
                if (page <= 0)
                {
                    // Method 3: Try by getting the range start position
                    var startRange = doc.Range(range.Start, range.Start);
                    page = (int)startRange.get_Information(WdInformation.wdActiveEndPageNumber);
                    Console.WriteLine($"[InteropExtractor] Method 3 (StartRange): page = {page}");
                }
                
                if (page <= 0)
                {
                    // Method 4: Try by calculating page from character position
                    page = CalculatePageFromPosition(range.Start, doc);
                    Console.WriteLine($"[InteropExtractor] Method 4 (CalculateFromPosition): page = {page}");
                }
                
                if (page <= 0)
                {
                    // Method 5: Fallback - estimate based on document structure
                    page = EstimatePageFromDocumentStructure(range, doc);
                    Console.WriteLine($"[InteropExtractor] Method 5 (EstimateFromStructure): page = {page}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[InteropExtractor] Error getting page number: {ex.Message}");
                page = 0;
            }
            
            Console.WriteLine($"[InteropExtractor] Final page number: {page}");
            return page;
        }
        
        /// <summary>
        /// Calculates page number from character position in document
        /// </summary>
        private static int CalculatePageFromPosition(int position, Document doc)
        {
            try
            {
                // Get total pages and characters to estimate
                int totalPages = doc.ComputeStatistics(WdStatistic.wdStatisticPages);
                int totalChars = doc.ComputeStatistics(WdStatistic.wdStatisticCharacters);
                if (totalChars > 0 && totalPages > 0)
                {
                    // Simple estimation: position / (total chars / total pages)
                    int estimatedPage = (int)Math.Ceiling((double)position / totalChars * totalPages);
                    return Math.Max(1, Math.Min(estimatedPage, totalPages));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[InteropExtractor] Error calculating page from position: {ex.Message}");
            }
            return 0;
        }
        
        /// <summary>
        /// Estimates the page number based on document structure (headings, footers, etc.)
        /// </summary>
        private static int EstimatePageFromDocumentStructure(Range range, Document doc)
        {
            int estimatedPage = 0;
            
            try
            {
                // Heuristic: check preceding and following content for page breaks or section breaks
                var prevRange = range.Previous();
                var nextRange = range.Next();
                
                // Check for page break character (manual page break)
                if (prevRange != null && prevRange.Text.Trim() == "\f")
                {
                    estimatedPage--;
                }
                if (nextRange != null && nextRange.Text.Trim() == "\f")
                {
                    estimatedPage++;
                }
                
                // Check for section breaks (next page)
                if (nextRange != null && nextRange.Sections.Count > 0)
                {
                    estimatedPage++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[InteropExtractor] Error estimating page from structure: {ex.Message}");
            }
            
            return estimatedPage;
        }
    }
}
