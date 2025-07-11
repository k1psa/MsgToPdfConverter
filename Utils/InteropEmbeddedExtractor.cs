using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Office.Interop.Word;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.Xml.Linq;
using MsgReader;

namespace MsgToPdfConverter.Utils
{
    public class InteropEmbeddedExtractor
    {
        public class ExtractedObjectInfo
        {
            public string FilePath { get; set; }
            public int PageNumber { get; set; } // 1-based page number
            public string OleClass { get; set; }
            public int DocumentOrderIndex { get; set; } // Order in document flow
            public string ExtractedFileName { get; set; }
        }

        /// <summary>
        /// Extracts embedded OLE objects from a .docx file using OpenXml, saving them to the specified output directory.
        /// Uses Interop only to map each object to its page number.
        /// Returns a list of extracted file info, including the page number where each object was found.
        /// </summary>
        public static List<ExtractedObjectInfo> ExtractEmbeddedObjects(string docxPath, string outputDir)
        {
            var results = new List<ExtractedObjectInfo>();
            if (!docxPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                throw new NotSupportedException("Only .docx files are supported for OpenXml extraction.");

            // --- 1. Extract embedded objects using OpenXml in document order ---
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
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[InteropExtractor] OpenXml document order parse error: {ex.Message}");
            }

            // --- 2. Extract objects in order ---
            var relIdToFile = new Dictionary<string, string>();
            var relIdToOleClass = new Dictionary<string, string>();
            int docOrderIndex = 0;
            try
            {
                using (var wordDoc = WordprocessingDocument.Open(docxPath, false))
                {
                    var embeddedParts = wordDoc.MainDocumentPart.EmbeddedObjectParts.ToList();
                    var relIdToPart = new Dictionary<string, EmbeddedObjectPart>();
                    foreach (var rel in wordDoc.MainDocumentPart.Parts)
                    {
                        if (rel.OpenXmlPart is EmbeddedObjectPart objPart)
                        {
                            relIdToPart[rel.RelationshipId] = objPart;
                        }
                    }
                    int xmlCounter = 1;
                    foreach (var relId in orderedRelIds)
                    {
                        if (relIdToPart.TryGetValue(relId, out var part))
                        {
                            string partExt = ".bin";
                            string uniqueSuffix = $"_{relId}_{Guid.NewGuid().ToString("N")}";
                            string partFile = Path.Combine(outputDir, $"Embedded_OpenXml_{xmlCounter}{uniqueSuffix}{partExt}");
                            var appTempDir = Path.Combine(Path.GetTempPath(), "MsgToPdfConverter");
                            string partFileFixed = Path.Combine(appTempDir, $"Embedded_OpenXml_{xmlCounter}{uniqueSuffix}{partExt}");
                            using (var fs = new FileStream(partFileFixed, FileMode.Create, FileAccess.Write))
                            {
                                part.GetStream().CopyTo(fs);
                            }
                            partFile = partFileFixed;
                            relIdToFile[relId] = partFile;
                            relIdToOleClass[relId] = "Package"; // Default, can be improved if needed
                            results.Add(new ExtractedObjectInfo { FilePath = partFile, PageNumber = -1, OleClass = "Package", DocumentOrderIndex = docOrderIndex, ExtractedFileName = Path.GetFileName(partFile) });
                            docOrderIndex++;
                            xmlCounter++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[InteropExtractor] OpenXml extraction error: {ex.Message}");
            }

            // --- 3. Extract real files from .bin using OlePackageExtractor ---
            var binObjs = results.Where(obj => obj.FilePath.EndsWith(".bin", StringComparison.OrdinalIgnoreCase)).ToList();
            var realResults = new List<ExtractedObjectInfo>();
            foreach (var obj in binObjs)
            {
                try
                {
                    var bytes = File.ReadAllBytes(obj.FilePath);
                    var pkg = MsgToPdfConverter.Utils.OlePackageExtractor.ExtractPackage(bytes);
                    if (pkg == null || pkg.Data == null || pkg.Data.Length == 0)
                        continue;
                    string realFileName = Path.GetFileNameWithoutExtension(obj.FilePath) + "_" + Guid.NewGuid().ToString("N") + Path.GetExtension(pkg.FileName ?? "");
                    string realFilePath = Path.Combine(Path.GetDirectoryName(obj.FilePath), realFileName);
                    File.WriteAllBytes(realFilePath, pkg.Data);
                    realResults.Add(new ExtractedObjectInfo {
                        FilePath = realFilePath,
                        PageNumber = -1,
                        OleClass = obj.OleClass,
                        DocumentOrderIndex = obj.DocumentOrderIndex,
                        ExtractedFileName = realFileName
                    });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[InteropExtractor] OLE bin extraction error: {ex.Message}");
                }
            }
            // Replace .bin objects with real extracted files
            results = results.Where(obj => !obj.FilePath.EndsWith(".bin", StringComparison.OrdinalIgnoreCase)).ToList();
            results.AddRange(realResults);
            results = results.OrderBy(r => r.DocumentOrderIndex).ToList();

            // --- 4. Map each extracted object to its page number using Interop (by order) ---
            try
            {
                Application wordApp = null;
                Document doc = null;
                try
                {
                    wordApp = new Application { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
                    doc = wordApp.Documents.Open(docxPath, ReadOnly: true, Visible: false);
                    var inlineShapeMeta = new List<(int Index, string ProgId, int Page)>();
                    for (int idx = 1; idx <= doc.InlineShapes.Count; idx++)
                    {
                        var shape = doc.InlineShapes[idx];
                        string progId = "";
                        try { progId = shape.OLEFormat?.ProgID ?? ""; } catch { }
                        int page = -1;
                        try
                        {
                            page = (int)shape.Range.get_Information(WdInformation.wdActiveEndPageNumber);
                            if (page <= 0)
                                page = (int)shape.Range.get_Information(WdInformation.wdActiveEndAdjustedPageNumber);
                            if (page <= 0)
                            {
                                var startRange = doc.Range(shape.Range.Start, shape.Range.Start);
                                page = (int)startRange.get_Information(WdInformation.wdActiveEndPageNumber);
                            }
                        }
                        catch { }
                        inlineShapeMeta.Add((idx, progId, page));
                    }
                    // Map by order: assign each extracted object to the corresponding InlineShape's page
                    for (int i = 0; i < results.Count && i < inlineShapeMeta.Count; i++)
                    {
                        results[i].PageNumber = inlineShapeMeta[i].Page > 0 ? inlineShapeMeta[i].Page : (i + 1);
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
                Console.WriteLine($"[InteropExtractor] Page mapping failed: {ex.Message}");
                for (int i = 0; i < results.Count; i++)
                {
                    results[i].PageNumber = i + 1;
                }
            }

            // --- 5. Return only one object per logical annex (no duplicates) ---
            // (Assume one per DocumentOrderIndex, as OpenXml order is correct)
            results = results
                .GroupBy(r => r.DocumentOrderIndex)
                .Select(g => g.First())
                .OrderBy(r => r.DocumentOrderIndex)
                .ToList();

            return results;
        }
    }
}
