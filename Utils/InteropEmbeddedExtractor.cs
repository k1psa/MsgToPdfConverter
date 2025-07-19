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
        public string ProgId { get; set; } // For better matching
        public int ParagraphIndex { get; set; } // For position-based matching
        public int RunIndex { get; set; } // For position-based matching
        public int CharPosition { get; set; } // Character position in document
    // ...existing code...
}
// ...existing code...


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
            List<(string relId, int paraIdx, int runIdx, int charPos)> orderedRelIdsWithPos = new List<(string, int, int, int)>();
            try
            {
                using (var wordDoc = WordprocessingDocument.Open(docxPath, false))
                {
                    var mainPart = wordDoc.MainDocumentPart;
                    if (mainPart != null)
                    {
                        var body = mainPart.Document.Body;
                        int paraIdx = 0;
                        int charPos = 0;
                        foreach (var para in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                        {
                            int runIdx = 0;
                            foreach (var run in para.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
                            {
                                foreach (var obj in run.Elements().Where(e => e.LocalName == "object"))
                                {
                                    var oleObj = obj.Elements().FirstOrDefault(e => e.LocalName == "OLEObject");
                                    if (oleObj != null)
                                    {
                                        var relIdAttr = oleObj.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                                        if (relIdAttr != null)
                                        {
                                            string relId = relIdAttr.Value;
                                            orderedRelIdsWithPos.Add((relId, paraIdx, runIdx, charPos));
                                            continue;
                                        }
                                    }
                                }
                                charPos++;
                                runIdx++;
                            }
                            charPos++;
                            paraIdx++;
                        }

                        // --- Explicitly extract all .xlsx files from /word/embeddings ---
                        int nextOrderIndex = results.Count > 0 ? results.Max(r => r.DocumentOrderIndex) + 1 : 0;
                        foreach (var rel in mainPart.Parts)
                        {
                            var partUri = rel.OpenXmlPart.Uri.ToString();
                            if (partUri.StartsWith("/word/embeddings/") && partUri.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                            {
                                string fileName = Path.GetFileName(partUri);
                                string outPath = Path.Combine(outputDir, fileName);
                                // Only extract if not already present
                                if (!File.Exists(outPath))
                                {
                                    using (var fs = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                                    {
                                        rel.OpenXmlPart.GetStream().CopyTo(fs);
                                    }
                                    Console.WriteLine($"[InteropExtractor] Extracted orphaned Excel: {outPath}");
                                }
                                // Prevent duplicate: only add orphaned Excel if no in-flow .xlsx with same file name exists
                                bool hasInFlow = results.Any(r =>
                                    Path.GetFileName(r.FilePath).Equals(fileName, StringComparison.OrdinalIgnoreCase)
                                    && r.ParagraphIndex >= 0
                                    && Path.GetExtension(r.FilePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase));
                                if (!hasInFlow && !results.Any(r => Path.GetFileName(r.FilePath).Equals(fileName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    results.Add(new ExtractedObjectInfo {
                                        FilePath = outPath,
                                        PageNumber = -1,
                                        OleClass = "Excel.Sheet",
                                        DocumentOrderIndex = nextOrderIndex++,
                                        ExtractedFileName = fileName,
                                        ProgId = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        ParagraphIndex = -1,
                                        RunIndex = -1,
                                        CharPosition = -1
                                    });
                                    Console.WriteLine($"[InteropExtractor] Added orphaned Excel to results: {fileName}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[InteropExtractor] OpenXml document order parse error: {ex.Message}");
            }

            // --- 2. Extract objects in order ---
            // Print extracted objects for debug
            Console.WriteLine("[InteropExtractor] Extracted OpenXml objects:");
            foreach (var tuple in orderedRelIdsWithPos)
            {
                Console.WriteLine($"  relId={tuple.relId}, paraIdx={tuple.paraIdx}, runIdx={tuple.runIdx}, charPos={tuple.charPos}");
            }

            var relIdToFile = new Dictionary<string, string>();
            var relIdToOleClass = new Dictionary<string, string>();
            var relIdToProgId = new Dictionary<string, string>();
            int docOrderIndex = 0;
            try
            {
                using (var wordDoc = WordprocessingDocument.Open(docxPath, false))
                {
                    // Map relId to OpenXmlPart (can be EmbeddedObjectPart or EmbeddedPackagePart)
                    var relIdToPart = new Dictionary<string, OpenXmlPart>();
                    foreach (var rel in wordDoc.MainDocumentPart.Parts)
                    {
                        if (rel.OpenXmlPart is EmbeddedObjectPart objPart)
                        {
                            relIdToPart[rel.RelationshipId] = objPart;
                            string progId = objPart.ContentType;
                            relIdToProgId[rel.RelationshipId] = progId;
                        }
                        else if (rel.OpenXmlPart is EmbeddedPackagePart pkgPart)
                        {
                            relIdToPart[rel.RelationshipId] = pkgPart;
                            string progId = pkgPart.ContentType;
                            relIdToProgId[rel.RelationshipId] = progId;
                        }
                    }
                    int xmlCounter = 1;
                    foreach (var tuple in orderedRelIdsWithPos)
                    {
                        var relId = tuple.relId;
                        var paraIdx = tuple.paraIdx;
                        var runIdx = tuple.runIdx;
                        var charPos = tuple.charPos;
                        if (relIdToPart.TryGetValue(relId, out var part))
                        {
                            string partExt = ".bin";
                            string progId = relIdToProgId.ContainsKey(relId) ? relIdToProgId[relId] : null;
                            // Direct extraction for Word/Excel files
                            if (progId == "application/msword") partExt = ".doc";
                            else if (progId == "application/vnd.openxmlformats-officedocument.wordprocessingml.document") partExt = ".docx";
                            else if (progId == "application/vnd.ms-excel") partExt = ".xls";
                            else if (progId == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet") partExt = ".xlsx";
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
                            relIdToOleClass[relId] = "Package";
                            results.Add(new ExtractedObjectInfo {
                                FilePath = partFile,
                                PageNumber = -1,
                                OleClass = "Package",
                                DocumentOrderIndex = docOrderIndex,
                                ExtractedFileName = Path.GetFileName(partFile),
                                ProgId = progId,
                                ParagraphIndex = paraIdx,
                                RunIndex = runIdx,
                                CharPosition = charPos
                            });
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


            // --- 2b. Explicitly extract all .xlsx files from /word/embeddings ---
            try
            {
                using (var wordDoc = WordprocessingDocument.Open(docxPath, false))
                {
                    int nextOrderIndex = results.Count > 0 ? results.Max(r => r.DocumentOrderIndex) + 1 : 0;
                    foreach (var rel in wordDoc.MainDocumentPart.Parts)
                    {
                        var partUri = rel.OpenXmlPart.Uri.ToString();
                        if (partUri.StartsWith("/word/embeddings/") && partUri.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            string fileName = Path.GetFileName(partUri);
                            string outPath = Path.Combine(outputDir, fileName);
                            // Only extract if not already present
                            if (!File.Exists(outPath))
                            {
                                using (var fs = new FileStream(outPath, FileMode.Create, FileAccess.Write))
                                {
                                    rel.OpenXmlPart.GetStream().CopyTo(fs);
                                }
                                Console.WriteLine($"[InteropExtractor] Extracted orphaned Excel: {outPath}");
                            }
                            // Only add to results if not already present
                            if (!results.Any(r => Path.GetFileName(r.FilePath).Equals(fileName, StringComparison.OrdinalIgnoreCase)))
                            {
                                results.Add(new ExtractedObjectInfo {
                                    FilePath = outPath,
                                    PageNumber = -1,
                                    OleClass = "Excel.Sheet",
                                    DocumentOrderIndex = nextOrderIndex++,
                                    ExtractedFileName = fileName,
                                    ProgId = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    ParagraphIndex = -1,
                                    RunIndex = -1,
                                    CharPosition = -1
                                });
                                Console.WriteLine($"[InteropExtractor] Added orphaned Excel to results: {fileName}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[InteropExtractor] Error extracting orphaned Excel: {ex.Message}");
            }

            // After results are populated
            Console.WriteLine("[InteropExtractor] Results after OpenXml extraction:");
            foreach (var obj in results)
            {
                Console.WriteLine($"  [{obj.DocumentOrderIndex}] File={obj.ExtractedFileName}, ProgId={obj.ProgId}, ParaIdx={obj.ParagraphIndex}, RunIdx={obj.RunIndex}, CharPos={obj.CharPosition}");
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
                    // Only keep valid extracted files (non-empty, known extension)
                    if (pkg == null || pkg.Data == null || pkg.Data.Length == 0)
                        continue;
                    string ext = Path.GetExtension(pkg.FileName ?? "").ToLowerInvariant();
                    bool isGenericExt = ext == "" || ext == ".bin" || ext == ".dat" || ext == ".tmp";
                    // Robust PDF detection: check for PDF header
                    bool isPdf = pkg.Data.Length > 4 && System.Text.Encoding.ASCII.GetString(pkg.Data, 0, 5) == "%PDF-";
                    if (isPdf)
                    {
                        ext = ".pdf";
                    }
                    // Only skip if not PDF and extension is generic
                    if (!isPdf && isGenericExt)
                        continue;
                    string realFileName = Path.GetFileNameWithoutExtension(obj.FilePath) + "_" + Guid.NewGuid().ToString("N") + ext;
                    string realFilePath = Path.Combine(Path.GetDirectoryName(obj.FilePath), realFileName);
                    File.WriteAllBytes(realFilePath, pkg.Data);
                    realResults.Add(new ExtractedObjectInfo {
                        FilePath = realFilePath,
                        PageNumber = -1,
                        OleClass = obj.OleClass,
                        DocumentOrderIndex = obj.DocumentOrderIndex,
                        ExtractedFileName = realFileName,
                        ProgId = obj.ProgId,
                        ParagraphIndex = obj.ParagraphIndex,
                        RunIndex = obj.RunIndex,
                        CharPosition = obj.CharPosition
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
            // Filter out bogus OpenXml objects (no valid extracted file)
            results = results.Where(obj => !string.IsNullOrWhiteSpace(obj.ExtractedFileName) && obj.ExtractedFileName != null && obj.ExtractedFileName != "" && Path.GetExtension(obj.ExtractedFileName).ToLowerInvariant() != ".bin" && Path.GetExtension(obj.ExtractedFileName).ToLowerInvariant() != ".dat" && Path.GetExtension(obj.ExtractedFileName).ToLowerInvariant() != ".tmp").OrderBy(r => r.DocumentOrderIndex).ToList();

            // --- 4. Map each extracted object to its page number using Interop (by order) ---
            try
            {
                Application wordApp = null;
                Document doc = null;
                try
                {
                    wordApp = new Application { Visible = false, DisplayAlerts = WdAlertLevel.wdAlertsNone };
                    doc = wordApp.Documents.Open(docxPath, ReadOnly: true, Visible: false);
                    var inlineShapeMeta = new List<(int Index, string ProgId, int Page, int CharPosition, string FileName)>();
                    Console.WriteLine("[InteropExtractor] InlineShapes from Interop:");
                    for (int idx = 1; idx <= doc.InlineShapes.Count; idx++)
                    {
                        var shape = doc.InlineShapes[idx];
                        string progId = "";
                        string fileName = "";
                        int charPos = -1;
                        try { progId = shape.OLEFormat?.ProgID ?? ""; } catch { }
                        try { fileName = shape.OLEFormat?.Object?.GetType().GetProperty("Name")?.GetValue(shape.OLEFormat.Object)?.ToString() ?? ""; } catch { }
                        try { charPos = shape.Range.Start; } catch { }
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
                        inlineShapeMeta.Add((idx, progId, page, charPos, fileName));
                        Console.WriteLine($"  [{idx}] File={fileName}, ProgId={progId}, CharPos={charPos}, Page={page}");
                    }
                    // Filter out bogus InlineShapes (empty ProgId or file name)
                    var validInlineShapes = inlineShapeMeta
                        .Where(meta => !string.IsNullOrWhiteSpace(meta.ProgId) && meta.ProgId != "" && meta.ProgId != null)
                        .ToList();
                    Console.WriteLine($"[InteropExtractor] Valid InlineShapes for mapping: {validInlineShapes.Count}");
                    foreach (var meta in validInlineShapes)
                    {
                        Console.WriteLine($"  [Valid] Index={meta.Index}, ProgId={meta.ProgId}, FileName={meta.FileName}, Page={meta.Page}");
                    }
                    // Mapping debug
                    Console.WriteLine("[InteropExtractor] Mapping results:");
                    var usedInlineShapes = new HashSet<int>();

                    // Split results into in-flow and orphaned
                    var inFlowObjs = results.Where(obj => obj.ParagraphIndex >= 0).ToList();
                    var orphanedObjs = results.Where(obj => obj.ParagraphIndex < 0).ToList();
                    // 1. Map in-flow objects first
                    foreach (var obj in inFlowObjs)
                    {
                        int matchedIdx = -1;
                        string ext = Path.GetExtension(obj.ExtractedFileName ?? "").ToLowerInvariant();
                        for (int j = 0; j < validInlineShapes.Count; j++)
                        {
                            if (usedInlineShapes.Contains(j)) continue;
                            var meta = validInlineShapes[j];
                            // PDF
                            if (ext == ".pdf" && meta.ProgId == "Package") { matchedIdx = j; break; }
                            // MSG
                            if (ext == ".msg" && meta.ProgId == "Package") { matchedIdx = j; break; }
                            // Word
                            if ((ext == ".doc" || ext == ".docx") && meta.ProgId.StartsWith("Word.Document")) { matchedIdx = j; break; }
                            // Excel
                            if ((ext == ".xls" || ext == ".xlsx") && meta.ProgId.StartsWith("Excel.Sheet")) { matchedIdx = j; break; }
                        }
                        if (matchedIdx == -1)
                        {
                            for (int j = 0; j < validInlineShapes.Count; j++)
                            {
                                if (!usedInlineShapes.Contains(j))
                                {
                                    matchedIdx = j;
                                    break;
                                }
                            }
                        }
                        if (matchedIdx != -1)
                        {
                            var meta = validInlineShapes[matchedIdx];
                            obj.PageNumber = meta.Page > 0 ? meta.Page : (meta.Index);
                            usedInlineShapes.Add(matchedIdx);
                            Console.WriteLine($"  [{obj.DocumentOrderIndex}] File={obj.ExtractedFileName} mapped to InlineShape[{meta.Index}] Page={obj.PageNumber}");
                        }
                        else
                        {
                            obj.PageNumber = 1;
                            Console.WriteLine($"  [{obj.DocumentOrderIndex}] File={obj.ExtractedFileName} could not be mapped to any InlineShape");
                        }
                    }
                    // 2. Map orphaned objects to any remaining InlineShapes
                    foreach (var obj in orphanedObjs)
                    {
                        int matchedIdx = -1;
                        string ext = Path.GetExtension(obj.ExtractedFileName ?? "").ToLowerInvariant();
                        for (int j = 0; j < validInlineShapes.Count; j++)
                        {
                            if (usedInlineShapes.Contains(j)) continue;
                            var meta = validInlineShapes[j];
                            // Excel (for orphaned)
                            if ((ext == ".xls" || ext == ".xlsx") && meta.ProgId.StartsWith("Excel.Sheet")) { matchedIdx = j; break; }
                        }
                        if (matchedIdx == -1)
                        {
                            for (int j = 0; j < validInlineShapes.Count; j++)
                            {
                                if (!usedInlineShapes.Contains(j))
                                {
                                    matchedIdx = j;
                                    break;
                                }
                            }
                        }
                        if (matchedIdx != -1)
                        {
                            var meta = validInlineShapes[matchedIdx];
                            obj.PageNumber = meta.Page > 0 ? meta.Page : (meta.Index);
                            usedInlineShapes.Add(matchedIdx);
                            Console.WriteLine($"  [{obj.DocumentOrderIndex}] File={obj.ExtractedFileName} mapped to InlineShape[{meta.Index}] Page={obj.PageNumber}");
                        }
                        else
                        {
                            obj.PageNumber = 1;
                            Console.WriteLine($"  [{obj.DocumentOrderIndex}] File={obj.ExtractedFileName} could not be mapped to any InlineShape");
                        }
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

            // --- 5. Remove duplicates by FilePath (if any), but keep all unique extracted objects ---

            // Remove orphaned .xlsx if an in-flow .xlsx exists (even if file names differ)

            // --- Improved deduplication: only remove true duplicates (same file path and same content hash) ---
            // Compute hash for each file
            foreach (var obj in results)
            {
                try
                {
                    if (!string.IsNullOrEmpty(obj.FilePath) && File.Exists(obj.FilePath))
                    {
                        using (var stream = File.OpenRead(obj.FilePath))
                        {
                            using (var sha = System.Security.Cryptography.SHA256.Create())
                            {
                                var hash = sha.ComputeHash(stream);
                                obj.ExtractedFileName += "|HASH:" + BitConverter.ToString(hash).Replace("-", "");
                            }
                        }
                    }
                }
                catch { }
            }

            // Remove only true duplicates (same file path and same content hash)
            results = results

                .GroupBy(r => r.FilePath + "|" + (r.ExtractedFileName ?? ""), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderBy(r => r.DocumentOrderIndex)
                .ToList();

            // Remove orphaned .xlsx if any in-flow .xlsx exists (even if file names or hashes differ)
            bool hasInFlowXlsx = results.Any(r => r.ParagraphIndex >= 0 && Path.GetExtension(r.FilePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase));
            if (hasInFlowXlsx) {
                results = results.Where(r =>
                    !(r.ParagraphIndex < 0 && Path.GetExtension(r.FilePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                ).ToList();
            }

            return results;
        }
    }
}
