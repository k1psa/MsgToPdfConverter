using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
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
        // Static progress callback for embedding operations
        private static Action s_currentProgressCallback = null;
        
        /// <summary>
        /// Sets the progress callback for current embedding operations
        /// </summary>
        public static void SetProgressCallback(Action progressCallback)
        {
            s_currentProgressCallback = progressCallback;
        }
        private static EmailConverterService _emailConverterService = new EmailConverterService();
        // --- Add static AttachmentService for 7z/zip recursive extraction ---
        private static AttachmentService _attachmentService = new AttachmentService(
            (path, text, _) => PdfService.AddHeaderPdf(path, text),
            OfficeConversionService.TryConvertOfficeToPdf,
            PdfAppendTest.AppendPdfs,
            _emailConverterService
        );

        /// <summary>
        /// Inserts extracted embedded files into the main PDF document after the pages where they were found
        /// </summary>
        /// <param name="mainPdfPath">Path to the main PDF file</param>
        /// <param name="extractedObjects">List of extracted embedded objects with page numbers</param>
        /// <param name="outputPdfPath">Path for the output PDF with embedded files inserted</param>
        public static void InsertEmbeddedFiles(string mainPdfPath, List<InteropEmbeddedExtractor.ExtractedObjectInfo> extractedObjects, string outputPdfPath, Action progressTick = null)
        {
            if (!File.Exists(mainPdfPath))
            {
#if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Main PDF not found: {mainPdfPath}");
#endif
                return;
            }


#if DEBUG
            DebugLogger.Log($"[PDF-INSERT] Starting insertion process. Main PDF: {mainPdfPath}");
#endif

#if DEBUG
            DebugLogger.Log($"[PDF-INSERT] Output PDF: {outputPdfPath}");
#endif

#if DEBUG
            DebugLogger.Log($"[PDF-INSERT] Found {extractedObjects?.Count ?? 0} extracted objects");
#endif

            var validObjects = new List<InteropEmbeddedExtractor.ExtractedObjectInfo>();
            foreach (var obj in extractedObjects)
            {
                if (!File.Exists(obj.FilePath))
                {
#if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] Warning: Extracted file not found: {obj.FilePath}");
#endif
                    continue;
                }
                var fileInfo = new FileInfo(obj.FilePath);
                if (fileInfo.Length == 0)
                {
#if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] Warning: Extracted file is empty: {obj.FilePath}");
#endif
                    continue;
                }
                validObjects.Add(obj);
            }


#if DEBUG
            DebugLogger.Log($"[PDF-INSERT] {validObjects.Count} valid objects to insert");
#endif

            if (validObjects.Count == 0)
            {

#if DEBUG
            DebugLogger.Log("[PDF-INSERT] No valid embedded files to insert, copying main PDF");
#endif
                File.Copy(mainPdfPath, outputPdfPath, true);
                return;
            }

            // --- Sequential Insertion Logic ---
            int mainPageCount = 0;
            using (var mainPdf = new PdfDocument(new PdfReader(mainPdfPath)))
            {
                mainPageCount = mainPdf.GetNumberOfPages();
            }



            // --- Strict Inline Shape Index Mapping ---
            var mappedObjects = new List<(InteropEmbeddedExtractor.ExtractedObjectInfo obj, int anchorPage, string reason)>();
            var unmappedObjects = new List<(InteropEmbeddedExtractor.ExtractedObjectInfo obj, string reason)>();


            // --- Fix: Map embedded files to inline shapes/pages by index ---
            // Get all valid inline shape pages (sorted by PageNumber ascending)
            var inlineShapePages = extractedObjects
                .Where(o => o.PageNumber > 0)
                .OrderBy(o => o.PageNumber)
                .Select(o => o.PageNumber)
                .Distinct()
                .ToList();

            // Map each embedded file to the corresponding inline shape page by index
            for (int i = 0; i < validObjects.Count; i++)
            {
                var obj = validObjects[i];
                if (i < inlineShapePages.Count)
                {
                    int anchorPage = inlineShapePages[i];
                    mappedObjects.Add((obj, anchorPage, $"Mapped to inline shape index {i} (anchor page {anchorPage})"));
                }
                else
                {
                    unmappedObjects.Add((obj, "Unmapped: not enough inline shapes/pages for mapping"));
                }
            }

            // Log mapping summary
#if DEBUG
            DebugLogger.Log("[PDF-INSERT] --- STRICT INLINE SHAPE INDEX MAPPING SUMMARY ---");
#endif
            foreach (var m in mappedObjects)
            {
#if DEBUG
                DebugLogger.Log($"[STRICT-MAPPED] {Path.GetFileName(m.obj.FilePath)} -> after page {m.anchorPage} ({m.reason})");
#endif
            }
            foreach (var u in unmappedObjects)
            {
#if DEBUG
                DebugLogger.Log($"[UNMAPPED] {Path.GetFileName(u.obj.FilePath)} -> after last page ({mainPageCount}) ({u.reason})");
#endif
            }

            // --- Recursively extract and convert all embedded objects and their attachments to PDF ---
            var pdfCache = new Dictionary<string, string>(); // Maps original file path to PDF path (may be same as original)
            // Track parent-child relationships for correct insertion order
            var expandedObjects = new List<(InteropEmbeddedExtractor.ExtractedObjectInfo obj, int anchorPage, string reason, string parentId)>();
            void RecursivelyConvertAndExpand(InteropEmbeddedExtractor.ExtractedObjectInfo obj, int anchorPage, string reason, string parentId)
            {
                var ext = Path.GetExtension(obj.FilePath)?.ToLowerInvariant();
                bool isPdf = IsPdfFile(obj.FilePath);
                if (isPdf)
                {
                    pdfCache[obj.FilePath] = obj.FilePath;
                    expandedObjects.Add((obj, anchorPage, reason, parentId));
                }
                else if (ext == ".docx" || ext == ".doc")
                {
                    string tempPdfPath = Path.Combine(Path.GetTempPath(), $"doc_temp_{Guid.NewGuid()}.pdf");
                    bool converted = OfficeConversionService.TryConvertOfficeToPdf(obj.FilePath, tempPdfPath);
                    if (converted && File.Exists(tempPdfPath))
                    {
                        pdfCache[obj.FilePath] = tempPdfPath;
                        pdfCache[tempPdfPath] = tempPdfPath;
                        var pdfObj = new InteropEmbeddedExtractor.ExtractedObjectInfo { FilePath = tempPdfPath, PageNumber = obj.PageNumber, CharPosition = obj.CharPosition, ProgId = obj.ProgId, OleClass = obj.OleClass, DocumentOrderIndex = obj.DocumentOrderIndex };
                        expandedObjects.Add((pdfObj, anchorPage, reason, parentId));
                    }
                    else
                    {
                        pdfCache[obj.FilePath] = null;
                        expandedObjects.Add((obj, anchorPage, reason, parentId));
                    }
                }
                else if (ext == ".xlsx")
                {
                    string tempPdfPath = Path.Combine(Path.GetTempPath(), $"xlsx_temp_{Guid.NewGuid()}.pdf");
                    if (TryConvertXlsxToPdf(obj.FilePath, tempPdfPath) && File.Exists(tempPdfPath))
                    {
                        pdfCache[obj.FilePath] = tempPdfPath;
                        pdfCache[tempPdfPath] = tempPdfPath;
                        var pdfObj = new InteropEmbeddedExtractor.ExtractedObjectInfo { FilePath = tempPdfPath, PageNumber = obj.PageNumber, CharPosition = obj.CharPosition, ProgId = obj.ProgId, OleClass = obj.OleClass, DocumentOrderIndex = obj.DocumentOrderIndex };
                        expandedObjects.Add((pdfObj, anchorPage, reason, parentId));
                    }
                    else
                    {
                        pdfCache[obj.FilePath] = null;
                        expandedObjects.Add((obj, anchorPage, reason, parentId));
                    }
                }
                else if (ext == ".msg")
                {
                    string tempPdfPath = Path.Combine(Path.GetTempPath(), $"msg_temp_{Guid.NewGuid()}.pdf");
                    var (converted, attachmentFiles) = TryConvertMsgToPdfWithAttachments(obj.FilePath, tempPdfPath);
                    string objId = obj.FilePath + "|" + obj.DocumentOrderIndex;
                    if (converted && File.Exists(tempPdfPath))
                    {
                        pdfCache[obj.FilePath] = tempPdfPath;
                        pdfCache[tempPdfPath] = tempPdfPath;
                        var pdfObj = new InteropEmbeddedExtractor.ExtractedObjectInfo { FilePath = tempPdfPath, PageNumber = obj.PageNumber, CharPosition = obj.CharPosition, ProgId = obj.ProgId, OleClass = obj.OleClass, DocumentOrderIndex = obj.DocumentOrderIndex };
                        expandedObjects.Add((pdfObj, anchorPage, reason, parentId));
                        // Recursively process only true attachments (not inline images)
                        foreach (var attPath in attachmentFiles)
                        {
                            if (File.Exists(attPath))
                            {
                                var attExt = Path.GetExtension(attPath)?.ToLowerInvariant();
                                if (attExt == ".jpg" || attExt == ".jpeg" || attExt == ".png" || attExt == ".bmp" || attExt == ".gif" || attExt == ".tif" || attExt == ".tiff" || attExt == ".webp")
                                {
                                    string imgPdfPath = Path.Combine(Path.GetTempPath(), $"img2pdf_{Guid.NewGuid()}.pdf");
                                    bool imgConverted = TryConvertImageToPdf(attPath, imgPdfPath);
                                    var imgObj = new InteropEmbeddedExtractor.ExtractedObjectInfo { FilePath = imgPdfPath, PageNumber = obj.PageNumber, CharPosition = obj.CharPosition, ProgId = null, OleClass = null, DocumentOrderIndex = obj.DocumentOrderIndex };
                                    if (imgConverted && File.Exists(imgPdfPath))
                                    {
                                        pdfCache[attPath] = imgPdfPath;
                                        pdfCache[imgPdfPath] = imgPdfPath;
                                        expandedObjects.Add((imgObj, anchorPage, "Image attachment of MSG", objId));
                                    }
                                    else
                                    {
                                        pdfCache[attPath] = null;
                                        pdfCache[imgPdfPath] = null;
                                        expandedObjects.Add((imgObj, anchorPage, "Image attachment conversion failed", objId));
                                    }
                                }
                                else
                                {
                                    var attObj = new InteropEmbeddedExtractor.ExtractedObjectInfo { FilePath = attPath, PageNumber = obj.PageNumber, CharPosition = obj.CharPosition, ProgId = null, OleClass = null, DocumentOrderIndex = obj.DocumentOrderIndex };
                                    RecursivelyConvertAndExpand(attObj, anchorPage, "Attachment of MSG", objId);
                                }
                            }
                        }
                    }
                    else
                    {
                        pdfCache[obj.FilePath] = null;
                        expandedObjects.Add((obj, anchorPage, reason, parentId));
                    }
                }
                else if (ext == ".zip")
                {
                    var zipEntries = ZipHelper.ExtractZipEntries(obj.FilePath);
                    foreach (var entry in zipEntries)
                    {
                        string tempFile = Path.Combine(Path.GetTempPath(), $"zip_entry_{Guid.NewGuid()}_{entry.FileName}");
                        File.WriteAllBytes(tempFile, entry.Data);
                        var attObj = new InteropEmbeddedExtractor.ExtractedObjectInfo { FilePath = tempFile, PageNumber = obj.PageNumber, CharPosition = obj.CharPosition, ProgId = null, OleClass = null, DocumentOrderIndex = obj.DocumentOrderIndex };
                        RecursivelyConvertAndExpand(attObj, anchorPage, "ZIP entry", obj.FilePath + "|" + obj.DocumentOrderIndex);
                        try { File.Delete(tempFile); } catch { }
                    }
                    pdfCache[obj.FilePath] = null; // ZIP itself not inserted
                }
                else if (ext == ".7z")
                {
                    string tempDir = Path.GetTempPath();
                    var allTempFiles = new List<string>();
                    string headerText = $"Extracted from {Path.GetFileName(obj.FilePath)}";
                    var parentChain = new List<string>();
                    string currentItem = Path.GetFileName(obj.FilePath);
                    string resultPdf = _attachmentService.Process7zAttachmentWithHierarchy(
                        obj.FilePath, tempDir, headerText, allTempFiles, parentChain, currentItem, false);
                    if (!string.IsNullOrEmpty(resultPdf) && File.Exists(resultPdf))
                    {
                        var attObj = new InteropEmbeddedExtractor.ExtractedObjectInfo { FilePath = resultPdf, PageNumber = obj.PageNumber, CharPosition = obj.CharPosition, ProgId = null, OleClass = null, DocumentOrderIndex = obj.DocumentOrderIndex };
                        RecursivelyConvertAndExpand(attObj, anchorPage, "7Z extracted PDF", obj.FilePath + "|" + obj.DocumentOrderIndex);
                    }
                    foreach (var f in allTempFiles) { try { File.Delete(f); } catch { } }
                    pdfCache[obj.FilePath] = null; // 7Z itself not inserted
                }
                else
                {
                    pdfCache[obj.FilePath] = null;
                    expandedObjects.Add((obj, anchorPage, reason, parentId));
                }
            }
            // Expand all mapped and unmapped objects, but only count progress for first depth
            foreach (var m in mappedObjects)
            {
                RecursivelyConvertAndExpand(m.obj, m.anchorPage, m.reason, null);
                // Progress tick for each top-level mapped object ONLY
                if (progressTick != null)
                    progressTick.Invoke();
                if (s_currentProgressCallback != null)
                    s_currentProgressCallback.Invoke();
            }
            foreach (var u in unmappedObjects)
            {
                RecursivelyConvertAndExpand(u.obj, mainPageCount, u.reason, null);
                // Progress tick for each top-level unmapped object ONLY
                if (progressTick != null)
                    progressTick.Invoke();
                if (s_currentProgressCallback != null)
                    s_currentProgressCallback.Invoke();
            }

            try
            {
                using (var outputStream = new FileStream(outputPdfPath, FileMode.Create, FileAccess.Write))
                using (var pdfWriter = new PdfWriter(outputStream))
                using (var outputPdf = new PdfDocument(pdfWriter))
                using (var mainPdf = new PdfDocument(new PdfReader(mainPdfPath)))
                {
                    for (int mainPage = 1; mainPage <= mainPageCount; mainPage++)
                    {
#if DEBUG
                        DebugLogger.Log($"[PDF-INSERT][DEBUG] Copying main page {mainPage} (output page {outputPdf.GetNumberOfPages() + 1})");
#endif
                        mainPdf.CopyPagesTo(mainPage, mainPage, outputPdf);

                        var pageObjects = expandedObjects
                            .Where(m => m.anchorPage == mainPage)
                            .OrderBy(m => m.obj.CharPosition >= 0 ? m.obj.CharPosition : int.MaxValue)
                            .ThenBy(m => m.obj.DocumentOrderIndex)
                            .ToList();
                        foreach (var m in pageObjects)
                        {
                            string pdfPath = pdfCache[m.obj.FilePath];
#if DEBUG
                            DebugLogger.Log($"[PDF-INSERT][DEBUG] Appending embedded file after main page {mainPage}: {Path.GetFileName(m.obj.FilePath)} (anchor={m.anchorPage}, order={m.obj.DocumentOrderIndex})");
#endif
                            if (!string.IsNullOrEmpty(pdfPath) && File.Exists(pdfPath))
                            {
                                try
                                {
                                    using (var embeddedReader = new PdfReader(pdfPath))
                                    using (var embeddedPdf = new PdfDocument(embeddedReader))
                                    {
                                        int embeddedPageCount = embeddedPdf.GetNumberOfPages();
#if DEBUG
                                        DebugLogger.Log($"[PDF-INSERT][DEBUG] Embedded file {Path.GetFileName(m.obj.FilePath)} has {embeddedPageCount} pages");
#endif
                                        for (int ep = 1; ep <= embeddedPageCount; ep++)
                                        {
                                            embeddedPdf.CopyPagesTo(ep, ep, outputPdf);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
#if DEBUG
                                    DebugLogger.Log($"[PDF-INSERT][ERROR] Failed to copy embedded PDF {pdfPath}: {ex.Message}");
#endif
                                }
                            }
                            else
                            {
                                int before = outputPdf.GetNumberOfPages();
                                InsertPlaceholderForFile(m.obj.FilePath, outputPdf, before, Path.GetExtension(m.obj.FilePath)?.ToLowerInvariant());
                            }
                        }
                    }
                    // Append unmapped objects after last page only
                    foreach (var u in unmappedObjects)
                    {
                        string pdfPath = pdfCache[u.obj.FilePath];
                        if (!string.IsNullOrEmpty(pdfPath) && File.Exists(pdfPath))
                        {
                            try
                            {
                                using (var embeddedReader = new PdfReader(pdfPath))
                                using (var embeddedPdf = new PdfDocument(embeddedReader))
                                {
                                    int embeddedPageCount = embeddedPdf.GetNumberOfPages();
                                    for (int ep = 1; ep <= embeddedPageCount; ep++)
                                    {
                                        embeddedPdf.CopyPagesTo(ep, ep, outputPdf);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
#if DEBUG
                                DebugLogger.Log($"[PDF-INSERT][ERROR] Failed to copy embedded PDF {pdfPath}: {ex.Message}");
#endif
                            }
                        }
                        else
                        {
                            int before = outputPdf.GetNumberOfPages();
                            InsertPlaceholderForFile(u.obj.FilePath, outputPdf, before, Path.GetExtension(u.obj.FilePath)?.ToLowerInvariant());
                        }
                    }
                }
                // Cleanup temp PDFs
                foreach (var kvp in pdfCache)
                {
                    if (kvp.Value != null && kvp.Value != kvp.Key && File.Exists(kvp.Value))
                    {
                        try { File.Delete(kvp.Value); } catch { }
                    }
                }


            }
            catch (Exception ex)
            {
#if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error creating PDF with embedded files: {ex.Message}");
#endif
                try
                {
                    File.Copy(mainPdfPath, outputPdfPath, true);
#if DEBUG
                    DebugLogger.Log("[PDF-INSERT] Fallback: copied main PDF without embedded files");
#endif
                }
                catch (Exception copyEx)
                {
#if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] Fallback copy failed: {copyEx.Message}");
#endif
                }
            }
        }

        // Insert embedded object without separator
        private static HashSet<string> processedArchiveHashes = new HashSet<string>(); // Archive deduplication
        private static int InsertEmbeddedObject_NoSeparator(InteropEmbeddedExtractor.ExtractedObjectInfo obj, PdfDocument outputPdf, int currentOutputPage, Action progressTick = null)
        {
            try
            {
                // Enhanced debug logging for file type detection
                string extMain = Path.GetExtension(obj.FilePath)?.ToLowerInvariant();
                bool isPdf = IsPdfFile(obj.FilePath);
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT][DEBUG] InsertEmbeddedObject_NoSeparator: {obj.FilePath} (ext: {extMain}) IsPdfFile={isPdf}");
                #endif

                // Treat any file that is a PDF by header as a PDF, regardless of extension
                if (isPdf)
                {
                    return InsertPdfFile_NoSeparator(obj.FilePath, outputPdf, currentOutputPage, obj.OleClass, progressTick);
                }
                else if (obj.FilePath.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                {
                    return InsertMsgFile_NoSeparator(obj.FilePath, outputPdf, currentOutputPage, progressTick);
                }
                else if (obj.FilePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    return InsertDocxFile_NoSeparator(obj.FilePath, outputPdf, currentOutputPage);
                }
                else if (obj.FilePath.EndsWith(".doc", StringComparison.OrdinalIgnoreCase))
                {
                    // Convert .doc to PDF using OfficeConversionService and insert
                    string tempPdfPath = Path.Combine(Path.GetTempPath(), $"doc_temp_{Guid.NewGuid()}.pdf");
                    try
                    {
                        bool converted = OfficeConversionService.TryConvertOfficeToPdf(obj.FilePath, tempPdfPath);
                        if (converted && File.Exists(tempPdfPath))
                        {
                            return InsertPdfFile_NoSeparator(tempPdfPath, outputPdf, currentOutputPage, "DOC", progressTick);
                        }
                        else
                        {
                            return InsertPlaceholderForFile(obj.FilePath, outputPdf, currentOutputPage, "DOC");
                        }
                    }
                    finally
                    {
                        try { if (File.Exists(tempPdfPath)) File.Delete(tempPdfPath); } catch { }
                    }
                }
                else if (obj.FilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    return InsertXlsxFile_NoSeparator(obj.FilePath, outputPdf, currentOutputPage);
                }
                else if (obj.FilePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    // Convert .xls to PDF using OfficeConversionService and insert
                    string tempPdfPath = Path.Combine(Path.GetTempPath(), $"xls_temp_{Guid.NewGuid()}.pdf");
                    try
                    {
                        bool converted = OfficeConversionService.TryConvertOfficeToPdf(obj.FilePath, tempPdfPath);
                        if (converted && File.Exists(tempPdfPath))
                        {
                            return InsertPdfFile_NoSeparator(tempPdfPath, outputPdf, currentOutputPage, "XLS", progressTick);
                        }
                        else
                        {
                            return InsertPlaceholderForFile(obj.FilePath, outputPdf, currentOutputPage, "XLS");
                        }
                    }
                    finally
                    {
                        try { if (File.Exists(tempPdfPath)) File.Delete(tempPdfPath); } catch { }
                    }
                }
                else if (obj.FilePath.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                {
                    // --- ZIP HANDLING ---
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** ZIP PROCESSING START *** Extracting and inserting ZIP: {Path.GetFileName(obj.FilePath)} after page {currentOutputPage}");
                    #endif
                    try
                    {
                        var zipEntries = ZipHelper.ExtractZipEntries(obj.FilePath);
                        foreach (var entry in zipEntries)
                        {
                            // Save each entry to a temp file
                            string tempFile = Path.Combine(Path.GetTempPath(), $"zip_entry_{Guid.NewGuid()}_{entry.FileName}");
                            File.WriteAllBytes(tempFile, entry.Data);
                            string entryExt = Path.GetExtension(entry.FileName).ToLowerInvariant(); // FIX: avoid shadowing
                            bool entryIsPdf = IsPdfFile(tempFile);
                            #if DEBUG
                            DebugLogger.Log($"[PDF-INSERT][DEBUG] ZIP entry: {entry.FileName} (ext: {entryExt}) IsPdfFile={entryIsPdf}");
                            #endif
                            if (entryIsPdf)
                            {
                                currentOutputPage = InsertPdfFile_NoSeparator(tempFile, outputPdf, currentOutputPage, "ZIP-PDF", progressTick);
                            }
                            else if (entryExt == ".docx")
                            {
                                currentOutputPage = InsertDocxFile_NoSeparator(tempFile, outputPdf, currentOutputPage);
                            }
                            else if (entryExt == ".xlsx")
                            {
                                currentOutputPage = InsertXlsxFile_NoSeparator(tempFile, outputPdf, currentOutputPage);
                            }
                            else if (entryExt == ".msg")
                            {
                                currentOutputPage = InsertMsgFile_NoSeparator(tempFile, outputPdf, currentOutputPage);
                            }
                            else
                            {
                                currentOutputPage = InsertPlaceholderForFile(tempFile, outputPdf, currentOutputPage, $"ZIP Entry ({entryExt})");
                            }
                            try { File.Delete(tempFile); } catch { }
                        }
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** ZIP PROCESSING COMPLETE *** {zipEntries.Count} entries processed from {Path.GetFileName(obj.FilePath)}");
                        #endif
                    }
                    catch (Exception zipEx)
                    {
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] Error processing ZIP {obj.FilePath}: {zipEx.Message}");
                        #endif
                        currentOutputPage = InsertErrorPlaceholder(obj.FilePath, outputPdf, currentOutputPage, zipEx.Message);
                    }
                    return currentOutputPage;
                }
                else if (obj.FilePath.EndsWith(".7z", StringComparison.OrdinalIgnoreCase))
                {
                    // --- 7Z HANDLING ---
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** 7Z PROCESSING START *** Extracting and inserting 7Z: {Path.GetFileName(obj.FilePath)} after page {currentOutputPage}");
                    #endif
                    try
                    {
                        string tempDir = Path.GetTempPath();
                        var allTempFiles = new List<string>();
                        string headerText = $"Extracted from {Path.GetFileName(obj.FilePath)}";
                        var parentChain = new List<string>();
                        string currentItem = Path.GetFileName(obj.FilePath);
                        // Recursively process 7z and get the resulting PDF (may be a merged PDF of all contents)
                        string resultPdf = _attachmentService.Process7zAttachmentWithHierarchy(
                            obj.FilePath, tempDir, headerText, allTempFiles, parentChain, currentItem, false);
                        if (!string.IsNullOrEmpty(resultPdf) && File.Exists(resultPdf))
                        {
                            bool resultIsPdf = IsPdfFile(resultPdf);
                            #if DEBUG
                            DebugLogger.Log($"[PDF-INSERT][DEBUG] 7Z result: {resultPdf} IsPdfFile={resultIsPdf}");
                            #endif
                            if (resultIsPdf)
                                currentOutputPage = InsertPdfFile_NoSeparator(resultPdf, outputPdf, currentOutputPage, "7Z-PDF", progressTick);
                            else
                            {
                                currentOutputPage = InsertPlaceholderForFile(resultPdf, outputPdf, currentOutputPage, "7Z");
                            }
                        }
                        else
                        {
                            currentOutputPage = InsertPlaceholderForFile(obj.FilePath, outputPdf, currentOutputPage, "7Z");
                        }
                        // Cleanup temp files
                        foreach (var f in allTempFiles) { try { File.Delete(f); } catch { } }
                    }
                    catch (Exception sevenZEx)
                    {
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] Error processing 7Z {obj.FilePath}: {sevenZEx.Message}");
                        #endif
                        currentOutputPage = InsertErrorPlaceholder(obj.FilePath, outputPdf, currentOutputPage, sevenZEx.Message);
                    }
                    return currentOutputPage;
                }
                else
                {
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT][DEBUG] File {obj.FilePath} not recognized as supported type, inserting placeholder.");
                    #endif
                    // Only for unsupported types, add a placeholder
                    return InsertPlaceholderForFile(obj.FilePath, outputPdf, currentOutputPage, obj.OleClass);
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error inserting {obj.FilePath}: {ex.Message}");
                #endif
                return InsertErrorPlaceholder(obj.FilePath, outputPdf, currentOutputPage, ex.Message);
            }
        }

        // --- Helper: Compute SHA256 hash of a file ---
        private static string ComputeFileHash(string filePath)
        {
            try
            {
                using (var stream = File.OpenRead(filePath))
                using (var sha = SHA256.Create())
                {
                    var hashBytes = sha.ComputeHash(stream);
                    return BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Failed to compute hash for {filePath}: {ex.Message}");
                #endif
                return null;
            }
        }

        // Helper to check if a file is a PDF by header (robust: scans first 1KB for %PDF-)
        private static bool IsPdfFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return false;
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    byte[] buffer = new byte[1024];
                    int read = fs.Read(buffer, 0, buffer.Length);
                    string content = System.Text.Encoding.ASCII.GetString(buffer, 0, read);
                    int idx = content.IndexOf("%PDF-");
                    if (idx >= 0)
                    {
                    if (idx > 0)
                    {
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT][IsPdfFile] Found '%PDF-' at offset {idx} in {filePath}, treating as PDF (nonzero offset)");
                        #endif
                    }
                        return true;
                    }
                    else
                    {
                        // Log first 32 bytes for debugging
                        string hex = BitConverter.ToString(buffer, 0, Math.Min(32, read)).Replace("-", " ");
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT][IsPdfFile] No '%PDF-' found in first 1KB of {filePath}. First 32 bytes: {hex}");
                        #endif
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT][IsPdfFile] Exception for {filePath}: {ex.Message}");
                #endif
                return false;
            }
        }

        // Insert PDF file without separator
        private static int InsertPdfFile_NoSeparator(string pdfPath, PdfDocument outputPdf, int currentPage, string oleClass, Action progressTick = null)
        {
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] *** PDF INSERTION START *** Inserting PDF: {Path.GetFileName(pdfPath)} after page {currentPage} (current total pages: {outputPdf.GetNumberOfPages()})");
            #endif
            try
            {
                if (!File.Exists(pdfPath))
                {
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] PDF file not found: {pdfPath}");
                    #endif
                    return InsertErrorPlaceholder(pdfPath, outputPdf, currentPage, "File not found");
                }
                var fileInfo = new FileInfo(pdfPath);
                if (fileInfo.Length == 0)
                {
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] PDF file is empty: {pdfPath}");
                    #endif
                    return InsertErrorPlaceholder(pdfPath, outputPdf, currentPage, "Empty file");
                }
                PdfReader reader = null;
                PdfDocument embeddedPdf = null;
                try
                {
                    reader = new PdfReader(pdfPath);
                    embeddedPdf = new PdfDocument(reader);
                    int embeddedPageCount = embeddedPdf.GetNumberOfPages();
                    
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** PDF CONTENT *** {Path.GetFileName(pdfPath)} has {embeddedPageCount} pages to insert");
                    #endif
                    
                    // Copy pages one by one to append them after currentPage
                    for (int pageNum = 1; pageNum <= embeddedPageCount; pageNum++)
                    {
                        int totalPagesBefore = outputPdf.GetNumberOfPages();
                        // CopyPagesTo appends to the end, which is what we want for sequential insertion
                        embeddedPdf.CopyPagesTo(pageNum, pageNum, outputPdf);
                        currentPage++;
                        
                        int totalPagesAfter = outputPdf.GetNumberOfPages();
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** PDF PAGE COPY *** Copied page {pageNum}/{embeddedPageCount} from {Path.GetFileName(pdfPath)}, output PDF went from {totalPagesBefore} to {totalPagesAfter} pages, tracking currentPage={currentPage}");
                        #endif
                    }
                    
                    // Progress tick REMOVED: Only top-level objects should trigger ticks
                    
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** PDF INSERTION COMPLETE *** Successfully inserted {embeddedPageCount} pages from {Path.GetFileName(pdfPath)}, final total pages: {outputPdf.GetNumberOfPages()}");
                    #endif
                }
                finally
                {
                    try { embeddedPdf?.Close(); reader?.Close(); } catch { }
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error reading PDF {pdfPath}: {ex.Message}");
                #endif
                currentPage = InsertErrorPlaceholder(pdfPath, outputPdf, currentPage, ex.Message);
            }
            return currentPage;
        }

        // Insert MSG file without separator
        private static int InsertMsgFile_NoSeparator(string msgPath, PdfDocument outputPdf, int currentPage, Action progressTick = null)
        {
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] Converting and inserting MSG: {Path.GetFileName(msgPath)} after page {currentPage}");
            #endif
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
                                #if DEBUG
                                DebugLogger.Log($"[PDF-INSERT] Inserting MSG attachment: {Path.GetFileName(attachmentPath)}");
                                #endif
                                currentPage = InsertAttachmentFile(attachmentPath, outputPdf, currentPage, progressTick);
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
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error processing MSG {msgPath}: {ex.Message}");
                #endif
                currentPage = InsertErrorPlaceholder(msgPath, outputPdf, currentPage, ex.Message);
            }
            return currentPage;
        }

        // Insert DOCX file without separator
        private static int InsertDocxFile_NoSeparator(string docxPath, PdfDocument outputPdf, int currentPage, Action progressTick = null)
        {
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] Converting and inserting DOCX: {Path.GetFileName(docxPath)} after page {currentPage}");
            #endif
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
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error processing DOCX {docxPath}: {ex.Message}");
                #endif
                currentPage = InsertErrorPlaceholder(docxPath, outputPdf, currentPage, ex.Message);
            }
            return currentPage;
        }

        // Insert XLSX file without separator
        private static int InsertXlsxFile_NoSeparator(string xlsxPath, PdfDocument outputPdf, int currentPage, Action progressTick = null)
        {
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] *** XLSX PROCESSING START *** Converting and inserting XLSX: {Path.GetFileName(xlsxPath)} after page {currentPage}");
            #endif
            try
            {
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"xlsx_temp_{Guid.NewGuid()}.pdf");
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] *** XLSX CONVERSION *** Temporary PDF path: {tempPdfPath}");
                #endif
                
                try
                {
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** XLSX CONVERSION *** Starting Excel to PDF conversion for {Path.GetFileName(xlsxPath)}");
                    #endif
                    bool converted = TryConvertXlsxToPdf(xlsxPath, tempPdfPath);
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** XLSX CONVERSION RESULT *** Conversion successful: {converted}");
                    #endif
                    
                    if (converted && File.Exists(tempPdfPath))
                    {
                        var fileInfo = new FileInfo(tempPdfPath);
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** XLSX PDF CREATED *** Temp PDF exists, size: {fileInfo.Length} bytes");
                        DebugLogger.Log($"[PDF-INSERT] *** XLSX PDF INSERTION *** Now treating converted XLSX as regular PDF");
                        #endif
                        currentPage = InsertPdfFile_NoSeparator(tempPdfPath, outputPdf, currentPage, "XLSX");
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** XLSX PDF INSERTED *** Successfully inserted converted XLSX as PDF");
                        #endif
                    }
                    else
                    {
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** XLSX CONVERSION FAILED *** Conversion failed or file doesn't exist, inserting placeholder");
                        #endif
                        currentPage = InsertPlaceholderForFile(xlsxPath, outputPdf, currentPage, "XLSX");
                    }
                }
                finally
                {
                    if (File.Exists(tempPdfPath)) 
                    { 
                        try 
                        { 
                            File.Delete(tempPdfPath); 
                            #if DEBUG
                            DebugLogger.Log($"[PDF-INSERT] *** XLSX CLEANUP *** Deleted temporary PDF: {Path.GetFileName(tempPdfPath)}");
                            #endif
                        } 
                        catch (Exception cleanupEx)
                        {
                            #if DEBUG
                            DebugLogger.Log($"[PDF-INSERT] *** XLSX CLEANUP ERROR *** Failed to delete temp file: {cleanupEx.Message}");
                            #endif
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] *** XLSX ERROR *** Error processing XLSX {xlsxPath}: {ex.Message}");
                #endif
                currentPage = InsertErrorPlaceholder(xlsxPath, outputPdf, currentPage, ex.Message);
            }
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] *** XLSX PROCESSING COMPLETE *** Final currentPage: {currentPage}");
            #endif
            return currentPage;
        }

        /// <summary>
        /// Attempts to convert MSG to PDF using the main HTML-to-PDF pipeline (DinkToPdf/HtmlToPdfWorker)
        /// </summary>
        private static bool TryConvertMsgToPdf(string msgPath, string outputPdfPath)
        {
            try
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Converting MSG to PDF (HTML pipeline): {msgPath} -> {outputPdfPath}");
                #endif
                using (var msg = new MsgReader.Outlook.Storage.Message(msgPath))
                {
                    // Build HTML with inline images using the main service
                    var htmlResult = _emailConverterService.BuildEmailHtmlWithInlineImages(msg, false);
                    string html = htmlResult.Html;
                    var tempHtmlPath = Path.Combine(Path.GetTempPath(), $"msg2pdf_{Guid.NewGuid()}.html");
                    var appTempDir = Path.Combine(Path.GetTempPath(), "MsgToPdfConverter");
                    var tempHtmlPathFixed = Path.Combine(appTempDir, $"msg2pdf_{Guid.NewGuid()}.html");
                    File.WriteAllText(tempHtmlPathFixed, html, System.Text.Encoding.UTF8);
                    tempHtmlPath = tempHtmlPathFixed;

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
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] Successfully converted MSG to PDF: {outputPdfPath}");
                        #endif
                        return true;
                    }
                    else
                    {
                        // Dump HTML to debug file for inspection
                        var debugHtmlPath = Path.Combine(Path.Combine(Path.GetTempPath(), "MsgToPdfConverter"), $"debug_email_{DateTime.Now:yyyyMMdd_HHmmss}_fail.html");
                        File.WriteAllText(debugHtmlPath, html, System.Text.Encoding.UTF8);
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] HtmlToPdfWorker failed.\nSTDOUT: {stdOut}\nSTDERR: {stdErr}\nHTML dumped to: {debugHtmlPath}");
                        #endif
                        // Optionally: track for cleanup if you have a temp file list in this context
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Failed to convert MSG to PDF: {ex.Message}\n{ex}");
                #endif
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
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Converting MSG to PDF with attachments: {msgPath} -> {outputPdfPath}");
                #endif
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
                                    #if DEBUG
                                    DebugLogger.Log($"[PDF-INSERT] Extracted MSG attachment: {fileAttachment.FileName} -> {tempAttachmentPath}");
                                    #endif
                                }
                                catch (Exception ex)
                                {
                                    #if DEBUG
                                    DebugLogger.Log($"[PDF-INSERT] Failed to extract attachment {fileAttachment.FileName}: {ex.Message}");
                                    #endif
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
                                    #if DEBUG
                                    DebugLogger.Log($"[PDF-INSERT] Extracted nested MSG: {nestedMsg.Subject} -> {tempMsgPath}");
                                    #endif
                                }
                                catch (Exception ex)
                                {
                                    #if DEBUG
                                    DebugLogger.Log($"[PDF-INSERT] Failed to extract nested MSG {nestedMsg.Subject}: {ex.Message}");
                                    #endif
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
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Failed to convert MSG with attachments: {ex.Message}");
                #endif
                
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
        private static int InsertAttachmentFile(string attachmentPath, PdfDocument outputPdf, int currentPage, Action progressTick = null)
        {
            try
            {
                var ext = Path.GetExtension(attachmentPath)?.ToLowerInvariant();
                switch (ext)
                {
                    case ".pdf":
                        return InsertPdfFile_NoSeparator(attachmentPath, outputPdf, currentPage, "Attachment", progressTick);
                    case ".docx":
                        return InsertDocxFile_NoSeparator(attachmentPath, outputPdf, currentPage, progressTick);
                    case ".xlsx":
                        return InsertXlsxFile_NoSeparator(attachmentPath, outputPdf, currentPage, progressTick);
                    case ".msg":
                        return InsertMsgFile_NoSeparator(attachmentPath, outputPdf, currentPage, progressTick);
                    case ".jpg":
                    case ".jpeg":
                    case ".png":
                    case ".bmp":
                    case ".gif":
                    case ".tif":
                    case ".tiff":
                    case ".webp":
                    {
                        // Create temp PDF with only the image, no scaling/margins/header (identical to single email logic)
                        string tempPdf = Path.Combine(Path.GetTempPath(), $"img2pdf_{Guid.NewGuid()}.pdf");
                        try
                        {
                            using (var writer = new iText.Kernel.Pdf.PdfWriter(tempPdf))
                            using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                            using (var docImg = new iText.Layout.Document(pdf))
                            {
                                var imgData = iText.IO.Image.ImageDataFactory.Create(attachmentPath);
                                var image = new iText.Layout.Element.Image(imgData);
                                docImg.Add(image);
                            }
                            int result = InsertPdfFile_NoSeparator(tempPdf, outputPdf, currentPage, "Image", progressTick);
                            return result;
                        }
                        catch (Exception ex)
                        {

                            return InsertErrorPlaceholder(attachmentPath, outputPdf, currentPage, ex.Message);
                        }
                        finally
                        {
                            try { if (File.Exists(tempPdf)) File.Delete(tempPdf); } catch { }
                        }
                    }
                    default:
                        return InsertPlaceholderForFile(attachmentPath, outputPdf, currentPage, $"Attachment ({ext})");
                }
            }
            catch (Exception ex)
            {

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
        private static int InsertEmbeddedObject(InteropEmbeddedExtractor.ExtractedObjectInfo obj, PdfDocument outputPdf, int currentOutputPage, Action progressTick = null)
        {
            // Route all calls to the new no-separator version
            return InsertEmbeddedObject_NoSeparator(obj, outputPdf, currentOutputPage, progressTick);
        }

        /// <summary>
        /// Attempts to convert DOCX to PDF using Word Interop
        /// </summary>
        private static bool TryConvertDocxToPdf(string docxPath, string outputPdfPath)
        {
            try
            {

                
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

                    
                    // Export to PDF with minimal settings
                    doc.ExportAsFixedFormat(outputPdfPath, 
                        Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF,
                        OpenAfterExport: false,
                        OptimizeFor: Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForPrint);
                    

                    
                    // Allow a moment for file to be written
                    System.Threading.Thread.Sleep(500);
                    
                    if (File.Exists(outputPdfPath) && new FileInfo(outputPdfPath).Length > 0)
                    {

                        return true;
                    }
                    else
                    {

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
                            #if DEBUG
                            DebugLogger.Log($"[PDF-INSERT] Warning: Document cleanup failed: {cleanupEx.Message}");
                            #endif
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
                            #if DEBUG
                            DebugLogger.Log($"[PDF-INSERT] Warning: Application cleanup failed: {cleanupEx.Message}");
                            #endif
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
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Failed to convert DOCX to PDF: {ex.Message}");
                DebugLogger.Log($"[PDF-INSERT] Exception details: {ex}");
                #endif
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

            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] *** EXCEL CONVERSION START *** Converting {Path.GetFileName(xlsxPath)} to PDF");
            #endif

            // Run Excel conversion in STA thread like OfficeConversionService to avoid popup issues
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** EXCEL INTEROP *** Creating Excel application in STA thread");
                    #endif
                    
                    var excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = false;
                    excelApp.DisplayAlerts = false;
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** EXCEL INTEROP *** Excel application created successfully");
                    #endif
                    
                    Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
                    Microsoft.Office.Interop.Excel.Workbook wb = null;
                    try
                    {
                        workbooks = excelApp.Workbooks;
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** EXCEL INTEROP *** Opening workbook: {Path.GetFileName(xlsxPath)}");
                        #endif
                        wb = workbooks.Open(xlsxPath);
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** EXCEL INTEROP *** Workbook opened successfully");
                        #endif
                        
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** EXCEL EXPORT *** Exporting to PDF: {outputPdfPath}");
                        #endif
                        wb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPdfPath);
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** EXCEL EXPORT *** Export completed successfully");
                        #endif
                        result = true;
                    }
                    finally
                    {
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** EXCEL CLEANUP *** Cleaning up Excel COM objects");
                        #endif
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
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] *** EXCEL CLEANUP *** Cleanup completed");
                        #endif
                    }
                }
                catch (Exception ex)
                {
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** EXCEL ERROR *** Exception in Excel conversion thread: {ex.Message}");
                    DebugLogger.Log($"[PDF-INSERT] *** EXCEL ERROR *** Stack trace: {ex.StackTrace}");
                    #endif
                    threadEx = ex;
                }
            });
            
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] *** EXCEL THREAD COMPLETE *** Thread finished, checking results");
            #endif
            
            if (threadEx != null)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] *** EXCEL CONVERSION FAILED *** Thread exception: {threadEx.Message}");
                #endif
                return false;
            }
            
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] *** EXCEL RESULT CHECK *** result={result}, file exists={File.Exists(outputPdfPath)}");
            #endif
            if (File.Exists(outputPdfPath))
            {
                var fileInfo = new FileInfo(outputPdfPath);
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] *** EXCEL RESULT CHECK *** Output file size: {fileInfo.Length} bytes");
                #endif
            }
            
            if (result && File.Exists(outputPdfPath) && new FileInfo(outputPdfPath).Length > 0)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] *** EXCEL SUCCESS *** Successfully converted XLSX to PDF: {outputPdfPath}");
                #endif
                return true;
            }
            else
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] *** EXCEL FAILURE *** XLSX conversion failed: output file not created or empty");
                #endif
                return false;
            }
        }

        // Insert embedded object at a specific position in the PDF (used for proper page ordering)
        private static void InsertEmbeddedObjectAtPosition(InteropEmbeddedExtractor.ExtractedObjectInfo obj, PdfDocument outputPdf, int afterPageNumber)
        {
            try
            {
                if (obj.FilePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    InsertPdfFileAtPosition(obj.FilePath, outputPdf, afterPageNumber, obj.OleClass);
                }
                else if (obj.FilePath.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                {
                    InsertMsgFileAtPosition(obj.FilePath, outputPdf, afterPageNumber);
                }
                else if (obj.FilePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    InsertDocxFileAtPosition(obj.FilePath, outputPdf, afterPageNumber);
                }
                else if (obj.FilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    InsertXlsxFileAtPosition(obj.FilePath, outputPdf, afterPageNumber);
                }
                else
                {
                    // Only for unsupported types, add a placeholder
                    InsertPlaceholderAtPosition(obj.FilePath, outputPdf, afterPageNumber, obj.OleClass);
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error inserting {obj.FilePath} at position {afterPageNumber}: {ex.Message}");
                #endif
                InsertErrorPlaceholderAtPosition(obj.FilePath, outputPdf, afterPageNumber, ex.Message);
            }
        }

        // Insert PDF file at specific position
        private static void InsertPdfFileAtPosition(string pdfPath, PdfDocument outputPdf, int afterPageNumber, string oleClass)
        {
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] *** PDF POSITION INSERTION *** Inserting PDF: {Path.GetFileName(pdfPath)} after page {afterPageNumber}");
            #endif
            try
            {
                if (!File.Exists(pdfPath))
                {
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] PDF file not found: {pdfPath}");
                    #endif
                    InsertErrorPlaceholderAtPosition(pdfPath, outputPdf, afterPageNumber, "File not found");
                    return;
                }
                var fileInfo = new FileInfo(pdfPath);
                if (fileInfo.Length == 0)
                {
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] PDF file is empty: {pdfPath}");
                    #endif
                    InsertErrorPlaceholderAtPosition(pdfPath, outputPdf, afterPageNumber, "Empty file");
                    return;
                }
                
                PdfReader reader = null;
                PdfDocument embeddedPdf = null;
                try
                {
                    reader = new PdfReader(pdfPath);
                    embeddedPdf = new PdfDocument(reader);
                    int embeddedPageCount = embeddedPdf.GetNumberOfPages();
                    
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** PDF CONTENT *** {Path.GetFileName(pdfPath)} has {embeddedPageCount} pages to insert after page {afterPageNumber}");
                    #endif
                    
                    // Since CopyPagesTo doesn't support insertion at specific position, 
                    // we need to copy to a temporary PDF and then insert the pages manually
                    using (var tempStream = new MemoryStream())
                    {
                        using (var tempWriter = new PdfWriter(tempStream))
                        using (var tempPdf = new PdfDocument(tempWriter))
                        {
                            // Copy all pages from embedded PDF to temp PDF
                            embeddedPdf.CopyPagesTo(1, embeddedPageCount, tempPdf);
                        }
                        
                        tempStream.Seek(0, SeekOrigin.Begin);
                        using (var tempReader = new PdfReader(tempStream))
                        using (var tempPdfDoc = new PdfDocument(tempReader))
                        {
                            // Now copy pages from temp PDF to output PDF at the specific position
                            for (int pageNum = 1; pageNum <= embeddedPageCount; pageNum++)
                            {
                                int totalPagesBefore = outputPdf.GetNumberOfPages();
                                var pageToCopy = tempPdfDoc.GetPage(pageNum);
                                var copiedPage = pageToCopy.CopyTo(outputPdf);
                                
                                // Move the copied page to the desired position
                                outputPdf.MovePage(outputPdf.GetNumberOfPages(), afterPageNumber + pageNum);
                                
                                int totalPagesAfter = outputPdf.GetNumberOfPages();
                                #if DEBUG
                                DebugLogger.Log($"[PDF-INSERT] *** PDF PAGE POSITION INSERT *** Inserted page {pageNum}/{embeddedPageCount} from {Path.GetFileName(pdfPath)} at position {afterPageNumber + pageNum}, PDF went from {totalPagesBefore} to {totalPagesAfter} pages");
                                #endif
                            }
                        }
                    }
                    
                    #if DEBUG
                    DebugLogger.Log($"[PDF-INSERT] *** PDF POSITION INSERTION COMPLETE *** Successfully inserted {embeddedPageCount} pages from {Path.GetFileName(pdfPath)} after page {afterPageNumber}");
                    #endif
                }
                finally
                {
                    try { embeddedPdf?.Close(); reader?.Close(); } catch { }
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error reading PDF {pdfPath}: {ex.Message}");
                #endif
                InsertErrorPlaceholderAtPosition(pdfPath, outputPdf, afterPageNumber, ex.Message);
            }
        }

        // Insert DOCX file at specific position
        private static void InsertDocxFileAtPosition(string docxPath, PdfDocument outputPdf, int afterPageNumber)
        {
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] Converting and inserting DOCX at position: {Path.GetFileName(docxPath)} after page {afterPageNumber}");
            #endif
            try
            {
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"docx_temp_{Guid.NewGuid()}.pdf");
                
                if (TryConvertDocxToPdf(docxPath, tempPdfPath))
                {
                    InsertPdfFileAtPosition(tempPdfPath, outputPdf, afterPageNumber, "Word.Document");
                    try { File.Delete(tempPdfPath); } catch { }
                }
                else
                {
                    InsertErrorPlaceholderAtPosition(docxPath, outputPdf, afterPageNumber, "DOCX conversion failed");
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error converting DOCX {docxPath}: {ex.Message}");
                #endif
                InsertErrorPlaceholderAtPosition(docxPath, outputPdf, afterPageNumber, ex.Message);
            }
        }

        // Insert XLSX file at specific position
        private static void InsertXlsxFileAtPosition(string xlsxPath, PdfDocument outputPdf, int afterPageNumber)
        {
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] Converting and inserting XLSX at position: {Path.GetFileName(xlsxPath)} after page {afterPageNumber}");
            #endif
            try
            {
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"xlsx_temp_{Guid.NewGuid()}.pdf");
                
                if (TryConvertXlsxToPdf(xlsxPath, tempPdfPath))
                {
                    InsertPdfFileAtPosition(tempPdfPath, outputPdf, afterPageNumber, "Excel.Sheet");
                    try { File.Delete(tempPdfPath); } catch { }
                }
                else
                {
                    InsertErrorPlaceholderAtPosition(xlsxPath, outputPdf, afterPageNumber, "XLSX conversion failed");
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error converting XLSX {xlsxPath}: {ex.Message}");
                #endif
                InsertErrorPlaceholderAtPosition(xlsxPath, outputPdf, afterPageNumber, ex.Message);
            }
        }

        // Insert MSG file at specific position
        private static void InsertMsgFileAtPosition(string msgPath, PdfDocument outputPdf, int afterPageNumber)
        {
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] Converting and inserting MSG at position: {Path.GetFileName(msgPath)} after page {afterPageNumber}");
            #endif
            try
            {
                string tempPdfPath = Path.Combine(Path.GetTempPath(), $"msg_temp_{Guid.NewGuid()}.pdf");
                
                var (converted, attachmentPaths) = TryConvertMsgToPdfWithAttachments(msgPath, tempPdfPath);
                if (converted)
                {
                    InsertPdfFileAtPosition(tempPdfPath, outputPdf, afterPageNumber, "Outlook.Message");
                    
                    // Insert attachments after the main MSG content
                    int currentPosition = afterPageNumber;
                    // Count how many pages were just inserted from the MSG
                    using (var tempReader = new PdfReader(tempPdfPath))
                    using (var tempPdf = new PdfDocument(tempReader))
                    {
                        currentPosition += tempPdf.GetNumberOfPages();
                    }
                    
                    foreach (string attachmentPath in attachmentPaths)
                    {
                        #if DEBUG
                        DebugLogger.Log($"[PDF-INSERT] Inserting MSG attachment at position: {Path.GetFileName(attachmentPath)}");
                        #endif
                        InsertPdfFileAtPosition(attachmentPath, outputPdf, currentPosition, "Attachment");
                        
                        // Update position for next attachment
                        using (var attReader = new PdfReader(attachmentPath))
                        using (var attPdf = new PdfDocument(attReader))
                        {
                            currentPosition += attPdf.GetNumberOfPages();
                        }
                    }
                    
                    try { File.Delete(tempPdfPath); } catch { }
                    foreach (string attPath in attachmentPaths)
                    {
                        try { File.Delete(attPath); } catch { }
                    }
                }
                else
                {
                    InsertErrorPlaceholderAtPosition(msgPath, outputPdf, afterPageNumber, "MSG conversion failed");
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Error converting MSG {msgPath}: {ex.Message}");
                #endif
                InsertErrorPlaceholderAtPosition(msgPath, outputPdf, afterPageNumber, ex.Message);
            }
        }

        // Insert placeholder at specific position
        private static void InsertPlaceholderAtPosition(string filePath, PdfDocument outputPdf, int afterPageNumber, string oleClass)
        {
            #if DEBUG
            DebugLogger.Log($"[PDF-INSERT] Inserting placeholder at position for unsupported file: {Path.GetFileName(filePath)} after page {afterPageNumber}");
            #endif
            InsertErrorPlaceholderAtPosition(filePath, outputPdf, afterPageNumber, $"Unsupported file type: {Path.GetExtension(filePath)}");
        }

        // Insert error placeholder at specific position
        private static void InsertErrorPlaceholderAtPosition(string filePath, PdfDocument outputPdf, int afterPageNumber, string errorMessage)
        {
            try
            {
                string fileName = Path.GetFileName(filePath);
                string errorInfo = $"File: {fileName}\nError: {errorMessage}";
                
                // Insert a new page at the specified position
                var page = outputPdf.AddNewPage(afterPageNumber + 1);
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
                
                // Use canvas text operations
                canvas.BeginText();
                
                try
                {
                    // Load default font
                    var font = iText.Kernel.Font.PdfFontFactory.CreateFont();
                    
                    // Title
                    var title = $"EMBEDDED FILE ERROR";
                    canvas.SetFontAndSize(font, 20);
                    var titleWidth = font.GetWidth(title, 20);
                    canvas.SetTextMatrix(1, 0, 0, 1, (pageSize.GetWidth() - titleWidth) / 2, pageSize.GetHeight() - 150);
                    canvas.ShowText(title);
                    
                    // Type
                    var typeText = "Type: ERROR";
                    canvas.SetFontAndSize(font, 14);
                    var typeWidth = font.GetWidth(typeText, 14);
                    canvas.SetTextMatrix(1, 0, 0, 1, (pageSize.GetWidth() - typeWidth) / 2, pageSize.GetHeight() - 200);
                    canvas.ShowText(typeText);
                    
                    // File info
                    canvas.SetFontAndSize(font, 12);
                    canvas.SetTextMatrix(1, 0, 0, 1, 60, pageSize.GetHeight() - 250);
                    canvas.ShowText($"File: {fileName}");
                    
                    // Error message
                    canvas.SetTextMatrix(1, 0, 0, 1, 60, pageSize.GetHeight() - 280);
                    canvas.ShowText($"Error: {errorMessage}");
                }
                finally
                {
                    canvas.EndText();
                }
                      
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Inserted error placeholder for {fileName} at position {afterPageNumber + 1}");
                #endif
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[PDF-INSERT] Failed to insert error placeholder: {ex.Message}");
                #endif
            }
        }

        // --- Add helper for image to PDF conversion ---
        private static bool TryConvertImageToPdf(string imagePath, string outputPdfPath)
        {
            try
            {
                using (var writer = new iText.Kernel.Pdf.PdfWriter(outputPdfPath))
                using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                using (var docImg = new iText.Layout.Document(pdf))
                {
                    var img = new iText.Layout.Element.Image(iText.IO.Image.ImageDataFactory.Create(imagePath));
                    docImg.Add(img);
                }
                return File.Exists(outputPdfPath);
            }
            catch
            {
                return false;
            }
        }
    }
}
