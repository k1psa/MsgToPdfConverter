using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using MsgReader.Outlook;
using System.Diagnostics;
using MsgToPdfConverter.Utils;
using SharpCompress.Archives;
using SharpCompress.Archives.SevenZip;
using SharpCompress.Common;
using iText.Kernel.Pdf;

namespace MsgToPdfConverter.Services
{
    public class AttachmentService
    {
        public class EmbeddedObjectInfo
        {
            public string FilePath { get; set; }
            public string ObjectType { get; set; } // e.g. "Word", "Excel", "OLE", etc.
            public int InsertionIndex { get; set; } // Paragraph or page index after which to insert

            public EmbeddedObjectInfo(string filePath, string objectType, int insertionIndex)
            {
                FilePath = filePath;
                ObjectType = objectType;
                InsertionIndex = insertionIndex;
            }
        }
        /// <summary>
        /// Deletes all files and folders in %temp%\msgtopdf
        /// </summary>
        public static void CleanupMsgToPdfTempFolder()
        {
            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "MsgToPdfConverter");
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
             
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[CLEANUP] Error deleting temp folder: {ex.Message}");
                #endif
            }
        }
        private readonly Action<string, string, string> _addHeaderPdf;
        private readonly Func<string, string, bool> _tryConvertOfficeToPdf;
        private readonly Action<List<string>, string> _appendPdfs;
        private readonly EmailConverterService _emailService;

        public AttachmentService(Action<string, string, string> addHeaderPdf, Func<string, string, bool> tryConvertOfficeToPdf, Action<List<string>, string> appendPdfs, EmailConverterService emailService)
        {
            _addHeaderPdf = addHeaderPdf;
            _tryConvertOfficeToPdf = tryConvertOfficeToPdf;
            _appendPdfs = appendPdfs;
            _emailService = emailService;
        }

        /// <summary>
        /// Creates a header PDF with hierarchy diagram image (falls back to text-only if image creation fails)
        /// </summary>
        private void CreateHierarchyHeaderPdf(List<string> parentChain, string currentItem, string headerText, string headerPdfPath)
        {
            string imagePath = null;
            try
            {
                #if DEBUG
                DebugLogger.Log($"[HIERARCHY] Creating hierarchy diagram for: {currentItem}");
                #endif
                #if DEBUG
                DebugLogger.Log($"[HIERARCHY] Parent chain: {string.Join(" -> ", parentChain ?? new List<string>())}");
                #endif

                // Build full hierarchy chain including current item
                var fullChain = new List<string>();
                if (parentChain != null)
                    fullChain.AddRange(parentChain);
                fullChain.Add(currentItem);

                // Create hierarchy image
                var hierarchyImageService = new HierarchyImageService();
                string outputFolder = Path.GetDirectoryName(headerPdfPath);
                imagePath = hierarchyImageService.CreateHierarchyImage(fullChain, currentItem, outputFolder);

                if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                {
                    #if DEBUG
                    DebugLogger.Log($"[HIERARCHY] Successfully created hierarchy image: {imagePath}");
                    #endif
                    // Create PDF with the hierarchy image
                    PdfService.AddImagePdf(headerPdfPath, imagePath, headerText);

                    // Clean up the temporary image file
                    try
                    {
                        File.Delete(imagePath);
                    }
                    catch (Exception cleanupEx)
                    {
                        #if DEBUG
                        DebugLogger.Log($"[HIERARCHY] Warning: Could not delete temporary image: {cleanupEx.Message}");
                        #endif
                    }
                    return;
                }
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[HIERARCHY] Failed to create hierarchy image, falling back to text: {ex.Message}");
                #endif
            }

            // Clean up failed image file if it exists
            if (!string.IsNullOrEmpty(imagePath))
            {
                try
                {
                    if (File.Exists(imagePath))
                        File.Delete(imagePath);
                }
                catch { }
            }

            // Fall back to enhanced text header
            string enhancedHeader = CreateHierarchyHeaderText(parentChain, currentItem, headerText);
            _addHeaderPdf(headerPdfPath, enhancedHeader, null);
        }

        /// <summary>
        /// Creates a header text with hierarchy tree structure (fallback method)
        /// </summary>
        private string CreateHierarchyHeaderText(List<string> parentChain, string currentItem, string originalHeaderText)
        {
            try
            {
                // Build tree structure using TreeHeaderHelper
                string treeHeader = TreeHeaderHelper.BuildTreeHeader(parentChain, currentItem);

                // Combine original header with tree structure
                return $"{originalHeaderText}\n\n{treeHeader}";
            }
            catch (Exception ex)
            {
                #if DEBUG
                DebugLogger.Log($"[HIERARCHY] Error creating hierarchy text, using original: {ex.Message}");
                #endif
                return originalHeaderText;
            }
        }

        public void ProcessMsgAttachmentsRecursively(Storage.Message msg, List<string> allPdfFiles, List<string> allTempFiles, string tempDir, bool extractOriginalOnly, int depth = 0, int maxDepth = 5, string headerText = null, List<string> parentChain = null, Action progressTick = null)
        {
            if (msg == null)
            {
                #if DEBUG
                DebugLogger.Log($"[ERROR] Null Storage.Message passed to ProcessMsgAttachmentsRecursively.");
                #endif
                return;
            }
            if (allPdfFiles == null)
            {
                #if DEBUG
                DebugLogger.Log($"[ERROR] allPdfFiles is null in ProcessMsgAttachmentsRecursively.");
                #endif
                return;
            }
            if (allTempFiles == null)
            {
                #if DEBUG
                DebugLogger.Log($"[ERROR] allTempFiles is null in ProcessMsgAttachmentsRecursively.");
                #endif
                return;
            }
            if (tempDir == null)
            {
                #if DEBUG
                DebugLogger.Log($"[ERROR] tempDir is null in ProcessMsgAttachmentsRecursively.");
                #endif
                return;
            }

            if (depth > maxDepth)
            {
                #if DEBUG
                DebugLogger.Log($"[MSG] Max recursion depth {maxDepth} reached, skipping further processing");
                #endif
                return;
            }
            if (parentChain == null)
            {
                parentChain = new List<string>();
            }

            // Additional null checks for MSG properties
            if (msg.Subject == null)
            {
                #if DEBUG
                DebugLogger.Log("[DEBUG] MSG subject is null.");
                #endif
            }
            if (msg.Attachments == null)
            {
                #if DEBUG
                DebugLogger.Log("[DEBUG] MSG attachments collection is null.");
                #endif
            }
            if (msg.BodyText == null && msg.BodyHtml == null)
            {
                #if DEBUG
                DebugLogger.Log("[DEBUG] MSG body is null (both text and HTML).");
                #endif
            }

            if (depth > 0)
            {
                // For nested MSGs, process the body as a PDF (if not extractOriginalOnly)
                try
                {
                    if (!extractOriginalOnly)
                    {
                        // Use only the main temp folder for nested MSGs
                        var htmlResult = _emailService.BuildEmailHtmlWithInlineImages(msg, false);
                        if (string.IsNullOrEmpty(htmlResult.Html))
                        {
                            #if DEBUG
                            DebugLogger.Log("[DEBUG] htmlResult.Html is null or empty for nested MSG.");
                            #endif
                        }
                        string nestedHtmlPath = Path.Combine(Path.Combine(Path.GetTempPath(), "MsgToPdfConverter"), Guid.NewGuid() + "_nested.html");
                        File.WriteAllText(nestedHtmlPath, htmlResult.Html ?? string.Empty, System.Text.Encoding.UTF8);
                        allTempFiles.Add(nestedHtmlPath);
                        string nestedPdf = Path.Combine(Path.Combine(Path.GetTempPath(), "MsgToPdfConverter"), Guid.NewGuid() + "_nested.pdf");
                        var psi = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName,
                            Arguments = $"--html2pdf \"{nestedHtmlPath}\" \"{nestedPdf}\"",
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardOutput = true,
                            RedirectStandardError = true
                        };
                        var proc = System.Diagnostics.Process.Start(psi);
                        proc.WaitForExit();
                        if (proc.ExitCode == 0 && File.Exists(nestedPdf))
                        {
                            allPdfFiles.Add(nestedPdf);
                            allTempFiles.Add(nestedPdf);
                        }
                        else
                        {
                            #if DEBUG
                            DebugLogger.Log($"[MSG] Failed to convert nested MSG to PDF");
                            #endif
                        }
                        // Progress tick for MSG body
                        progressTick?.Invoke();
                    }
                }
                catch (Exception ex)
                {
                    #if DEBUG
                    DebugLogger.Log($"[MSG] Error processing nested MSG body: {ex.Message}");
                    #endif
                }
            }

            // Now process attachments if they exist
            if (msg.Attachments == null || msg.Attachments.Count == 0)
            {
                #if DEBUG
                DebugLogger.Log($"[MSG] Depth {depth} - No attachments to process");
                #endif
                return;
            }

            #if DEBUG
            DebugLogger.Log($"[MSG] Processing attachments at depth {depth}, found {msg.Attachments.Count} attachments");
            #endif

            var inlineContentIds = _emailService.GetInlineContentIds(msg.BodyHtml ?? "");
            var typedAttachments = new List<Storage.Attachment>();
            var nestedMessages = new List<Storage.Message>();

            // Separate attachments and nested messages + DEDUPLICATION LOGIC
            var allAttachments = new List<Storage.Attachment>();
            foreach (var att in msg.Attachments)
            {
                if (att is Storage.Attachment a)
                {
                    #if DEBUG
                    DebugLogger.Log($"[MSG] Depth {depth} - Examining attachment: {a.FileName} (IsInline: {a.IsInline}, ContentId: {a.ContentId})");
                    #endif

                    // Skip attachments if they have a ContentId that's actually referenced in the email body as an inline image
                    if (!string.IsNullOrEmpty(a.ContentId) && inlineContentIds.Contains(a.ContentId.Trim('<', '>', '\"', '\'', ' ')))
                    {
                        #if DEBUG
                        DebugLogger.Log($"[MSG] Depth {depth} - Skipping inline attachment (referenced in email body): {a.FileName}");
                        #endif
                        continue;
                    }

                    // Skip small images that are likely signature images or decorative elements
                    if (IsLikelySignatureImage(a))
                    {
                        #if DEBUG
                        DebugLogger.Log($"[MSG] Depth {depth} - Skipping likely signature/decorative image: {a.FileName}");
                        #endif
                        continue;
                    }

                    allAttachments.Add(a);
                }
            }

            // Deduplication only applies within the current attachment group, not across hierarchy boundaries.
            // If two files have the same name but come from different sources (e.g., one outside and one inside a nested MSG/ZIP/7z), both are processed independently.
            // Only skip true duplicates within the same attachment list (e.g., two identical files attached to the same email).
            var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var a in allAttachments)
            {
                if (a.FileName == null)
                {
                    typedAttachments.Add(a);
                    continue;
                }
                if (!seenNames.Contains(a.FileName))
                {
                    typedAttachments.Add(a);
                    seenNames.Add(a.FileName);
                }
                else
                {

#if DEBUG
                    DebugLogger.Log($"[MSG-DEDUP] Depth {depth} - SKIPPING true duplicate: {a.FileName}");
#endif
#if DEBUG
                    DebugLogger.Log($"[MSG-DEDUP] Depth {depth} - SKIPPING true duplicate: {a.FileName}");
#endif
                }
            }

            // Handle nested messages
            foreach (var att in msg.Attachments)
            {
                if (att is Storage.Message nestedMsg)
                {

#if DEBUG
                    DebugLogger.Log($"[MSG] Depth {depth} - Found nested MSG: {nestedMsg.Subject ?? "No Subject"}");
#endif
#if DEBUG
                    DebugLogger.Log($"[MSG] Depth {depth} - Found nested MSG: {nestedMsg.Subject ?? "No Subject"}");
#endif
                    nestedMessages.Add(nestedMsg);
                }
            }

            // Process regular attachments
            int totalAttachments = typedAttachments.Count;
            for (int attIndex = 0; attIndex < typedAttachments.Count; attIndex++)
            {
                var a = typedAttachments[attIndex];
                string attPath = Path.Combine(tempDir, a.FileName ?? $"attachment_{attIndex}");
                File.WriteAllBytes(attPath, a.Data);
                allTempFiles.Add(attPath);
                string attHeader = $"Attachment {attIndex + 1}/{totalAttachments} - {a.FileName}";
                string currentItem = a.FileName;
                var parentChainForAtt = new List<string>(parentChain);
                parentChainForAtt.Add(msg.Subject ?? "MSG");
                // Use hierarchy-aware processing and pass progressTick
                var attPdf = ProcessSingleAttachmentWithHierarchy(a, attPath, tempDir, attHeader, allTempFiles, allPdfFiles, parentChainForAtt, currentItem, extractOriginalOnly, progressTick);
                if (attPdf != null)
                {
                    allPdfFiles.Add(attPdf);
                }
                // Progress ticks are now handled inside ProcessSingleAttachmentWithHierarchy during embedding extraction
            }

            // Process nested MSG files recursively (this will handle both their body content and attachments)
            for (int msgIndex = 0; msgIndex < nestedMessages.Count; msgIndex++)
            {
                var nestedMsg = nestedMessages[msgIndex];
                var parentChainForMsg = new List<string>(parentChain);
                parentChainForMsg.Add(msg.Subject ?? "MSG");
                // Recursively process nested MSG, pass progressTick
                ProcessMsgAttachmentsRecursively(nestedMsg, allPdfFiles, allTempFiles, tempDir, extractOriginalOnly, depth + 1, maxDepth, $"Nested Email: {nestedMsg.Subject ?? "MSG"}", parentChainForMsg, progressTick);
            }
        }

        public string ProcessSingleAttachment(Storage.Attachment att, string attPath, string tempDir, string headerText, List<string> allTempFiles, List<string> allPdfFiles = null, Action progressTick = null)
        {
            if (att == null)
            {

#if DEBUG
                DebugLogger.Log($"[ERROR] Null Storage.Attachment passed to ProcessSingleAttachment.");
#endif
#if DEBUG
                DebugLogger.Log($"[ERROR] Null Storage.Attachment passed to ProcessSingleAttachment.");
#endif
                return null;
            }
            if (attPath == null)
            {

#if DEBUG
                DebugLogger.Log($"[ERROR] attPath is null in ProcessSingleAttachment.");
#endif
#if DEBUG
                DebugLogger.Log($"[ERROR] attPath is null in ProcessSingleAttachment.");
#endif
                return null;
            }
            if (tempDir == null)
            {

#if DEBUG
                DebugLogger.Log($"[ERROR] tempDir is null in ProcessSingleAttachment.");
#endif
#if DEBUG
                DebugLogger.Log($"[ERROR] tempDir is null in ProcessSingleAttachment.");
#endif
                return null;
            }
            if (allTempFiles == null)
            {

#if DEBUG
                DebugLogger.Log($"[ERROR] allTempFiles is null in ProcessSingleAttachment.");
#endif
#if DEBUG
                DebugLogger.Log($"[ERROR] allTempFiles is null in ProcessSingleAttachment.");
#endif
                return null;
            }

            string attName = att.FileName ?? "attachment";
            string ext = Path.GetExtension(attName).ToLowerInvariant();
            string attPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName) + ".pdf");
            string finalAttachmentPdf = null;

            try
            {
                if (ext == ".pdf")
                {
                    // Return PDF directly without header
                    finalAttachmentPdf = attPath;
                }
                else if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp" || ext == ".gif")
                {
                    // Create image-only PDF without header
                    // 1. Create header PDF (with hierarchy graphic/text)
                    // string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                    // _addHeaderPdf(headerPdf, headerText, null);
                    // 2. Create image-only PDF
                    string imagePdf = Path.Combine(tempDir, Guid.NewGuid() + "_image.pdf");
                    using (var writer = new iText.Kernel.Pdf.PdfWriter(imagePdf))
                    using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                    using (var docImg = new iText.Layout.Document(pdf))
                    {
                        var imgData = iText.IO.Image.ImageDataFactory.Create(attPath);
                        var image = new iText.Layout.Element.Image(imgData);
                        docImg.Add(image);
                    }
                    // 3. Merge header and image PDF
                    finalAttachmentPdf = imagePdf;
                    // finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                    // _appendPdfs(new List<string> { headerPdf, imagePdf }, finalAttachmentPdf);
                    // allTempFiles.Add(headerPdf);
                    // allTempFiles.Add(imagePdf);
                    allTempFiles.Add(finalAttachmentPdf);
                }
                else if (ext == ".doc" || ext == ".docx" || ext == ".xls" || ext == ".xlsx")
                {
                    if (_tryConvertOfficeToPdf(attPath, attPdf))
                    {
                        // Return converted PDF directly without header
                        finalAttachmentPdf = attPdf;
                        allTempFiles.Add(attPdf);
                        // --- Embedded OLE/Package extraction progress ---
                        if (File.Exists(attPath))
                        {
                            int embeddedCount = ExtractEmbeddedObjectsWithProgress(attPath, tempDir, allTempFiles, allPdfFiles, progressTick);
                        }
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                        _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Conversion failed)", null);
                        allTempFiles.Add(finalAttachmentPdf);
                        // No progressTick here
                    }
                }
                else if (ext == ".zip")
                {
                    finalAttachmentPdf = ProcessZipAttachmentWithHierarchy(attPath, tempDir, headerText, allTempFiles, new List<string>(), attName, false, progressTick);
                    if (finalAttachmentPdf != null)
                    {
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else if (ext == ".7z")
                {
                    finalAttachmentPdf = Process7zAttachmentWithHierarchy(attPath, tempDir, headerText, allTempFiles, new List<string>(), attName, false, progressTick);
                    if (finalAttachmentPdf != null)
                    {
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else
                {
                    finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                    _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Unsupported file type)", null);
                    allTempFiles.Add(finalAttachmentPdf);
                }
            }
            catch (Exception ex)
            {

#if DEBUG
                DebugLogger.Log($"[ATTACH] Error processing attachment {attName}: {ex.Message}");
#endif
#if DEBUG
                DebugLogger.Log($"[ATTACH] Error processing attachment {attName}: {ex.Message}");
#endif
                finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Processing error: {ex.Message})", null);
                allTempFiles.Add(finalAttachmentPdf);
            }

            return finalAttachmentPdf;
        }

        public string ProcessZipAttachmentWithHierarchy(string attPath, string tempDir, string headerText, List<string> allTempFiles, List<string> parentChain, string currentItem, bool extractOriginalOnly = false, Action progressTick = null)
        {
            if (attPath == null)
            {

#if DEBUG
                DebugLogger.Log($"[ERROR] attPath is null in ProcessZipAttachmentWithHierarchy.");
#endif
#if DEBUG
                DebugLogger.Log($"[ERROR] attPath is null in ProcessZipAttachmentWithHierarchy.");
#endif
                return null;
            }
            if (tempDir == null)
            {

#if DEBUG
                DebugLogger.Log($"[ERROR] tempDir is null in ProcessZipAttachmentWithHierarchy.");
#endif
#if DEBUG
                DebugLogger.Log($"[ERROR] tempDir is null in ProcessZipAttachmentWithHierarchy.");
#endif
                return null;
            }
            if (allTempFiles == null)
            {

#if DEBUG
                DebugLogger.Log($"[ERROR] allTempFiles is null in ProcessZipAttachmentWithHierarchy.");
#endif
#if DEBUG
                DebugLogger.Log($"[ERROR] allTempFiles is null in ProcessZipAttachmentWithHierarchy.");
#endif
                return null;
            }
            if (parentChain == null)
            {

#if DEBUG
                DebugLogger.Log($"[ERROR] parentChain is null in ProcessZipAttachmentWithHierarchy.");
#endif
#if DEBUG
                DebugLogger.Log($"[ERROR] parentChain is null in ProcessZipAttachmentWithHierarchy.");
#endif
                return null;
            }

            try
            {

#if DEBUG
                DebugLogger.Log($"[ZIP] Processing ZIP file: {attPath}");
#endif
#if DEBUG
                DebugLogger.Log($"[ZIP] Processing ZIP file: {attPath}");
#endif

                // Track files that could not be converted
                var unconvertibleFiles = new List<string>();

                // Create enhanced header text with hierarchy
                string enhancedHeaderText = CreateHierarchyHeaderText(parentChain, currentItem, headerText);

                using (var archive = System.IO.Compression.ZipFile.OpenRead(attPath))
                {
                    var zipPdfFiles = new List<string>();

                    // Only count files (not directories) for the header
                    var fileEntries = archive.Entries.Where(e => e.Length > 0).ToList();
                    int fileCount = fileEntries.Count;
                    var folderEntries = archive.Entries.Where(e => e.Length == 0).ToList();
                    int folderCount = folderEntries.Count;

                    int fileIndex = 0;
                    int folderIndex = 0;
                    int totalEntries = archive.Entries.Count;

                    foreach (var entry in archive.Entries)
                    {
                        // Build comprehensive parent chain for ZIP entries including folder structure
                        var zipEntryParentChain = new List<string>(parentChain);
                        zipEntryParentChain.Add(currentItem);

                        // For nested folder structures, add each folder level to the parent chain
                        var pathParts = entry.FullName.Split('/', '\\');
                        for (int i = 0; i < pathParts.Length - 1; i++) // Exclude the filename itself
                        {
                            if (!string.IsNullOrEmpty(pathParts[i]))
                            {
                                zipEntryParentChain.Add(pathParts[i] + "/");
                            }
                        }

                        if (entry.Length == 0)
                        {
                            folderIndex++;
                            // This is a directory - skip it entirely (no header creation)

#if DEBUG
                            DebugLogger.Log($"[ZIP] Found directory: {entry.FullName} - skipping");
#endif
#if DEBUG
                            DebugLogger.Log($"[ZIP] Found directory: {entry.FullName} - skipping");
#endif
                            continue;
                        }
                        fileIndex++;

                        // Get the final filename for the current item in hierarchy
                        string currentFileName = Path.GetFileName(entry.FullName);
                        string entryExt = Path.GetExtension(entry.Name).ToLowerInvariant();

                        // Skip signature images before extracting the file
                        if ((entryExt == ".jpg" || entryExt == ".jpeg" || entryExt == ".png" || entryExt == ".bmp" || entryExt == ".gif") &&
                            IsLikelySignatureImageByNameAndSize(currentFileName, entry.Length))
                        {

#if DEBUG
                            DebugLogger.Log($"[ZIP] Skipping likely signature image: {currentFileName}");
#endif
#if DEBUG
                            DebugLogger.Log($"[ZIP] Skipping likely signature image: {currentFileName}");
#endif
                            continue; // Skip this image file entirely - don't extract or process
                        }

                        string entryPath = Path.Combine(tempDir, $"zip_{Guid.NewGuid()}_{Path.GetFileName(entry.Name)}");
                        entry.ExtractToFile(entryPath, true);
                        allTempFiles.Add(entryPath);

                        string entryPdf = null;

                        try
                        {
                            if (entryExt == ".pdf")
                            {
                                // Return PDF directly without header
                                entryPdf = entryPath;
                            }
                            else if (entryExt == ".jpg" || entryExt == ".jpeg" || entryExt == ".png" || entryExt == ".bmp" || entryExt == ".gif")
                            {
                                // Create image-only PDF without header
                                // 1. Create header PDF (with hierarchy graphic/text)
                                // string entryHeaderPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_entry_header.pdf");
                                // CreateHierarchyHeaderPdf(zipEntryParentChain, currentFileName, $"Attachment {fileIndex}/{fileCount} - {currentFileName}", entryHeaderPdf);
                                // 2. Create image-only PDF
                                string imagePdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_image.pdf");
                                using (var writer = new iText.Kernel.Pdf.PdfWriter(imagePdf))
                                using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                                using (var docImg = new iText.Layout.Document(pdf))
                                {
                                    var imgData = iText.IO.Image.ImageDataFactory.Create(entryPath);
                                    var image = new iText.Layout.Element.Image(imgData);
                                    docImg.Add(image);
                                }
                                // 3. Merge header and image PDF
                                entryPdf = imagePdf;
                                // entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_entry_merged.pdf");
                                // _appendPdfs(new List<string> { entryHeaderPdf, imagePdf }, entryPdf);
                                // allTempFiles.Add(entryHeaderPdf);
                                // allTempFiles.Add(imagePdf);
                                allTempFiles.Add(entryPdf);
                            }
                            else if (entryExt == ".doc" || entryExt == ".docx" || entryExt == ".xls" || entryExt == ".xlsx")
                            {
                                string convertedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_converted.pdf");
                                if (_tryConvertOfficeToPdf(entryPath, convertedPdf))
                                {
                                    // Return converted PDF directly without header
                                    entryPdf = convertedPdf;
                                    allTempFiles.Add(convertedPdf);
                                    // string entryHeaderPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_entry_header.pdf");
                                    // CreateHierarchyHeaderPdf(zipEntryParentChain, currentFileName, $"Attachment {fileIndex}/{fileCount} - {currentFileName}", entryHeaderPdf);
                                    // entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_entry_merged.pdf");
                                    // _appendPdfs(new List<string> { entryHeaderPdf, convertedPdf }, entryPdf);
                                    // allTempFiles.Add(entryHeaderPdf);
                                    // allTempFiles.Add(convertedPdf);
                                    // allTempFiles.Add(entryPdf);
                                }
                                else
                                {
                                    // Create simple text PDF for conversion failure
                                    entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_entry_placeholder.pdf");
                                    _addHeaderPdf(entryPdf, $"File: {currentFileName}\n(Conversion failed)", null);
                                    allTempFiles.Add(entryPdf);
                                    unconvertibleFiles.Add(currentFileName);
                                }
                            }
                            else if (entryExt == ".msg")
                            {
                                // Handle nested MSG files in ZIP with full recursive processing (including attachments)
                                try
                                {
                                    using (var nestedMsg = new Storage.Message(entryPath))
                                    {

#if DEBUG
                                        DebugLogger.Log($"[ZIP] Processing nested MSG with full recursion: {currentFileName}");
#endif
#if DEBUG
                                        DebugLogger.Log($"[ZIP] Processing nested MSG with full recursion: {currentFileName}");
#endif

                                        // Create a temporary list to collect all PDFs from this nested MSG
                                        var nestedPdfFiles = new List<string>();
                                        var nestedTempFiles = new List<string>();

                                        // Process the MSG recursively (this will handle the email body + all attachments based on extractOriginalOnly flag)
                                        ProcessMsgAttachmentsRecursively(nestedMsg, nestedPdfFiles, nestedTempFiles, tempDir, extractOriginalOnly, 1, 5,
                                            $"Nested Email from ZIP: {nestedMsg.Subject ?? currentFileName}",
                                            new List<string>(zipEntryParentChain));

                                        // Add all temp files to our main cleanup list
                                        allTempFiles.AddRange(nestedTempFiles);

                                        if (nestedPdfFiles.Count > 0)
                                        {
                                            if (nestedPdfFiles.Count == 1)
                                            {
                                                // Single PDF from nested MSG
                                                entryPdf = nestedPdfFiles[0];

#if DEBUG
                                                DebugLogger.Log($"[ZIP] Nested MSG produced single PDF: {entryPdf}");
#endif
#if DEBUG
                                                DebugLogger.Log($"[ZIP] Nested MSG produced single PDF: {entryPdf}");
#endif
                                            }
                                            else
                                            {
                                                // Multiple PDFs from nested MSG - merge them
                                                entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_nested_merged.pdf");
                                                _appendPdfs(nestedPdfFiles, entryPdf);

#if DEBUG
                                                DebugLogger.Log($"[ZIP] Nested MSG produced {nestedPdfFiles.Count} PDFs, merged into: {entryPdf}");
#endif
#if DEBUG
                                                DebugLogger.Log($"[ZIP] Nested MSG produced {nestedPdfFiles.Count} PDFs, merged into: {entryPdf}");
#endif

                                                // Clean up individual PDFs after merging
                                                foreach (var pdf in nestedPdfFiles)
                                                {
                                                    try
                                                    {
                                                        if (File.Exists(pdf) && pdf != entryPdf)
                                                        {
                                                            File.Delete(pdf);
                                                        }
                                                    }
                                                    catch { } // Ignore cleanup errors
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // No PDFs produced - create error PDF
                                            entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_msg_error.pdf");
                                            _addHeaderPdf(entryPdf, $"File: {currentFileName}\n(No content could be extracted)", null);
                                            allTempFiles.Add(entryPdf);
                                        }
                                    }
                                }
                                catch (Exception msgEx)
                                {

#if DEBUG
                                    DebugLogger.Log($"[ZIP] Error processing nested MSG {currentFileName}: {msgEx.Message}");
#endif
#if DEBUG
                                    DebugLogger.Log($"[ZIP] Error processing nested MSG {currentFileName}: {msgEx.Message}");
#endif
                                    entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_msg_error.pdf");
                                    _addHeaderPdf(entryPdf, $"File: {currentFileName}\n(MSG processing error: {msgEx.Message})", null);
                                    allTempFiles.Add(entryPdf);
                                }
                                
                                // Progress tick for MSG file processing
                                progressTick?.Invoke();
                            }
                            else
                            {
                                // Create simple text PDF for unsupported file types
                                entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_entry_placeholder.pdf");
                                _addHeaderPdf(entryPdf, $"File: {currentFileName}\n(Unsupported file type: {entryExt})", null);
                                allTempFiles.Add(entryPdf);
                                unconvertibleFiles.Add(currentFileName);
                            }

                            if (entryPdf != null)
                                zipPdfFiles.Add(entryPdf);
                        }
                        catch (Exception entryEx)
                        {

#if DEBUG
                            DebugLogger.Log($"[ZIP] Error processing entry {entry.Name}: {entryEx.Message}");
#endif
#if DEBUG
                            DebugLogger.Log($"[ZIP] Error processing entry {entry.Name}: {entryEx.Message}");
#endif
                            entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_entry_error.pdf");
                            string errorFileName = Path.GetFileName(entry.FullName);
                            _addHeaderPdf(entryPdf, $"File: {errorFileName}\n(Processing error: {entryEx.Message})", null);
                            zipPdfFiles.Add(entryPdf);
                            allTempFiles.Add(entryPdf);
                            unconvertibleFiles.Add(errorFileName);
                        }

                        // Progress tick for every file processed in ZIP
                        progressTick?.Invoke();
                    }

                    // Skip unconvertible files notification (no header creation)
                    // if (unconvertibleFiles.Count > 0)
                    // {
                    //     string notifyText = "WARNING: The following files could not be converted to PDF and are not included as content pages:\n" + string.Join("\n", unconvertibleFiles) + "\n\n";
                    //     string newHeaderPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_header_notify.pdf");
                    //     CreateHierarchyHeaderPdf(parentChain, currentItem, notifyText + enhancedHeaderText + $"\n\nZIP Archive Contents ({fileCount} files):", newHeaderPdf);
                    //     zipPdfFiles[0] = newHeaderPdf;
                    //     allTempFiles.Add(newHeaderPdf);
                    // }

                    // Merge all ZIP entry PDFs (if any)
                    if (zipPdfFiles.Count > 1)
                    {
                        string finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_final.pdf");
                        _appendPdfs(zipPdfFiles, finalZipPdf);

#if DEBUG
                        DebugLogger.Log($"[ZIP] Created final merged PDF with {zipPdfFiles.Count} files: {finalZipPdf}");
#endif
#if DEBUG
                        DebugLogger.Log($"[ZIP] Created final merged PDF with {zipPdfFiles.Count} files: {finalZipPdf}");
#endif

                        // Add individual PDFs to cleanup since they're now merged into finalZipPdf
                        foreach (var pdf in zipPdfFiles)
                        {
                            if (File.Exists(pdf) && pdf != finalZipPdf)
                            {
                                try
                                {
                                    File.Delete(pdf);

#if DEBUG
                                    DebugLogger.Log($"[ZIP] Cleaned up individual PDF: {pdf}");
#endif
#if DEBUG
                                    DebugLogger.Log($"[ZIP] Cleaned up individual PDF: {pdf}");
#endif
                                }
                                catch
                                {
                                    allTempFiles.Add(pdf); // Add to cleanup list if manual delete failed
                                }
                            }
                        }

                        // DON'T add finalZipPdf to allTempFiles - it needs to be returned for the main output
                        return finalZipPdf;
                    }
                    else if (zipPdfFiles.Count == 1)
                    {

                        // Don't add to allTempFiles - it needs to be returned for the main output
                        return zipPdfFiles[0];
                    }
                    else
                    {
                        // No files processed - create simple text PDF
                        string emptyPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_empty.pdf");
                        _addHeaderPdf(emptyPdf, $"ZIP Archive: {currentItem}\n(No supported files found)", null);
                        allTempFiles.Add(emptyPdf);
                        return emptyPdf;
                    }
                }
            }
            catch (Exception ex)
            {

                string errorPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_error.pdf");
                _addHeaderPdf(errorPdf, $"ZIP Archive: {currentItem}\n(Processing error: {ex.Message})", null);
                allTempFiles.Add(errorPdf);
                return errorPdf;
            }
        }

        public string Process7zAttachmentWithHierarchy(string attPath, string tempDir, string headerText, List<string> allTempFiles, List<string> parentChain, string currentItem, bool extractOriginalOnly = false, Action progressTick = null)
        {
            try
            {


                // Track files that could not be converted
                var unconvertibleFiles = new List<string>();

                // Create enhanced header text with hierarchy
                string enhancedHeaderText = CreateHierarchyHeaderText(parentChain, currentItem, headerText);

                using (var archive = SevenZipArchive.Open(attPath))
                {
                    var sevenZipPdfFiles = new List<string>();

                    // Only count files (not directories) for the header
                    var fileEntries = archive.Entries.Where(e => !e.IsDirectory).ToList();
                    int fileCount = fileEntries.Count;
                    var folderEntries = archive.Entries.Where(e => e.IsDirectory).ToList();
                    int folderCount = folderEntries.Count;

                    int fileIndex = 0;
                    int folderIndex = 0;
                    int totalEntries = archive.Entries.Count();

                    foreach (var entry in archive.Entries)
                    {
                        // Build comprehensive parent chain for 7z entries including folder structure
                        var sevenZipEntryParentChain = new List<string>(parentChain);
                        sevenZipEntryParentChain.Add(currentItem);

                        // For nested folder structures, add each folder level to the parent chain
                        var pathParts = entry.Key.Split('/', '\\');
                        for (int i = 0; i < pathParts.Length - 1; i++) // Exclude the filename itself
                        {
                            if (!string.IsNullOrEmpty(pathParts[i]))
                            {
                                sevenZipEntryParentChain.Add(pathParts[i] + "/");
                            }
                        }

                        if (entry.IsDirectory)
                        {
                            folderIndex++;
                            // This is a directory - skip it entirely (no header creation)
#if DEBUG
                            DebugLogger.Log($"[7Z] Found directory: {entry.Key} - skipping");
#endif
                            continue;
                        }
                        fileIndex++;

                        string currentFileName = Path.GetFileName(entry.Key);
                        string entryExt = Path.GetExtension(entry.Key).ToLowerInvariant();

                        // Skip signature images before extracting the file
                        if ((entryExt == ".jpg" || entryExt == ".jpeg" || entryExt == ".png" || entryExt == ".bmp" || entryExt == ".gif") &&
                            IsLikelySignatureImageByNameAndSize(currentFileName, entry.Size))
                        {

#if DEBUG
                            DebugLogger.Log($"[7Z] Skipping likely signature image: {currentFileName}");
#endif
                            continue;
                        }

                        string entryPath = Path.Combine(tempDir, $"7z_{Guid.NewGuid()}_{Path.GetFileName(entry.Key)}");
                        using (var entryStream = entry.OpenEntryStream())
                        using (var fileStream = File.Create(entryPath))
                        {
                            entryStream.CopyTo(fileStream);
                        }
                        allTempFiles.Add(entryPath);

                        string entryPdf = null;

                        try
                        {
                            if (entryExt == ".pdf")
                            {
                                // Return PDF directly without header
                                entryPdf = entryPath;
                            }
                            else if (entryExt == ".jpg" || entryExt == ".jpeg" || entryExt == ".png" || entryExt == ".bmp" || entryExt == ".gif")
                            {
                                // Create image-only PDF without header
                                // 1. Create header PDF (with hierarchy graphic/text)
                                // string entryHeaderPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_entry_header.pdf");
                                // CreateHierarchyHeaderPdf(sevenZipEntryParentChain, currentFileName, $"Attachment {fileIndex}/{fileCount} - {currentFileName}", entryHeaderPdf);
                                // 2. Create image-only PDF
                                string imagePdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_image.pdf");
                                using (var writer = new iText.Kernel.Pdf.PdfWriter(imagePdf))
                                using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                                using (var docImg = new iText.Layout.Document(pdf))
                                {
                                    var imgData = iText.IO.Image.ImageDataFactory.Create(entryPath);
                                    var image = new iText.Layout.Element.Image(imgData);
                                    docImg.Add(image);
                                }
                                // 3. Merge header and image PDF
                                entryPdf = imagePdf;
                                // entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_entry_merged.pdf");
                                // _appendPdfs(new List<string> { entryHeaderPdf, imagePdf }, entryPdf);
                                // allTempFiles.Add(entryHeaderPdf);
                                // allTempFiles.Add(imagePdf);
                                allTempFiles.Add(entryPdf);
                            }
                            else if (entryExt == ".doc" || entryExt == ".docx" || entryExt == ".xls" || entryExt == ".xlsx")
                            {
                                string convertedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_converted.pdf");
                                if (_tryConvertOfficeToPdf(entryPath, convertedPdf))
                                {
                                    // Return converted PDF directly without header
                                    entryPdf = convertedPdf;
                                    allTempFiles.Add(convertedPdf);
                                    // string entryHeaderPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_entry_header.pdf");
                                    // CreateHierarchyHeaderPdf(sevenZipEntryParentChain, currentFileName, $"Attachment {fileIndex}/{fileCount} - {currentFileName}", entryHeaderPdf);
                                    // entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_entry_merged.pdf");
                                    // _appendPdfs(new List<string> { entryHeaderPdf, convertedPdf }, entryPdf);
                                    // allTempFiles.Add(entryHeaderPdf);
                                    // allTempFiles.Add(convertedPdf);
                                    // allTempFiles.Add(entryPdf);
                                }
                                else
                                {
                                    // Create simple text PDF for conversion failure
                                    entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_entry_placeholder.pdf");
                                    _addHeaderPdf(entryPdf, $"File: {currentFileName}\n(Conversion failed)", null);
                                    allTempFiles.Add(entryPdf);
                                    unconvertibleFiles.Add(currentFileName);
                                }
                            }
                            else if (entryExt == ".msg")
                            {
                                // Handle nested MSG files in 7z with full recursive processing (including attachments)
                                try
                                {
                                    using (var nestedMsg = new Storage.Message(entryPath))
                                    {

#if DEBUG
                                        DebugLogger.Log($"[7Z] Processing nested MSG with full recursion: {currentFileName}");
#endif

                                        // Create a temporary list to collect all PDFs from this nested MSG
                                        var nestedPdfFiles = new List<string>();
                                        var nestedTempFiles = new List<string>();

                                        // Process the MSG recursively (this will handle the email body + all attachments based on extractOriginalOnly flag)
                                        ProcessMsgAttachmentsRecursively(nestedMsg, nestedPdfFiles, nestedTempFiles, tempDir, extractOriginalOnly, 1, 5,
                                            $"Nested Email from 7z: {nestedMsg.Subject ?? currentFileName}",
                                            new List<string>(sevenZipEntryParentChain));

                                        // Add all temp files to our main cleanup list
                                        allTempFiles.AddRange(nestedTempFiles);

                                        if (nestedPdfFiles.Count > 0)
                                        {
                                            if (nestedPdfFiles.Count == 1)
                                            {
                                                // Single PDF from nested MSG
                                                entryPdf = nestedPdfFiles[0];

#if DEBUG
                                                DebugLogger.Log($"[7Z] Nested MSG produced single PDF: {entryPdf}");
#endif
                                            }
                                            else
                                            {
                                                // Multiple PDFs from nested MSG - merge them
                                                entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_nested_merged.pdf");
                                                _appendPdfs(nestedPdfFiles, entryPdf);

#if DEBUG
                                                DebugLogger.Log($"[7Z] Nested MSG produced {nestedPdfFiles.Count} PDFs, merged into: {entryPdf}");
#endif

                                                // Clean up individual PDFs after merging
                                                foreach (var pdf in nestedPdfFiles)
                                                {
                                                    try
                                                    {
                                                        if (File.Exists(pdf) && pdf != entryPdf)
                                                        {
                                                            File.Delete(pdf);
                                                        }
                                                    }
                                                    catch { } // Ignore cleanup errors
                                                }
                                            }
                                        }
                                        else
                                        {
                                            // No PDFs produced - create error PDF
                                            entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_msg_error.pdf");
                                            _addHeaderPdf(entryPdf, $"File: {currentFileName}\n(No content could be extracted)", null);
                                            allTempFiles.Add(entryPdf);
                                        }
                                    }
                                }
                                catch (Exception msgEx)
                                {

#if DEBUG
                                    DebugLogger.Log($"[7Z] Error processing nested MSG {currentFileName}: {msgEx.Message}");
#endif
                                    entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_msg_error.pdf");
                                    _addHeaderPdf(entryPdf, $"File: {currentFileName}\n(MSG processing error: {msgEx.Message})", null);
                                    allTempFiles.Add(entryPdf);
                                }
                                
                                // Progress tick for MSG file processing
                                progressTick?.Invoke();
                            }
                            else
                            {
                                // Create simple text PDF for unsupported file types
                                entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_entry_placeholder.pdf");
                                _addHeaderPdf(entryPdf, $"File: {currentFileName}\n(Unsupported file type: {entryExt})", null);
                                allTempFiles.Add(entryPdf);
                                unconvertibleFiles.Add(currentFileName);
                            }

                            if (entryPdf != null)
                                sevenZipPdfFiles.Add(entryPdf);
                        }
                        catch (Exception entryEx)
                        {

#if DEBUG
                            DebugLogger.Log($"[7Z] Error processing entry {entry.Key}: {entryEx.Message}");
#endif
                            entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_entry_error.pdf");
                            string errorFileName = Path.GetFileName(entry.Key);
                            _addHeaderPdf(entryPdf, $"File: {errorFileName}\n(Processing error: {entryEx.Message})", null);
                            sevenZipPdfFiles.Add(entryPdf);
                            allTempFiles.Add(entryPdf);
                            unconvertibleFiles.Add(errorFileName);
                        }

                        // Progress tick for every file processed in 7z
                        progressTick?.Invoke();
                    }

                    // Skip unconvertible files notification (no header creation)
                    // if (unconvertibleFiles.Count > 0)
                    // {
                    //     string notifyText = "WARNING: The following files could not be converted to PDF and are not included as content pages:\n" + string.Join("\n", unconvertibleFiles) + "\n\n";
                    //     string newHeaderPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_header_notify.pdf");
                    //     CreateHierarchyHeaderPdf(parentChain, currentItem, notifyText + enhancedHeaderText + $"\n\n7z Archive Contents ({fileCount} files):", newHeaderPdf);
                    //     sevenZipPdfFiles[0] = newHeaderPdf;
                    //     allTempFiles.Add(newHeaderPdf);
                    // }

                    // Merge all 7z entry PDFs (if any)
                    if (sevenZipPdfFiles.Count > 1)
                    {
                        string final7zPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_final.pdf");
                        _appendPdfs(sevenZipPdfFiles, final7zPdf);

#if DEBUG
                        DebugLogger.Log($"[7Z] Created final merged PDF with {sevenZipPdfFiles.Count} files: {final7zPdf}");
#endif

                        // Add individual PDFs to cleanup since they're now merged into final7zPdf
                        foreach (var pdf in sevenZipPdfFiles)
                        {
                            if (File.Exists(pdf) && pdf != final7zPdf)
                            {
                                try
                                {
                                    File.Delete(pdf);

#if DEBUG
                                    DebugLogger.Log($"[7Z] Cleaned up individual PDF: {pdf}");
#endif
                                }
                                catch
                                {
                                    allTempFiles.Add(pdf); // Add to cleanup list if manual delete failed
                                }
                            }
                        }

                        // DON'T add final7zPdf to allTempFiles - it needs to be returned for the main output
                        return final7zPdf;
                    }
                    else if (sevenZipPdfFiles.Count == 1)
                    {

                        // Don't add to allTempFiles - it needs to be returned for the main output
                        return sevenZipPdfFiles[0];
                    }
                    else
                    {
                        // No files processed - create simple text PDF
                        string emptyPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_empty.pdf");
                        _addHeaderPdf(emptyPdf, $"7z Archive: {currentItem}\n(No supported files found)", null);
                        allTempFiles.Add(emptyPdf);
                        return emptyPdf;
                    }
                }
            }
            catch (Exception ex)
            {

                string errorPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_error.pdf");
                _addHeaderPdf(errorPdf, $"7z Archive: {currentItem}\n(Processing error: {ex.Message})", null);
                allTempFiles.Add(errorPdf);
                return errorPdf;
            }
        }

        /// <summary>
        /// Processes a single attachment with SmartArt hierarchy support
        /// </summary>
        public string ProcessSingleAttachmentWithHierarchy(Storage.Attachment att, string attPath, string tempDir, string headerText, List<string> allTempFiles, List<string> allPdfFiles, List<string> parentChain, string currentItem, bool extractOriginalOnly = false, Action progressTick = null)
        {

            
            // Handle both .msg attachments and standalone files
            string attName;
            if (att != null)
            {
                attName = att.FileName ?? "attachment";
            }
            else
            {
                attName = Path.GetFileName(attPath) ?? "file";
            }

            string ext = Path.GetExtension(attName).ToLowerInvariant();
            string attPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName) + ".pdf");
            string finalAttachmentPdf = null;

            try
            {
                // Ensure tempDir exists before any file operations
                if (!Directory.Exists(tempDir))
                {
                    Directory.CreateDirectory(tempDir);
                }
                if (ext == ".pdf")
                {
                    // If dropped file is a PDF, copy it to temp/output folder as a new file
                    string outputPdf = Path.Combine(tempDir, Guid.NewGuid() + "_copy.pdf");
                    if (File.Exists(attPath))
                    {
                        File.Copy(attPath, outputPdf, true);
                        finalAttachmentPdf = outputPdf;
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                        _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Source PDF not found)", null);
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else if (ext == ".msg")
                {
                    // If dropped file is a MSG, process it as a nested message
                    if (File.Exists(attPath))
                    {
                        using (var nestedMsg = new Storage.Message(attPath))
                        {
                            var nestedPdfFiles = new List<string>();
                            var nestedTempFiles = new List<string>();
                            ProcessMsgAttachmentsRecursively(nestedMsg, nestedPdfFiles, nestedTempFiles, tempDir, extractOriginalOnly, 1, 5, $"Nested Email: {nestedMsg.Subject ?? attName}", new List<string>(parentChain));
                            allTempFiles.AddRange(nestedTempFiles);
                            if (nestedPdfFiles.Count > 0)
                            {
                                if (nestedPdfFiles.Count == 1)
                                {
                                    finalAttachmentPdf = nestedPdfFiles[0];
                                }
                                else
                                {
                                    finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_msg_merged.pdf");
                                    _appendPdfs(nestedPdfFiles, finalAttachmentPdf);
                                    allTempFiles.Add(finalAttachmentPdf);
                                }
                            }
                            else
                            {
                                finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_msg_error.pdf");
                                _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(No content could be extracted)", null);
                                allTempFiles.Add(finalAttachmentPdf);
                            }
                        }
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_msg_error.pdf");
                        _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Source MSG not found)", null);
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp" || ext == ".gif")
                {
                    if (File.Exists(attPath))
                    {
                        string imagePdf = Path.Combine(tempDir, Guid.NewGuid() + "_image.pdf");
                        using (var writer = new iText.Kernel.Pdf.PdfWriter(imagePdf))
                        using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                        using (var docImg = new iText.Layout.Document(pdf))
                        {
                            var imgData = iText.IO.Image.ImageDataFactory.Create(attPath);
                            var image = new iText.Layout.Element.Image(imgData);
                            docImg.Add(image);
                        }
                        finalAttachmentPdf = imagePdf;
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_image_error.pdf");
                        _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Source image not found)", null);
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else if (ext == ".doc" || ext == ".docx" || ext == ".xls" || ext == ".xlsx")
                {
                    if (progressTick != null)
                    {
                        PdfEmbeddedInsertionService.SetProgressCallback(progressTick);
                    }
                    if (File.Exists(attPath))
                    {
                        if (_tryConvertOfficeToPdf(attPath, attPdf))
                        {
                            finalAttachmentPdf = attPdf;
                            allTempFiles.Add(attPdf);
                            int embeddedCount = ExtractEmbeddedObjectsWithProgress(attPath, tempDir, allTempFiles, allPdfFiles, progressTick);
                            if (embeddedCount == 0 && progressTick != null)
                            {
                                progressTick();
                            }
                        }
                        else
                        {
                            finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                            _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Conversion failed)", null);
                            allTempFiles.Add(finalAttachmentPdf);
                        }
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_office_error.pdf");
                        _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Source Office file not found)", null);
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                    PdfEmbeddedInsertionService.SetProgressCallback(null);
                }
                else if (ext == ".zip")
                {
                    if (File.Exists(attPath))
                    {
                        finalAttachmentPdf = ProcessZipAttachmentWithHierarchy(attPath, tempDir, headerText, allTempFiles, new List<string>(), attName, false, progressTick);
                        if (finalAttachmentPdf != null)
                        {
                            allTempFiles.Add(finalAttachmentPdf);
                        }
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_error.pdf");
                        _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Source ZIP not found)", null);
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else if (ext == ".7z")
                {
                    if (File.Exists(attPath))
                    {
                        finalAttachmentPdf = Process7zAttachmentWithHierarchy(attPath, tempDir, headerText, allTempFiles, new List<string>(), attName, false, progressTick);
                        if (finalAttachmentPdf != null)
                        {
                            allTempFiles.Add(finalAttachmentPdf);
                        }
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_error.pdf");
                        _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Source 7z not found)", null);
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else
                {
                    finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                    _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Unsupported file type)", null);
                    allTempFiles.Add(finalAttachmentPdf);
                }
            }
            catch (Exception ex)
            {

                finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Processing error: {ex.Message})", null);
                allTempFiles.Add(finalAttachmentPdf);
            }


            return finalAttachmentPdf;
        }

        /// <summary>
        /// Unified method for processing MSG files - converts MSG to PDF using consistent logic
        /// </summary>
        /// <param name="msgFilePath">Path to the MSG file to process</param>
        /// <param name="tempDir">Temporary directory for intermediate files</param>
        /// <param name="allTempFiles">List to track temporary files for cleanup</param>
        /// <returns>Path to the generated PDF, or null if conversion failed</returns>
        private string ProcessMsgToPdf(string msgFilePath, string tempDir, List<string> allTempFiles)
        {
            try
            {


                using (var nestedMsg = new Storage.Message(msgFilePath))
                {
                    // Convert nested MSG to HTML and then PDF
                    var htmlResult = _emailService.BuildEmailHtmlWithInlineImages(nestedMsg, false);
                    string nestedHtmlPath = Path.Combine(tempDir, Guid.NewGuid() + "_nested.html");
                    File.WriteAllText(nestedHtmlPath, htmlResult.Html, System.Text.Encoding.UTF8);
                    allTempFiles.Add(nestedHtmlPath); // Always add HTML to temp files for cleanup

                    string nestedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_nested.pdf");

                    // Use HtmlToPdfWorker for conversion
                    var psi = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName,
                        Arguments = $"--html2pdf \"{nestedHtmlPath}\" \"{nestedPdf}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true
                    };

                    var proc = System.Diagnostics.Process.Start(psi);
                    proc.WaitForExit();

                    if (proc.ExitCode == 0 && File.Exists(nestedPdf))
                    {

                        // Return the PDF path - caller is responsible for managing this file
                        return nestedPdf;
                    }
                    else
                    {
#if DEBUG
                        DebugLogger.Log($"[MSG] Failed to convert MSG to PDF: {Path.GetFileName(msgFilePath)}");
#endif
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                DebugLogger.Log($"[MSG] Error processing MSG file {Path.GetFileName(msgFilePath)}: {ex.Message}");
#endif
                return null;
            }
        }

        /// <summary>
        /// Debug helper: Print the contents of the temp file list and files to protect before cleanup
        /// </summary>
        public static void DebugPrintTempAndProtectedFiles(IEnumerable<string> allTempFiles, IEnumerable<string> filesToProtect = null)
        {
#if DEBUG
            DebugLogger.Log("[DEBUG] --- Temp file cleanup about to run ---");
#endif
#if DEBUG
            DebugLogger.Log($"[DEBUG] allTempFiles ({(allTempFiles == null ? 0 : allTempFiles.Count())}):");
#endif
            if (allTempFiles != null)
            {
                foreach (var f in allTempFiles)
                {
#if DEBUG
                    DebugLogger.Log($"[DEBUG]   TEMP: {f}");
#endif
                }
            }
            if (filesToProtect != null)
            {
#if DEBUG
                DebugLogger.Log($"[DEBUG] filesToProtect ({filesToProtect.Count()}):");
#endif
                foreach (var f in filesToProtect)
                {
#if DEBUG
                    DebugLogger.Log($"[DEBUG]   PROTECT: {f}");
#endif
                }
            }
            else
            {
#if DEBUG
                DebugLogger.Log("[DEBUG] filesToProtect: (none provided)");
#endif
            }
#if DEBUG
            DebugLogger.Log("[DEBUG] --- End of temp/protected file debug ---");
#endif
        }

        /// <summary>
        /// Determines if an attachment is likely a signature image or decorative element that should be skipped
        /// </summary>
        public bool IsLikelySignatureImage(Storage.Attachment attachment)
        {
            try
            {
                string fileName = attachment.FileName ?? "";
                string ext = Path.GetExtension(fileName).ToLowerInvariant();

                // Only check image files
                if (ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" && ext != ".bmp")
                {
                    return false; // Not an image, so not a signature image
                }

                // Check file size - signature images are typically small (less than 50KB)
                int fileSizeKB = (attachment.Data?.Length ?? 0) / 1024;
                bool isSmallImage = fileSizeKB < 50;

                // Check for common signature image patterns in filename
                string lowerFileName = fileName.ToLowerInvariant();
                bool hasSignaturePattern = lowerFileName.Contains("image") ||
                                         lowerFileName.Contains("signature") ||
                                         lowerFileName.Contains("logo") ||
                                         lowerFileName.Contains("banner") ||
                                         lowerFileName.StartsWith("oledata.mso");

                // If it's a small image with signature patterns, likely a signature
                if (isSmallImage && hasSignaturePattern)
                {
#if DEBUG
                    DebugLogger.Log($"[FILTER] Detected signature image: {fileName} ({fileSizeKB}KB)");
#endif
                    return true;
                }

                // If it's marked as inline AND small, likely decorative/signature
                if (attachment.IsInline == true && isSmallImage)
                {
#if DEBUG
                    DebugLogger.Log($"[FILTER] Detected small inline image: {fileName} ({fileSizeKB}KB)");
#endif
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
#if DEBUG
                DebugLogger.Log($"[FILTER] Error checking signature image {attachment.FileName}: {ex.Message}");
#endif
                return false; // If in doubt, don't filter out
            }
        }

        /// <summary>
        /// Determines if a file is likely a signature image based on name and size (for archive processing)
        /// </summary>
        private bool IsLikelySignatureImageByNameAndSize(string fileName, long fileSize)
        {
            try
            {
                string ext = Path.GetExtension(fileName).ToLowerInvariant();

                // Only check image files
                if (ext != ".jpg" && ext != ".jpeg" && ext != ".png" && ext != ".gif" && ext != ".bmp")
                {
                    return false; // Not an image, so not a signature image
                }

                // Check file size - signature images are typically small (less than 50KB)
                int fileSizeKB = (int)(fileSize / 1024);
                bool isSmallImage = fileSizeKB < 50;

                // Check for common signature image patterns in filename
                string lowerFileName = fileName.ToLowerInvariant();
                bool hasSignaturePattern = lowerFileName.Contains("image") ||
                                         lowerFileName.Contains("signature") ||
                                         lowerFileName.Contains("logo") ||
                                         lowerFileName.Contains("banner") ||
                                         lowerFileName.StartsWith("oledata.mso");

                // If it's a small image with signature patterns, likely a signature
                if (isSmallImage && hasSignaturePattern)
                {
#if DEBUG
                    DebugLogger.Log($"[FILTER] Detected signature image: {fileName} ({fileSizeKB}KB)");
#endif
                    return true;
                }

                // If it's very small (less than 10KB), likely decorative/signature
                if (fileSizeKB < 10)
                {
#if DEBUG
                    DebugLogger.Log($"[FILTER] Detected very small image: {fileName} ({fileSizeKB}KB)");
#endif
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
#if DEBUG
                DebugLogger.Log($"[FILTER] Error checking signature image {fileName}: {ex.Message}");
#endif
                return false; // If in doubt, don't filter out
            }
        }

        /// <summary>
        /// Recursively counts all processable items (MSG files, regular files in ZIP/7z)
        /// This method counts only top-level user-visible files to match progress reporting
        /// </summary>
        public int CountAllProcessableItems(Storage.Message msg)
        {
            int count = 0;
            if (msg.Attachments != null)
            {
                foreach (var att in msg.Attachments)
                {
                    if (att is Storage.Message nestedMsg)
                    {
                        // Count nested MSG as 1 item
                        count += 1;
                    }
                    else if (att is Storage.Attachment a)
                    {
                        string ext = System.IO.Path.GetExtension(a.FileName ?? "").ToLowerInvariant();
                        if (ext == ".zip" || ext == ".7z")
                        {
                            // Extract to temp dir and count recursively
                            string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "MsgToPdf_CountAllItems_" + Guid.NewGuid().ToString());
                            System.IO.Directory.CreateDirectory(tempDir);
                            try
                            {
                                if (ext == ".zip")
                                {
                                    using (var archive = System.IO.Compression.ZipFile.OpenRead(a.FileName))
                                    {
                                        foreach (var entry in archive.Entries)
                                        {
                                            if (entry.Length == 0) continue; // skip directories
                                            string entryPath = System.IO.Path.Combine(tempDir, entry.FullName);
                                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(entryPath));
                                            entry.ExtractToFile(entryPath, true);
                                            string entryExt = System.IO.Path.GetExtension(entryPath).ToLowerInvariant();
                                            if (entryExt == ".zip" || entryExt == ".7z")
                                            {
                                                try
                                                {
                                                    count += CountAllProcessableItemsFromFile(entryPath);
                                                }
                                                catch { count++; }
                                            }
                                            else
                                            {
                                                // Count every file (not just MSG)
                                                count++;
                                            }
                                        }
                                    }
                                }
                                else if (ext == ".7z")
                                {
                                    using (var archive = SharpCompress.Archives.SevenZip.SevenZipArchive.Open(a.FileName))
                                    {
                                        foreach (var entry in archive.Entries)
                                        {
                                            if (entry.IsDirectory) continue;
                                            string entryPath = System.IO.Path.Combine(tempDir, entry.Key);
                                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(entryPath));
                                            using (var entryStream = entry.OpenEntryStream())
                                            using (var fileStream = System.IO.File.Create(entryPath))
                                            {
                                                entryStream.CopyTo(fileStream);
                                            }
                                            string entryExt = System.IO.Path.GetExtension(entryPath).ToLowerInvariant();
                                            if (entryExt == ".zip" || entryExt == ".7z")
                                            {
                                                try
                                                {
                                                    count += CountAllProcessableItemsFromFile(entryPath);
                                                }
                                                catch { count++; }
                                            }
                                            else
                                            {
                                                // Count every file (not just MSG)
                                                count++;
                                            }
                                        }
                                    }
                                }
                            }
                            catch { count++; }
                            finally
                            {
                                try { System.IO.Directory.Delete(tempDir, true); } catch { }
                            }
                        }
                        else
                        {
                            // Count every regular attachment (not just MSG)
                            count++;

                            // If Office file, count embedded files too
                            if (ext == ".doc" || ext == ".docx" || ext == ".xls" || ext == ".xlsx")
                            {
                                try
                                {
                                    var embeddedFiles = MsgToPdfConverter.Utils.DocxEmbeddedExtractor.ExtractEmbeddedFiles(a.FileName);
                                    if (embeddedFiles != null)
                                        count += embeddedFiles.Count;
                                }
                                catch { }
                            }
                        }
                    }
                }
            }
            return count;
        }

        /// <summary>
        /// Helper to count all processable items from a file (MSG, ZIP, 7z)
        /// This method counts only MSG files to match progress reporting
        /// </summary>
        public int CountAllProcessableItemsFromFile(string filePath)
        {
            string ext = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
            if (ext == ".msg")
            {
                try
                {
                    using (var msg = new Storage.Message(filePath))
                    {
                        return CountAllProcessableItems(msg);
                    }
                }
                catch { return 1; }
            }
            else if (ext == ".zip")
            {
                int zipCount = 0;
                string zipTempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "MsgToPdf_CountAllItems_" + Guid.NewGuid().ToString());
                System.IO.Directory.CreateDirectory(zipTempDir);
                try
                {
                    using (var zipArchive = System.IO.Compression.ZipFile.OpenRead(filePath))
                    {
                        foreach (var zipEntry in zipArchive.Entries)
                        {
                            if (zipEntry.Length == 0) continue;
                            string zipEntryPath = System.IO.Path.Combine(zipTempDir, zipEntry.FullName);
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(zipEntryPath));
                            zipEntry.ExtractToFile(zipEntryPath, true);
                            string zipEntryExt = System.IO.Path.GetExtension(zipEntryPath).ToLowerInvariant();
                            if (zipEntryExt == ".zip" || zipEntryExt == ".7z")
                            {
                                try
                                {
                                    zipCount += CountAllProcessableItemsFromFile(zipEntryPath);
                                }
                                catch { zipCount++; }
                            }
                            else
                            {
                                // Count every file (not just MSG)
                                zipCount++;
                            }
                        }
                    }
                }
                catch { zipCount++; }
                finally
                {
                    try { System.IO.Directory.Delete(zipTempDir, true); } catch { }
                }
                return zipCount;
            }
            else if (ext == ".7z")
            {
                int sevenCount = 0;
                string sevenTempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "MsgToPdf_CountAllItems_" + Guid.NewGuid().ToString());
                System.IO.Directory.CreateDirectory(sevenTempDir);
                try
                {
                    using (var sevenArchive = SharpCompress.Archives.SevenZip.SevenZipArchive.Open(filePath))
                    {
                        foreach (var sevenEntry in sevenArchive.Entries)
                        {
                            if (sevenEntry.IsDirectory) continue;
                            string sevenEntryPath = System.IO.Path.Combine(sevenTempDir, sevenEntry.Key);
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(sevenEntryPath));
                            using (var entryStream = sevenEntry.OpenEntryStream())
                            using (var fileStream = System.IO.File.Create(sevenEntryPath))
                            {
                                entryStream.CopyTo(fileStream);
                            }
                            string sevenEntryExt = System.IO.Path.GetExtension(sevenEntryPath).ToLowerInvariant();
                            if (sevenEntryExt == ".zip" || sevenEntryExt == ".7z")
                            {
                                try
                                {
                                    sevenCount += CountAllProcessableItemsFromFile(sevenEntryPath);
                                }
                                catch { sevenCount++; }
                            }
                            else
                            {
                                // Count every file (not just MSG)
                                sevenCount++;
                            }
                        }
                    }
                }
                catch { sevenCount++; }
                finally
                {
                    try { System.IO.Directory.Delete(sevenTempDir, true); } catch { }
                }
                return sevenCount;
            }
            else
            {
                int total = 1;
                if (ext == ".docx" || ext == ".xlsx")
                {
                    int embedCount = 0;
                    try
                    {
                        using (var archive = System.IO.Compression.ZipFile.OpenRead(filePath))
                        {
                            foreach (var entry in archive.Entries)
                            {
                                if ((ext == ".docx" && entry.FullName.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase)) ||
                                    (ext == ".xlsx" && entry.FullName.StartsWith("xl/embeddings/", StringComparison.OrdinalIgnoreCase)))
                                {
                                    // Count all files in the embeddings folder, regardless of extension (including .bin)
                                    embedCount++;
                                }
                            }
                        }
                    }
                    catch { }
                    total = 1 + 2 * embedCount;
                }
                else if (ext == ".doc" || ext == ".xls")
                {
                    // Only count embedded files using DocxEmbeddedExtractor for doc/xls
                    try
                    {
                        var embeddedFiles = MsgToPdfConverter.Utils.DocxEmbeddedExtractor.ExtractEmbeddedFiles(filePath);
                        if (embeddedFiles != null && embeddedFiles.Count > 0)
                            total += embeddedFiles.Count;
                    }
                    catch { }
                }
                return total;
            }
        }

        /// <summary>
        /// Extracts embedded OLE/Package objects from a DOCX/XLSX file. Progress ticks happen during PDF insertion.
        /// </summary>
        private int ExtractEmbeddedObjectsWithProgress(string officeFilePath, string tempDir, List<string> allTempFiles, List<string> allPdfFiles, Action progressTick)
        {
            int count = 0;
            try
            {
                // Extract embedded Office files from DOCX (ZIP) 'embeddings' folder
                if (Path.GetExtension(officeFilePath).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                {
                    using (var archive = ZipFile.OpenRead(officeFilePath))
                    {
                        foreach (var entry in archive.Entries)
                        {
                            // Only process files in 'word/embeddings/'
                            if (entry.FullName.StartsWith("word/embeddings/", StringComparison.OrdinalIgnoreCase))
                            {
                                // Use unified app temp directory for all temp files
                                string outDir = MsgToPdfConverter.Services.PdfEmbeddedInsertionService.AppTempDir;
                                Directory.CreateDirectory(outDir);
                                string outPath = Path.Combine(outDir, entry.Name);
                                entry.ExtractToFile(outPath, true);
                                allTempFiles.Add(outPath);
                                // Convert to PDF using existing delegate
                                string pdfPath = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(entry.Name) + ".pdf");
                                if (_tryConvertOfficeToPdf != null && _tryConvertOfficeToPdf(outPath, pdfPath))
                                {
                                    allTempFiles.Add(pdfPath);
                                    // Do NOT add to allPdfFiles! Embedded PDFs should only be inserted at mapped positions, not merged at start/end.
                                }
                                // Always tick for each file in the embeddings folder (regardless of extension)
                                progressTick?.Invoke();
                                count++;
                            }
                        }
                    }
                }
                // Fallback to DocxEmbeddedExtractor for other embedded files
                var embeddedFiles = MsgToPdfConverter.Utils.DocxEmbeddedExtractor.ExtractEmbeddedFiles(officeFilePath);
                if (embeddedFiles != null)
                {
                    foreach (var embedded in embeddedFiles)
                    {
                        string safeName = string.IsNullOrWhiteSpace(embedded.FileName) ? $"Embedded_{Guid.NewGuid()}" : embedded.FileName;
                        string outPath = Path.Combine(tempDir, safeName);
                        File.WriteAllBytes(outPath, embedded.Data);
                        allTempFiles.Add(outPath);
                        // Always tick for each file extracted here (regardless of extension)
                        progressTick?.Invoke();
                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                DebugLogger.Log($"[OLE-EXTRACT] Error extracting embedded objects: {ex.Message}");
#endif
            }
            return count;
        }
    }
}
