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

namespace MsgToPdfConverter.Services
{
    public class AttachmentService
    {
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
                Console.WriteLine($"[HIERARCHY] Creating hierarchy diagram for: {currentItem}");
                Console.WriteLine($"[HIERARCHY] Parent chain: {string.Join(" -> ", parentChain ?? new List<string>())}");

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
                    Console.WriteLine($"[HIERARCHY] Successfully created hierarchy image: {imagePath}");
                    // Create PDF with the hierarchy image
                    PdfService.AddImagePdf(headerPdfPath, imagePath, headerText);

                    // Clean up the temporary image file
                    try
                    {
                        File.Delete(imagePath);
                    }
                    catch (Exception cleanupEx)
                    {
                        Console.WriteLine($"[HIERARCHY] Warning: Could not delete temporary image: {cleanupEx.Message}");
                    }
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[HIERARCHY] Failed to create hierarchy image, falling back to text: {ex.Message}");
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
                Console.WriteLine($"[HIERARCHY] Error creating hierarchy text, using original: {ex.Message}");
                return originalHeaderText;
            }
        }

        public void ProcessMsgAttachmentsRecursively(Storage.Message msg, List<string> allPdfFiles, List<string> allTempFiles, string tempDir, bool extractOriginalOnly, int depth = 0, int maxDepth = 5, string headerText = null, List<string> parentChain = null, Action progressTick = null)
        {
            if (depth > maxDepth)
            {
                Console.WriteLine($"[MSG] Max recursion depth {maxDepth} reached, skipping further processing");
                return;
            }
            if (parentChain == null)
            {
                parentChain = new List<string>();
            }
            if (depth > 0)
            {
                // For nested MSGs, process the body as a PDF (if not extractOriginalOnly)
                try
                {
                    if (!extractOriginalOnly)
                    {
                        var htmlResult = _emailService.BuildEmailHtmlWithInlineImages(msg, false);
                        string nestedHtmlPath = Path.Combine(tempDir, Guid.NewGuid() + "_nested.html");
                        File.WriteAllText(nestedHtmlPath, htmlResult.Html, System.Text.Encoding.UTF8);
                        allTempFiles.Add(nestedHtmlPath);
                        string nestedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_nested.pdf");
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
                            Console.WriteLine($"[MSG] Failed to convert nested MSG to PDF");
                        }
                        // Progress tick for MSG body
                        progressTick?.Invoke();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MSG] Error processing nested MSG body: {ex.Message}");
                }
            }

            // Now process attachments if they exist
            if (msg.Attachments == null || msg.Attachments.Count == 0)
            {
                Console.WriteLine($"[MSG] Depth {depth} - No attachments to process");
                return;
            }

            Console.WriteLine($"[MSG] Processing attachments at depth {depth}, found {msg.Attachments.Count} attachments");

            var inlineContentIds = _emailService.GetInlineContentIds(msg.BodyHtml ?? "");
            var typedAttachments = new List<Storage.Attachment>();
            var nestedMessages = new List<Storage.Message>();

            // Separate attachments and nested messages
            foreach (var att in msg.Attachments)
            {
                if (att is Storage.Attachment a)
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Examining attachment: {a.FileName} (IsInline: {a.IsInline}, ContentId: {a.ContentId})");

                    // Skip attachments if they have a ContentId that's actually referenced in the email body as an inline image
                    if (!string.IsNullOrEmpty(a.ContentId) && inlineContentIds.Contains(a.ContentId.Trim('<', '>', '\"', '\'', ' ')))
                    {
                        Console.WriteLine($"[MSG] Depth {depth} - Skipping inline attachment (referenced in email body): {a.FileName}");
                        continue;
                    }

                    // Skip small images that are likely signature images or decorative elements
                    if (IsLikelySignatureImage(a))
                    {
                        Console.WriteLine($"[MSG] Depth {depth} - Skipping likely signature/decorative image: {a.FileName}");
                        continue;
                    }

                    Console.WriteLine($"[MSG] Depth {depth} - Including attachment for processing: {a.FileName}");

                    typedAttachments.Add(a);
                }
                else if (att is Storage.Message nestedMsg)
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Found nested MSG: {nestedMsg.Subject ?? "No Subject"}");
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
                var attPdf = ProcessSingleAttachmentWithHierarchy(a, attPath, tempDir, attHeader, allTempFiles, parentChainForAtt, currentItem, extractOriginalOnly, progressTick);
                if (attPdf != null)
                {
                    allPdfFiles.Add(attPdf);
                }
                // Ensure progress tick for each attachment (if not handled inside)
                progressTick?.Invoke();
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

        public string ProcessSingleAttachment(Storage.Attachment att, string attPath, string tempDir, string headerText, List<string> allTempFiles)
        {
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
                            // Only call ExtractEmbeddedObjectsWithProgress here, do NOT call progressTick (not available in this method)
                            int embeddedCount = ExtractEmbeddedObjectsWithProgress(attPath, tempDir, allTempFiles, null);
                            // embeddedCount is the number of OLE/Package objects found and processed
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
                    finalAttachmentPdf = ProcessZipAttachmentWithHierarchy(attPath, tempDir, headerText, allTempFiles, new List<string>(), attName, false, null);
                    if (finalAttachmentPdf != null)
                    {
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else if (ext == ".7z")
                {
                    finalAttachmentPdf = Process7zAttachmentWithHierarchy(attPath, tempDir, headerText, allTempFiles, new List<string>(), attName, false, null);
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
                Console.WriteLine($"[ATTACH] Error processing attachment {attName}: {ex.Message}");
                finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Processing error: {ex.Message})", null);
                allTempFiles.Add(finalAttachmentPdf);
            }

            return finalAttachmentPdf;
        }

        public string ProcessZipAttachmentWithHierarchy(string attPath, string tempDir, string headerText, List<string> allTempFiles, List<string> parentChain, string currentItem, bool extractOriginalOnly = false, Action progressTick = null)
        {
            try
            {
                Console.WriteLine($"[ZIP] Processing ZIP file: {attPath}");

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
                            Console.WriteLine($"[ZIP] Found directory: {entry.FullName} - skipping");
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
                            Console.WriteLine($"[ZIP] Skipping likely signature image: {currentFileName}");
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
                                        Console.WriteLine($"[ZIP] Processing nested MSG with full recursion: {currentFileName}");

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
                                                Console.WriteLine($"[ZIP] Nested MSG produced single PDF: {entryPdf}");
                                            }
                                            else
                                            {
                                                // Multiple PDFs from nested MSG - merge them
                                                entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_nested_merged.pdf");
                                                _appendPdfs(nestedPdfFiles, entryPdf);
                                                Console.WriteLine($"[ZIP] Nested MSG produced {nestedPdfFiles.Count} PDFs, merged into: {entryPdf}");

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
                                    Console.WriteLine($"[ZIP] Error processing nested MSG {currentFileName}: {msgEx.Message}");
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
                            Console.WriteLine($"[ZIP] Error processing entry {entry.Name}: {entryEx.Message}");
                            entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_entry_error.pdf");
                            string errorFileName = Path.GetFileName(entry.FullName);
                            _addHeaderPdf(entryPdf, $"File: {errorFileName}\n(Processing error: {entryEx.Message})", null);
                            zipPdfFiles.Add(entryPdf);
                            allTempFiles.Add(entryPdf);
                            unconvertibleFiles.Add(errorFileName);
                        }

                        // No progress tick here - progress is handled at a higher level
                        // to avoid double-counting files inside ZIP archives
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
                        Console.WriteLine($"[ZIP] Created final merged PDF with {zipPdfFiles.Count} files: {finalZipPdf}");

                        // Add individual PDFs to cleanup since they're now merged into finalZipPdf
                        foreach (var pdf in zipPdfFiles)
                        {
                            if (File.Exists(pdf) && pdf != finalZipPdf)
                            {
                                try
                                {
                                    File.Delete(pdf);
                                    Console.WriteLine($"[ZIP] Cleaned up individual PDF: {pdf}");
                                }
                                catch (Exception cleanupEx)
                                {
                                    Console.WriteLine($"[ZIP] Warning - could not delete individual PDF {pdf}: {cleanupEx.Message}");
                                    allTempFiles.Add(pdf); // Add to cleanup list if manual delete failed
                                }
                            }
                        }

                        // DON'T add finalZipPdf to allTempFiles - it needs to be returned for the main output
                        return finalZipPdf;
                    }
                    else if (zipPdfFiles.Count == 1)
                    {
                        Console.WriteLine($"[ZIP] Returning single PDF: {zipPdfFiles[0]}");
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
                Console.WriteLine($"[ZIP] Error processing ZIP file {attPath}: {ex.Message}");
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
                Console.WriteLine($"[7Z] Processing 7z file: {attPath}");

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
                            Console.WriteLine($"[7Z] Found directory: {entry.Key} - skipping");
                            continue;
                        }
                        fileIndex++;

                        string currentFileName = Path.GetFileName(entry.Key);
                        string entryExt = Path.GetExtension(entry.Key).ToLowerInvariant();

                        // Skip signature images before extracting the file
                        if ((entryExt == ".jpg" || entryExt == ".jpeg" || entryExt == ".png" || entryExt == ".bmp" || entryExt == ".gif") &&
                            IsLikelySignatureImageByNameAndSize(currentFileName, entry.Size))
                        {
                            Console.WriteLine($"[7Z] Skipping likely signature image: {currentFileName}");
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
                                        Console.WriteLine($"[7Z] Processing nested MSG with full recursion: {currentFileName}");

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
                                                Console.WriteLine($"[7Z] Nested MSG produced single PDF: {entryPdf}");
                                            }
                                            else
                                            {
                                                // Multiple PDFs from nested MSG - merge them
                                                entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_nested_merged.pdf");
                                                _appendPdfs(nestedPdfFiles, entryPdf);
                                                Console.WriteLine($"[7Z] Nested MSG produced {nestedPdfFiles.Count} PDFs, merged into: {entryPdf}");

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
                                    Console.WriteLine($"[7Z] Error processing nested MSG {currentFileName}: {msgEx.Message}");
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
                            Console.WriteLine($"[7Z] Error processing entry {entry.Key}: {entryEx.Message}");
                            entryPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_entry_error.pdf");
                            string errorFileName = Path.GetFileName(entry.Key);
                            _addHeaderPdf(entryPdf, $"File: {errorFileName}\n(Processing error: {entryEx.Message})", null);
                            sevenZipPdfFiles.Add(entryPdf);
                            allTempFiles.Add(entryPdf);
                            unconvertibleFiles.Add(errorFileName);
                        }

                        // No progress tick here - progress is handled at a higher level
                        // to avoid double-counting files inside 7z archives
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
                        Console.WriteLine($"[7Z] Created final merged PDF with {sevenZipPdfFiles.Count} files: {final7zPdf}");

                        // Add individual PDFs to cleanup since they're now merged into final7zPdf
                        foreach (var pdf in sevenZipPdfFiles)
                        {
                            if (File.Exists(pdf) && pdf != final7zPdf)
                            {
                                try
                                {
                                    File.Delete(pdf);
                                    Console.WriteLine($"[7Z] Cleaned up individual PDF: {pdf}");
                                }
                                catch (Exception cleanupEx)
                                {
                                    Console.WriteLine($"[7Z] Warning - could not delete individual PDF {pdf}: {cleanupEx.Message}");
                                    allTempFiles.Add(pdf); // Add to cleanup list if manual delete failed
                                }
                            }
                        }

                        // DON'T add final7zPdf to allTempFiles - it needs to be returned for the main output
                        return final7zPdf;
                    }
                    else if (sevenZipPdfFiles.Count == 1)
                    {
                        Console.WriteLine($"[7Z] Returning single PDF: {sevenZipPdfFiles[0]}");
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
                Console.WriteLine($"[7Z] Error processing 7z file {attPath}: {ex.Message}");
                string errorPdf = Path.Combine(tempDir, Guid.NewGuid() + "_7z_error.pdf");
                _addHeaderPdf(errorPdf, $"7z Archive: {currentItem}\n(Processing error: {ex.Message})", null);
                allTempFiles.Add(errorPdf);
                return errorPdf;
            }
        }

        /// <summary>
        /// Processes a single attachment with SmartArt hierarchy support
        /// </summary>
        public string ProcessSingleAttachmentWithHierarchy(Storage.Attachment att, string attPath, string tempDir, string headerText, List<string> allTempFiles, List<string> parentChain, string currentItem, bool extractOriginalOnly = false, Action progressTick = null)
        {
            Console.WriteLine($"[ATTACH-DEBUG] ENTER: attName={att?.FileName}, attPath={attPath}, tempDir={tempDir}, headerText={headerText}, parentChain=[{string.Join(" -> ", parentChain ?? new List<string>())}], currentItem={currentItem}, extractOriginalOnly={extractOriginalOnly}");
            
            // Handle both .msg attachments and standalone files
            string attName;
            if (att != null)
            {
                attName = att.FileName ?? "attachment";
            }
            else
            {
                // For standalone files, use the file name from the path
                attName = Path.GetFileName(attPath) ?? "file";
            }
            
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
                    // CreateHierarchyHeaderPdf(parentChain, currentItem, headerText, headerPdf);
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
                        // --- Embedded OLE/Package extraction ---
                        if (File.Exists(attPath))
                        {
                            // Extract embedded objects (they now contribute to progress count)
                            int embeddedCount = ExtractEmbeddedObjectsWithProgress(attPath, tempDir, allTempFiles, progressTick);
                            // embeddedCount is the number of OLE/Package objects found and processed
                        }
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                        _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Conversion failed)", null);
                        allTempFiles.Add(finalAttachmentPdf);
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
                Console.WriteLine($"[ATTACH] Error processing attachment {attName}: {ex.Message}");
                finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                _addHeaderPdf(finalAttachmentPdf, $"File: {attName}\n(Processing error: {ex.Message})", null);
                allTempFiles.Add(finalAttachmentPdf);
            }

            Console.WriteLine($"[ATTACH-DEBUG] EXIT: attName={att?.FileName}, resultPdf={finalAttachmentPdf}");
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
                Console.WriteLine($"[MSG] Processing MSG file: {Path.GetFileName(msgFilePath)}");

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
                        Console.WriteLine($"[MSG] Successfully converted MSG to PDF: {Path.GetFileName(msgFilePath)} -> {nestedPdf}");
                        // Return the PDF path - caller is responsible for managing this file
                        return nestedPdf;
                    }
                    else
                    {
                        Console.WriteLine($"[MSG] Failed to convert MSG to PDF: {Path.GetFileName(msgFilePath)}");
                        return null;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[MSG] Error processing MSG file {Path.GetFileName(msgFilePath)}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Debug helper: Print the contents of the temp file list and files to protect before cleanup
        /// </summary>
        public static void DebugPrintTempAndProtectedFiles(IEnumerable<string> allTempFiles, IEnumerable<string> filesToProtect = null)
        {
            Console.WriteLine("[DEBUG] --- Temp file cleanup about to run ---");
            Console.WriteLine($"[DEBUG] allTempFiles ({(allTempFiles == null ? 0 : allTempFiles.Count())}):");
            if (allTempFiles != null)
            {
                foreach (var f in allTempFiles)
                {
                    Console.WriteLine($"[DEBUG]   TEMP: {f}");
                }
            }
            if (filesToProtect != null)
            {
                Console.WriteLine($"[DEBUG] filesToProtect ({filesToProtect.Count()}):");
                foreach (var f in filesToProtect)
                {
                    Console.WriteLine($"[DEBUG]   PROTECT: {f}");
                }
            }
            else
            {
                Console.WriteLine("[DEBUG] filesToProtect: (none provided)");
            }
            Console.WriteLine("[DEBUG] --- End of temp/protected file debug ---");
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
                    Console.WriteLine($"[FILTER] Detected signature image: {fileName} ({fileSizeKB}KB)");
                    return true;
                }

                // If it's marked as inline AND small, likely decorative/signature
                if (attachment.IsInline == true && isSmallImage)
                {
                    Console.WriteLine($"[FILTER] Detected small inline image: {fileName} ({fileSizeKB}KB)");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[FILTER] Error checking signature image {attachment.FileName}: {ex.Message}");
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
                    Console.WriteLine($"[FILTER] Detected signature image: {fileName} ({fileSizeKB}KB)");
                    return true;
                }

                // If it's very small (less than 10KB), likely decorative/signature
                if (fileSizeKB < 10)
                {
                    Console.WriteLine($"[FILTER] Detected very small image: {fileName} ({fileSizeKB}KB)");
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[FILTER] Error checking signature image {fileName}: {ex.Message}");
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
                        // Count nested MSG as 1 item (don't count its internal attachments)
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
                                            if (entryExt == ".msg")
                                            {
                                                // Count MSG files only, not their internal attachments
                                                count += 1;
                                            }
                                            else if (entryExt == ".zip" || entryExt == ".7z")
                                            {
                                                // Recurse into nested archive without double-counting
                                                try
                                                {
                                                    count += CountAllProcessableItemsFromFile(entryPath);
                                                }
                                                catch { count++; }
                                            }
                                            else
                                            {
                                                // Don't count regular files in ZIP - only MSG files matter for progress
                                                // count++; // Count each file as 1 item
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
                                            if (entryExt == ".msg")
                                            {
                                                // Count MSG files only, not their internal attachments
                                                count += 1;
                                            }
                                            else if (entryExt == ".zip" || entryExt == ".7z")
                                            {
                                                try
                                                {
                                                    count += CountAllProcessableItemsFromFile(entryPath);
                                                }
                                                catch { count++; }
                                            }
                                            else
                                            {
                                                // Don't count regular files in 7z - only MSG files matter for progress
                                                // count++; // Count each file as 1 item
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
                            // Don't count regular attachments - only count at the MSG level
                            // count++; // Count each attachment as 1 item (including Office files)
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
                // Count MSG files as 1 item each (don't count their internal attachments)
                return 1;
            }
            else if (ext == ".zip")
            {
                int count = 0;
                string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "MsgToPdf_CountAllItems_" + Guid.NewGuid().ToString());
                System.IO.Directory.CreateDirectory(tempDir);
                try
                {
                    using (var archive = System.IO.Compression.ZipFile.OpenRead(filePath))
                    {
                        foreach (var entry in archive.Entries)
                        {
                            if (entry.Length == 0) continue; // Skip directories
                            string entryPath = System.IO.Path.Combine(tempDir, entry.FullName);
                            System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(entryPath));
                            entry.ExtractToFile(entryPath, true);
                            string entryExt = System.IO.Path.GetExtension(entryPath).ToLowerInvariant();
                            if (entryExt == ".msg")
                            {
                                // Count MSG files only
                                count += 1;
                            }
                            else if (entryExt == ".zip" || entryExt == ".7z")
                            {
                                // For nested archives, count recursively
                                count += CountAllProcessableItemsFromFile(entryPath);
                            }
                            else
                            {
                                // Don't count regular files - only MSG files matter for progress
                                // count++; // Count each file as 1 item
                            }
                        }
                    }
                }
                catch { count++; }
                finally
                {
                    try { System.IO.Directory.Delete(tempDir, true); } catch { }
                }
                return count;
            }
            else if (ext == ".7z")
            {
                int count = 0;
                string tempDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "MsgToPdf_CountAllItems_" + Guid.NewGuid().ToString());
                System.IO.Directory.CreateDirectory(tempDir);
                try
                {
                    using (var archive = SharpCompress.Archives.SevenZip.SevenZipArchive.Open(filePath))
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
                            if (entryExt == ".msg")
                            {
                                // Count MSG files only
                                count += 1;
                            }
                            else if (entryExt == ".zip" || entryExt == ".7z")
                            {
                                // For nested archives, count recursively
                                count += CountAllProcessableItemsFromFile(entryPath);
                            }
                            else
                            {
                                // Don't count regular files - only MSG files matter for progress
                                // count++; // Count each file as 1 item
                            }
                        }
                    }
                }
                catch { count++; }
                finally
                {
                    try { System.IO.Directory.Delete(tempDir, true); } catch { }
                }
                return count;
            }
            else
            {
                // For regular files (including Office files), count as 1 item plus embedded objects
                int count = 1; // The main file itself
                
                // If it's an Office file, also count embedded objects
                string fileExt = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
                if (fileExt == ".doc" || fileExt == ".docx" || fileExt == ".xls" || fileExt == ".xlsx")
                {
                    try
                    {
                        int embeddedCount = CountEmbeddedObjects(filePath);
                        count += embeddedCount;
                        Console.WriteLine($"[COUNT] File {filePath}: 1 main + {embeddedCount} embedded = {count} total");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[COUNT] Error counting embedded objects in {filePath}: {ex.Message}");
                    }
                }
                
                return count;
            }
        }

        /// <summary>
        /// Counts embedded objects in an Office file without extracting them
        /// </summary>
        private static int CountEmbeddedObjects(string officeFilePath)
        {
            try
            {
                // Use DocxEmbeddedExtractor to get count of embedded files
                var embeddedFiles = MsgToPdfConverter.Utils.DocxEmbeddedExtractor.ExtractEmbeddedFiles(officeFilePath);
                return embeddedFiles?.Count ?? 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[COUNT-EMBED] Error counting embedded objects: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Extracts embedded OLE/Package objects from a DOCX/XLSX file and calls progressTick for each.
        /// </summary>
        private static int ExtractEmbeddedObjectsWithProgress(string officeFilePath, string tempDir, List<string> allTempFiles, Action progressTick)
        {
            int count = 0;
            try
            {
                // Use DocxEmbeddedExtractor to extract embedded files (OLE/Package objects)
                var embeddedFiles = MsgToPdfConverter.Utils.DocxEmbeddedExtractor.ExtractEmbeddedFiles(officeFilePath);
                if (embeddedFiles != null)
                {
                    foreach (var embedded in embeddedFiles)
                    {
                        // Save the embedded file to disk
                        string safeName = string.IsNullOrWhiteSpace(embedded.FileName) ? $"Embedded_{Guid.NewGuid()}" : embedded.FileName;
                        string outPath = Path.Combine(tempDir, safeName);
                        File.WriteAllBytes(outPath, embedded.Data);
                        allTempFiles.Add(outPath);
                        // Call progressTick for each embedded object since they are now counted
                        progressTick?.Invoke();
                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[OLE-EXTRACT] Error extracting embedded objects: {ex.Message}");
            }
            return count;
        }
    }
}
