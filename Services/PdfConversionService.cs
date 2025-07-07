using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DinkToPdf;
using DinkToPdf.Contracts;
using MsgReader.Outlook;

namespace MsgToPdfConverter.Services
{
    public static class PdfConversionService
    {
        /// <summary>
        /// Kill any lingering wkhtmltopdf processes
        /// </summary>
        public static void KillWkhtmltopdfProcesses()
        {
            try
            {
                var procs = Process.GetProcessesByName("wkhtmltopdf");
                foreach (var proc in procs)
                {
                    try { proc.Kill(); } catch { }
                }
                Console.WriteLine($"Killed {procs.Length} lingering wkhtmltopdf processes.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error killing wkhtmltopdf processes: {ex.Message}");
            }
        }

        /// <summary>
        /// Configure DinkToPdf to find wkhtmltopdf binaries
        /// </summary>
        public static void ConfigureDinkToPdfPath(PdfTools pdfTools)
        {
            try
            {
                // Try to find wkhtmltopdf binaries in various locations
                string appDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string architecture = Environment.Is64BitProcess ? "x64" : "x86";

                // Check if architecture folder exists in the same directory as exe
                string archPath = Path.Combine(appDir, architecture);
                if (Directory.Exists(archPath))
                {
                    Console.WriteLine($"[DEBUG] Found architecture folder: {archPath}");
                    return; // DinkToPdf should find it automatically
                }

                // Check if architecture folder exists in libraries subfolder
                string librariesArchPath = Path.Combine(appDir, "libraries", architecture);
                if (Directory.Exists(librariesArchPath))
                {
                    Console.WriteLine($"[DEBUG] Found architecture folder in libraries: {librariesArchPath}");
                    // Copy the architecture folder to the main directory temporarily
                    string tempArchPath = Path.Combine(appDir, architecture);
                    if (!Directory.Exists(tempArchPath))
                    {
                        Console.WriteLine($"[DEBUG] Copying {librariesArchPath} to {tempArchPath}");
                        FileService.DirectoryCopy(librariesArchPath, tempArchPath, true);
                    }
                    return;
                }

                Console.WriteLine("[DEBUG] No wkhtmltopdf architecture folder found");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Error configuring DinkToPdf path: {ex.Message}");
            }
        }

        /// <summary>
        /// Run DinkToPdf conversion in STA thread
        /// </summary>
        public static void RunDinkToPdfConversion(HtmlToPdfDocument doc)
        {
            Exception threadEx = null;
            Console.WriteLine("[DEBUG] About to create STA thread for DinkToPdf");
            var staThread = new System.Threading.Thread(() =>
            {
                try
                {
                    Console.WriteLine("[DEBUG] Inside STA thread: Killing lingering wkhtmltopdf processes");
                    KillWkhtmltopdfProcesses();
                    Console.WriteLine("[DEBUG] Inside STA thread: Creating SynchronizedConverter");

                    // Configure DinkToPdf to use the correct path for wkhtmltopdf binaries
                    var pdfTools = new PdfTools();
                    ConfigureDinkToPdfPath(pdfTools);

                    var converter = new SynchronizedConverter(pdfTools);
                    Console.WriteLine("[DEBUG] Inside STA thread: Starting converter.Convert");
                    converter.Convert(doc);
                    Console.WriteLine("[DEBUG] Inside STA thread: Finished converter.Convert");
                    KillWkhtmltopdfProcesses();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception ex)
                {
                    threadEx = ex;
                    Console.WriteLine($"[DINKTOPDF] Exception: {ex}");
                }
            });
            staThread.SetApartmentState(System.Threading.ApartmentState.STA);
            Console.WriteLine("[DEBUG] About to start STA thread");
            staThread.Start();
            Console.WriteLine("[DEBUG] Waiting for STA thread to finish");
            staThread.Join();
            Console.WriteLine("[DEBUG] STA thread finished");
            if (threadEx != null)
                throw new Exception("DinkToPdf conversion failed", threadEx);
        }

        /// <summary>
        /// Appends attachments as PDFs to the main email PDF
        /// </summary>
        public static void AppendAttachmentsToPdf(string mainPdfPath, List<Storage.Attachment> attachments, SynchronizedConverter converter, bool deleteSourcePdf)
        {
            if (attachments == null || attachments.Count == 0)
                return;

            var tempPdfFiles = new List<string>();
            string tempDir = Path.GetTempPath();

            foreach (var att in attachments)
            {
                try
                {
                    string attName = att.FileName ?? "unnamed_attachment";
                    Console.WriteLine($"[ATTACH] Processing: {attName}");

                    // Save attachment to temp file
                    string attPath = Path.Combine(tempDir, Guid.NewGuid() + "_" + Path.GetFileName(attName));
                    string attPdf = Path.Combine(tempDir, Guid.NewGuid() + "_attachment.pdf");

                    using (var fileStream = new FileStream(attPath, FileMode.Create))
                    {
                        fileStream.Write(att.Data, 0, att.Data.Length);
                    }

                    string ext = Path.GetExtension(attName).ToLowerInvariant();
                    string headerText = $"Attachment: {attName}";

                    if (ext == ".pdf")
                    {
                        string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                        PdfService.AddHeaderPdf(headerPdf, headerText);
                        string mergedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                        PdfAppendTest.AppendPdfs(new List<string> { headerPdf, attPath }, mergedPdf);
                        tempPdfFiles.Add(mergedPdf);
                        File.Delete(headerPdf);
                    }
                    else if (ext == ".doc" || ext == ".docx" || ext == ".xls" || ext == ".xlsx")
                    {
                        if (OfficeConversionService.TryConvertOfficeToPdf(attPath, attPdf))
                        {
                            string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                            PdfService.AddHeaderPdf(headerPdf, headerText);
                            string mergedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                            PdfAppendTest.AppendPdfs(new List<string> { headerPdf, attPdf }, mergedPdf);
                            tempPdfFiles.Add(mergedPdf);
                            File.Delete(headerPdf);
                            File.Delete(attPdf);
                        }
                        else
                        {
                            PdfService.AddPlaceholderPdf(attPdf, $"Could not convert attachment: {attName}");
                            tempPdfFiles.Add(attPdf);
                        }
                    }
                    else if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".gif" || ext == ".bmp")
                    {
                        using (var writer = new iText.Kernel.Pdf.PdfWriter(attPdf))
                        using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                        using (var docImg = new iText.Layout.Document(pdf))
                        {
                            var p = new iText.Layout.Element.Paragraph(headerText)
                                .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                                .SetFontSize(16);
                            docImg.Add(p);
                            var imgData = iText.IO.Image.ImageDataFactory.Create(attPath);
                            var image = new iText.Layout.Element.Image(imgData);
                            docImg.Add(image);
                        }
                        tempPdfFiles.Add(attPdf);
                    }
                    else if (ext == ".zip")
                    {
                        using (var archive = System.IO.Compression.ZipFile.OpenRead(attPath))
                        {
                            var zipPdfFiles = new List<string>();
                            foreach (var entry in archive.Entries)
                            {
                                if (entry.FullName.EndsWith("/") || string.IsNullOrEmpty(entry.Name))
                                    continue; // Skip directories

                                // Sanitize entry.Name to prevent path traversal
                                string safeEntryName = Path.GetFileName(entry.Name);
                                string zfExt = Path.GetExtension(safeEntryName).ToLowerInvariant();
                                string zf = Path.Combine(tempDir, Guid.NewGuid() + "_" + safeEntryName);
                                string zfPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zipfile.pdf");

                                // Validate that zf is within tempDir to prevent path traversal
                                if (!Path.GetFullPath(zf).StartsWith(Path.GetFullPath(tempDir), StringComparison.OrdinalIgnoreCase))
                                    throw new InvalidOperationException("Unsafe file path detected.");

                                using (var entryStream = entry.Open())
                                using (var outputStream = new FileStream(zf, FileMode.Create))
                                {
                                    entryStream.CopyTo(outputStream);
                                }

                                if (zfExt == ".pdf")
                                {
                                    string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                                    PdfService.AddHeaderPdf(headerPdf, $"{headerText} (ZIP: {safeEntryName})");
                                    string mergedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                                    PdfAppendTest.AppendPdfs(new List<string> { headerPdf, zf }, mergedPdf);
                                    zipPdfFiles.Add(mergedPdf);
                                    File.Delete(headerPdf);
                                }
                                else if (zfExt == ".doc" || zfExt == ".docx" || zfExt == ".xls" || zfExt == ".xlsx")
                                {
                                    if (OfficeConversionService.TryConvertOfficeToPdf(zf, zfPdf))
                                    {
                                        string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                                        PdfService.AddHeaderPdf(headerPdf, $"{headerText} (ZIP: {safeEntryName})");
                                        string mergedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                                        PdfAppendTest.AppendPdfs(new List<string> { headerPdf, zfPdf }, mergedPdf);
                                        zipPdfFiles.Add(mergedPdf);
                                        File.Delete(headerPdf);
                                        File.Delete(zfPdf);
                                    }
                                    else
                                    {
                                        PdfService.AddPlaceholderPdf(zfPdf, $"Could not convert attachment: {safeEntryName}");
                                        Console.WriteLine($"[ATTACH] ZIP Office failed to convert: {zf}");
                                    }
                                }
                                else
                                {
                                    PdfService.AddPlaceholderPdf(zfPdf, $"Unsupported attachment: {safeEntryName}");
                                    zipPdfFiles.Add(zfPdf);
                                    Console.WriteLine($"[ATTACH] ZIP unsupported: {zf}");
                                }

                                File.Delete(zf);
                            }

                            if (zipPdfFiles.Count > 0)
                            {
                                if (zipPdfFiles.Count == 1)
                                {
                                    tempPdfFiles.Add(zipPdfFiles[0]);
                                }
                                else
                                {
                                    string zipMergedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_merged.pdf");
                                    PdfAppendTest.AppendPdfs(zipPdfFiles, zipMergedPdf);
                                    tempPdfFiles.Add(zipMergedPdf);
                                    foreach (var zf in zipPdfFiles)
                                        File.Delete(zf);
                                }
                            }
                            else
                            {
                                PdfService.AddPlaceholderPdf(attPdf, $"Unsupported attachment: {attName}");
                                tempPdfFiles.Add(attPdf);
                                Console.WriteLine($"[ATTACH] ZIP unsupported: {attName}");
                            }
                        }
                    }
                    else
                    {
                        PdfService.AddPlaceholderPdf(attPdf, $"Unsupported attachment: {attName}");
                        tempPdfFiles.Add(attPdf);
                        Console.WriteLine($"[ATTACH] Unsupported type: {attName}");
                    }
                    Console.WriteLine($"[ATTACH] Finished: {attName}");

                    File.Delete(attPath);
                }
                catch (Exception ex)
                {
                    string attPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                    PdfService.AddPlaceholderPdf(attPdf, $"Error processing attachment: {ex.Message}");
                    tempPdfFiles.Add(attPdf);
                    Console.WriteLine($"[ATTACH] Exception: {att.FileName ?? "unnamed"} - {ex}");
                }
            }

            // Append all attachment PDFs to the main PDF
            if (tempPdfFiles.Count > 0)
            {
                tempPdfFiles.Insert(0, mainPdfPath);
                string mergedMainPdf = Path.Combine(tempDir, Guid.NewGuid() + "_final_merged.pdf");
                PdfAppendTest.AppendPdfs(tempPdfFiles, mergedMainPdf);

                // Only delete/replace the original main PDF if it is a PDF and deleteSourcePdf is true
                string ext = Path.GetExtension(mainPdfPath)?.ToLowerInvariant();
                if (ext == ".pdf" && deleteSourcePdf)
                {
                    File.Delete(mainPdfPath);
                    File.Move(mergedMainPdf, mainPdfPath);
                }
                else
                {
                    // If not deleting, just copy the merged PDF to the output location, do not delete the source
                    File.Copy(mergedMainPdf, mainPdfPath, true);
                }

                // Clean up temp files
                foreach (var tempFile in tempPdfFiles.Skip(1)) // Skip the first one which was the original main PDF
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[CLEANUP] Failed to delete temp file {tempFile}: {ex.Message}");
                    }
                }
            }
        }
    }
}
