using System;
using System.Collections.Generic;
using System.Windows;
using MsgToPdfConverter.Utils;
using System.IO;
using MsgReader.Outlook;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Diagnostics;
using DinkToPdf;
using DinkToPdf.Contracts;
using System.Text.RegularExpressions;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.IO.Image;
using iText.Layout.Element;

namespace MsgToPdfConverter
{
    public partial class MainWindow : Window
    {
        private List<string> selectedFiles = new List<string>();

        public MainWindow()
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            InitializeComponent();
            CheckDotNetRuntime();
        }

        private void CheckDotNetRuntime()
        {
            // Only check if not running in design mode
            if (!System.ComponentModel.DesignerProperties.GetIsInDesignMode(this))
            {
                if (!IsDotNetDesktopRuntimeInstalled())
                {
                    var result = MessageBox.Show(
                        ".NET Desktop Runtime 5.0 is required to run this application. Would you like to download it now?",
                        ".NET Runtime Required",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        string url = "https://dotnet.microsoft.com/en-us/download/dotnet/5.0/runtime";
                        try
                        {
                            Process.Start("explorer", url);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Could not open browser. Please visit this URL manually:\n{url}\nError: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        Application.Current.Shutdown();
                    }
                    else
                    {
                        Application.Current.Shutdown();
                    }
                }
            }
        }

        private bool IsDotNetDesktopRuntimeInstalled()
        {
            // Simple check: look for a known .NET 5+ runtime folder
            string windir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string dotnetDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet", "shared", "Microsoft.WindowsDesktop.App");
            if (Directory.Exists(dotnetDir))
            {
                var versions = Directory.GetDirectories(dotnetDir);
                foreach (var v in versions)
                {
                    if (v.Contains("5.0")) return true;
                }
            }
            return false;
        }

        private void SelectFilesButton_Click(object sender, RoutedEventArgs e)
        {
            selectedFiles = FileDialogHelper.OpenMsgFileDialog();
            FilesListBox.Items.Clear();
            if (selectedFiles != null && selectedFiles.Count > 0)
            {
                foreach (var file in selectedFiles)
                {
                    FilesListBox.Items.Add(file);
                }
                ConvertButton.IsEnabled = true;
            }
            else
            {
                ConvertButton.IsEnabled = false;
            }
        }

        private void SelectFolderButton_Click(object sender, RoutedEventArgs e)
        {
            selectedFiles = FileDialogHelper.OpenMsgFolderDialog();
            FilesListBox.Items.Clear();
            if (selectedFiles != null && selectedFiles.Count > 0)
            {
                foreach (var file in selectedFiles)
                {
                    FilesListBox.Items.Add(file);
                }
                ConvertButton.IsEnabled = true;
            }
            else
            {
                ConvertButton.IsEnabled = false;
            }
        }

        private string GetMimeTypeFromFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "image/png";
            string ext = System.IO.Path.GetExtension(fileName).ToLowerInvariant();
            switch (ext)
            {
                case ".jpg":
                case ".jpeg": return "image/jpeg";
                case ".png": return "image/png";
                case ".gif": return "image/gif";
                case ".bmp": return "image/bmp";
                case ".tif":
                case ".tiff": return "image/tiff";
                case ".svg": return "image/svg+xml";
                default: return "image/png";
            }
        }

        private string EmbedInlineImages(Storage.Message msg)
        {
            string html = msg.BodyHtml ?? msg.BodyText;
            if (string.IsNullOrEmpty(html) || msg.Attachments == null || msg.Attachments.Count == 0)
                return html;

            var regex = new Regex("<img[^>]+src=\"cid:([^\"]+)\"", RegexOptions.IgnoreCase);
            return regex.Replace(html, match =>
            {
                string cid = match.Groups[1].Value;
                Storage.Attachment found = null;
                foreach (var att in msg.Attachments)
                {
                    if (att is Storage.Attachment attachment && attachment.ContentId != null && attachment.ContentId.Trim('<', '>') == cid.Trim('<', '>'))
                    {
                        found = attachment;
                        break;
                    }
                }
                if (found != null)
                {
                    string mimeType = GetMimeTypeFromFileName(found.FileName);
                    string base64 = Convert.ToBase64String(found.Data);
                    return match.Value.Replace($"cid:{cid}", $"data:{mimeType};base64,{base64}");
                }
                return match.Value;
            });
        }

        private string BuildEmailHtml(Storage.Message msg)
        {
            string from = msg.Sender?.DisplayName ?? msg.Sender?.Email ?? "";
            string sent = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("f") : "";
            string to = string.Join(", ", msg.Recipients?.FindAll(r => r.Type == Storage.Recipient.RecipientType.To)?.ConvertAll(r => r.DisplayName + (string.IsNullOrEmpty(r.Email) ? "" : $" <{r.Email}>")) ?? new List<string>());
            string cc = string.Join(", ", msg.Recipients?.FindAll(r => r.Type == Storage.Recipient.RecipientType.Cc)?.ConvertAll(r => r.DisplayName + (string.IsNullOrEmpty(r.Email) ? "" : $" <{r.Email}>")) ?? new List<string>());
            string subject = msg.Subject ?? "";
            string body = EmbedInlineImages(msg) ?? "";

            string header = $@"
                <div style='font-family:Segoe UI,Arial,sans-serif;font-size:12pt;margin-bottom:16px;'>
                    <div><b>From:</b> {System.Net.WebUtility.HtmlEncode(from)}</div>
                    <div><b>Sent:</b> {System.Net.WebUtility.HtmlEncode(sent)}</div>
                    <div><b>To:</b> {System.Net.WebUtility.HtmlEncode(to)}</div>
                    {(string.IsNullOrWhiteSpace(cc) ? "" : $"<div><b>Cc:</b> {System.Net.WebUtility.HtmlEncode(cc)}</div>")}
                    <div><b>Subject:</b> {System.Net.WebUtility.HtmlEncode(subject)}</div>
                </div>";
            return header + body;
        }

        private void LogDebug(string message)
        {
            try
            {
                string logPath = Path.Combine(Path.GetTempPath(), "MsgToPdfConverter_debug.log");
                File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\r\n");
            }
            catch { }
        }

        private async void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("Convert clicked");
            LogDebug("ConvertButton_Click started");
            if (selectedFiles == null || selectedFiles.Count == 0)
            {
                LogDebug("No files selected");
                Console.WriteLine("No files selected.");
                return;
            }
            Console.WriteLine("Files selected, starting conversion");
            int success = 0, fail = 0;
            ProgressBar.Visibility = Visibility.Visible;
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = selectedFiles.Count;
            ProgressBar.Value = 0;
            var converter = new SynchronizedConverter(new PdfTools());
            bool appendAttachments = false;
            Dispatcher.Invoke(() => appendAttachments = AppendAttachmentsCheckBox.IsChecked == true);

            await System.Threading.Tasks.Task.Run(() =>
            {
                Console.WriteLine("[TASK] Started conversion task");
                try
                {
                    for (int i = 0; i < selectedFiles.Count; i++)
                    {
                        Console.WriteLine($"[TASK] Processing file {i + 1} of {selectedFiles.Count}: {selectedFiles[i]}");
                        var msgFilePath = selectedFiles[i];
                        LogDebug($"Processing file: {msgFilePath}");
                        try
                        {
                            var msg = new Storage.Message(msgFilePath);
                            Console.WriteLine($"[TASK] Loaded MSG: {msgFilePath}");
                            LogDebug($"Loaded MSG: {msgFilePath}");
                            string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
                            string baseName = Path.GetFileNameWithoutExtension(msgFilePath);
                            string dir = Path.GetDirectoryName(msgFilePath);
                            string pdfFilePath = Path.Combine(dir, $"{baseName} - {datePart}.pdf");
                            string htmlWithHeader = BuildEmailHtml(msg);
                            Console.WriteLine($"[TASK] Built HTML for: {msgFilePath}");
                            LogDebug($"Building PDF: {pdfFilePath}");
                            var doc = new HtmlToPdfDocument()
                            {
                                GlobalSettings = new GlobalSettings
                                {
                                    ColorMode = ColorMode.Color,
                                    Orientation = Orientation.Portrait,
                                    PaperSize = PaperKind.A4,
                                    Out = pdfFilePath
                                },
                                Objects = {
                                    new ObjectSettings
                                    {
                                        PagesCount = true,
                                        HtmlContent = htmlWithHeader,
                                        WebSettings = { DefaultEncoding = "utf-8" }
                                    }
                                }
                            };
                            Console.WriteLine("[TASK] About to call converter.Convert(doc);");
                            converter.Convert(doc);
                            Console.WriteLine("[TASK] Finished converter.Convert(doc);");
                            Console.WriteLine($"[TASK] Email PDF created: {pdfFilePath}");
                            LogDebug($"Email PDF created: {pdfFilePath}");

                            if (appendAttachments && msg.Attachments != null && msg.Attachments.Count > 0)
                            {
                                // Step 1: Extract and convert attachments to PDF, collect all PDF paths
                                var typedAttachments = new List<Storage.Attachment>();
                                foreach (var att in msg.Attachments)
                                {
                                    if (att is Storage.Attachment a) typedAttachments.Add(a);
                                }
                                LogDebug($"Extracting and converting {typedAttachments.Count} attachments for {pdfFilePath}");
                                Console.WriteLine($"Extracting and converting {typedAttachments.Count} attachments for {pdfFilePath}");
                                var allPdfFiles = new List<string> { pdfFilePath };
                                var allTempFiles = new List<string>(); // Track all temp files (originals and PDFs)
                                string tempDir = Path.GetDirectoryName(pdfFilePath);
                                foreach (var att in typedAttachments)
                                {
                                    string attName = att.FileName ?? "attachment";
                                    string ext = Path.GetExtension(attName).ToLowerInvariant();
                                    string attPath = Path.Combine(tempDir, attName);
                                    string attPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName) + ".pdf");
                                    try
                                    {
                                        File.WriteAllBytes(attPath, att.Data);
                                        allTempFiles.Add(attPath); // Track the extracted file
                                        if (ext == ".pdf")
                                        {
                                            allPdfFiles.Add(attPath);
                                        }
                                        else if (ext == ".jpg" || ext == ".jpeg")
                                        {
                                            using (var writer = new iText.Kernel.Pdf.PdfWriter(attPdf))
                                            using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                                            using (var docImg = new iText.Layout.Document(pdf))
                                            {
                                                var imgData = iText.IO.Image.ImageDataFactory.Create(attPath);
                                                var image = new iText.Layout.Element.Image(imgData);
                                                docImg.Add(image);
                                            }
                                            allPdfFiles.Add(attPdf);
                                            allTempFiles.Add(attPdf); // Track the converted PDF
                                        }
                                        else if (ext == ".doc" || ext == ".docx" || ext == ".xls" || ext == ".xlsx")
                                        {
                                            if (TryConvertOfficeToPdf(attPath, attPdf))
                                            {
                                                allPdfFiles.Add(attPdf);
                                                allTempFiles.Add(attPdf);
                                            }
                                        }
                                        else if (ext == ".zip")
                                        {
                                            string extractDir = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName));
                                            System.IO.Compression.ZipFile.ExtractToDirectory(attPath, extractDir);
                                            allTempFiles.Add(attPath); // Track the zip file
                                            var zipFiles = Directory.GetFiles(extractDir, "*.*", SearchOption.AllDirectories);
                                            foreach (var zf in zipFiles)
                                            {
                                                allTempFiles.Add(zf); // Track all extracted files
                                                string zfPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(zf) + ".pdf");
                                                string zfExt = Path.GetExtension(zf).ToLowerInvariant();
                                                if (zfExt == ".pdf")
                                                {
                                                    allPdfFiles.Add(zf);
                                                }
                                                else if (zfExt == ".doc" || zfExt == ".docx" || zfExt == ".xls" || zfExt == ".xlsx")
                                                {
                                                    if (TryConvertOfficeToPdf(zf, zfPdf))
                                                    {
                                                        allPdfFiles.Add(zfPdf);
                                                        allTempFiles.Add(zfPdf);
                                                    }
                                                }
                                            }
                                            allTempFiles.Add(extractDir); // Track the extracted directory for cleanup
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogDebug($"Exception processing attachment: {ex}");
                                        Console.WriteLine($"[ATTACH] Exception: {attName} - {ex}");
                                    }
                                }
                                // Step 2: Merge all PDFs into a temp file
                                string mergedPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(pdfFilePath) + "_merged.pdf");
                                PdfAppendTest.AppendPdfs(allPdfFiles, mergedPdf);
                                // Step 3: Delete original files (main and attachments, and all temp files except merged)
                                foreach (var f in allTempFiles)
                                {
                                    if (File.Exists(f) && !string.Equals(f, mergedPdf, StringComparison.OrdinalIgnoreCase) && !string.Equals(f, pdfFilePath, StringComparison.OrdinalIgnoreCase))
                                    {
                                        try { File.Delete(f); } catch { }
                                    }
                                    else if (Directory.Exists(f))
                                    {
                                        try { Directory.Delete(f, true); } catch { }
                                    }
                                }
                                // Step 4: Rename merged file back to original name
                                if (File.Exists(mergedPdf))
                                {
                                    File.Move(mergedPdf, pdfFilePath, true);
                                }
                                LogDebug($"Merged and replaced {pdfFilePath}");
                                Console.WriteLine($"Merged and replaced {pdfFilePath}");
                            }
                            success++;
                            LogDebug($"Success: {msgFilePath}");
                        }
                        catch (Exception ex)
                        {
                            fail++;
                            LogDebug($"Failed: {msgFilePath} - {ex}");
                            Console.WriteLine($"Failed to convert: {msgFilePath}\nError: {ex.Message}");
                        }
                        Dispatcher.Invoke(() => ProgressBar.Value = i + 1);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[TASK] UNHANDLED EXCEPTION: {ex}");
                }
            });
            ProgressBar.Visibility = Visibility.Collapsed;
            LogDebug($"Conversion completed. Success: {success}, Failed: {fail}");
            Console.WriteLine($"Conversion completed. Success: {success}, Failed: {fail}");
        }

        // Appends attachments as PDFs to the main email PDF
        private void AppendAttachmentsToPdf(string mainPdfPath, List<Storage.Attachment> attachments, SynchronizedConverter converter)
        {
            LogDebug($"Total attachments: {attachments.Count}");
            string allNames = string.Join(", ", attachments.ConvertAll(a => a.FileName ?? "(no name)"));
            LogDebug($"Attachment names: {allNames}");
            Console.WriteLine($"Total attachments: {attachments.Count}\nNames: {allNames}");
            var tempPdfFiles = new List<string> { mainPdfPath };
            // Use the original file directory for temp files
            string tempDir = Path.GetDirectoryName(mainPdfPath);
            // Directory.CreateDirectory(tempDir); // Not needed, should already exist

            foreach (var att in attachments)
            {
                string attName = att.FileName ?? "attachment";
                string ext = Path.GetExtension(attName).ToLowerInvariant();
                string attPath = Path.Combine(tempDir, attName);
                string attPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName) + ".pdf");
                try
                {
                    File.WriteAllBytes(attPath, att.Data);
                    if (ext == ".pdf")
                    {
                        tempPdfFiles.Add(attPath);
                    }
                    else if (ext == ".jpg" || ext == ".jpeg")
                    {
                        // Convert JPG to PDF using iText7
                        using (var writer = new iText.Kernel.Pdf.PdfWriter(attPdf))
                        using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                        using (var doc = new iText.Layout.Document(pdf))
                        {
                            var imgData = iText.IO.Image.ImageDataFactory.Create(attPath);
                            var image = new iText.Layout.Element.Image(imgData);
                            doc.Add(image);
                        }
                        tempPdfFiles.Add(attPdf);
                    }
                    else if (ext == ".doc" || ext == ".docx" || ext == ".xls" || ext == ".xlsx")
                    {
                        if (TryConvertOfficeToPdf(attPath, attPdf))
                        {
                            tempPdfFiles.Add(attPdf);
                        }
                    }
                    else if (ext == ".zip")
                    {
                        Console.WriteLine($"[ATTACH] ZIP detected, extracting: {attName}");
                        string extractDir = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName));
                        System.IO.Compression.ZipFile.ExtractToDirectory(attPath, extractDir);
                        var zipFiles = Directory.GetFiles(extractDir, "*.*", SearchOption.AllDirectories);
                        foreach (var zf in zipFiles)
                        {
                            string zfPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(zf) + ".pdf");
                            string zfExt = Path.GetExtension(zf).ToLowerInvariant();
                            if (zfExt == ".pdf")
                            {
                                string pdfCopy = Path.Combine(tempDir, Guid.NewGuid() + "_zipattachment.pdf");
                                File.Copy(zf, pdfCopy, true);
                                tempPdfFiles.Add(pdfCopy);
                                Console.WriteLine($"[ATTACH] ZIP PDF added: {pdfCopy}");
                            }
                            else if (zfExt == ".doc" || zfExt == ".docx" || zfExt == ".xls" || zfExt == ".xlsx")
                            {
                                if (TryConvertOfficeToPdf(zf, zfPdf))
                                {
                                    tempPdfFiles.Add(zfPdf);
                                    Console.WriteLine($"[ATTACH] ZIP Office converted: {zfPdf}");
                                }
                                else
                                {
                                    AddPlaceholderPdf(zfPdf, $"Could not convert attachment: {Path.GetFileName(zf)}");
                                    Console.WriteLine($"[ATTACH] ZIP Office failed to convert: {zf}");
                                }
                            }
                            else
                            {
                                AddPlaceholderPdf(zfPdf, $"Unsupported attachment: {Path.GetFileName(zf)}");
                                tempPdfFiles.Add(zfPdf);
                                Console.WriteLine($"[ATTACH] ZIP unsupported: {zf}");
                            }
                        }
                    }
                    else
                    {
                        AddPlaceholderPdf(attPdf, $"Unsupported attachment: {attName}");
                        tempPdfFiles.Add(attPdf);
                        Console.WriteLine($"[ATTACH] Unsupported type: {attName}");
                    }
                    Console.WriteLine($"[ATTACH] Finished: {attName}");
                }
                catch (Exception ex)
                {
                    AddPlaceholderPdf(attPdf, $"Error processing attachment: {ex.Message}");
                    tempPdfFiles.Add(attPdf);
                    LogDebug($"Exception processing attachment: {ex}");
                    Console.WriteLine($"[ATTACH] Exception: {attName} - {ex}");
                }
            }

            // Merge all tempPdfFiles using the robust iText7 method from PdfAppendTest
            try
            {
                PdfAppendTest.AppendPdfs(tempPdfFiles, mainPdfPath);
            }
            finally
            {
                // Do not delete temp files for now
            }
        }

        // Merges multiple PDFs into one using iText7, never including the output file as an input
        private void MergePdfs(string[] pdfFiles, string outputPdf)
        {
            Console.WriteLine($"[MERGE] (iText7) Merging PDFs into: {outputPdf}");
            // Filter out the output file if present in the input list
            var inputFiles = new List<string>();
            foreach (var f in pdfFiles)
            {
                if (!string.Equals(f, outputPdf, StringComparison.OrdinalIgnoreCase))
                    inputFiles.Add(f);
            }
            // Filter out PDFs that are empty or invalid
            var validInputFiles = new List<string>();
            foreach (var pdf in inputFiles)
            {
                try
                {
                    using (var reader = new iText.Kernel.Pdf.PdfReader(pdf))
                    using (var doc = new iText.Kernel.Pdf.PdfDocument(reader))
                    {
                        int n = doc.GetNumberOfPages();
                        if (n > 0)
                        {
                            validInputFiles.Add(pdf);
                        }
                        else
                        {
                            Console.WriteLine($"[MERGE] Skipping empty PDF: {pdf}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MERGE] Skipping invalid PDF: {pdf} - {ex.Message}");
                }
            }
            if (validInputFiles.Count == 0)
            {
                Console.WriteLine("[MERGE] No valid PDFs to merge. Aborting.");
                return;
            }
            using (var stream = new FileStream(outputPdf, FileMode.Create, FileAccess.Write))
            using (var pdfWriter = new iText.Kernel.Pdf.PdfWriter(stream))
            using (var pdfDoc = new iText.Kernel.Pdf.PdfDocument(pdfWriter))
            {
                foreach (var pdf in validInputFiles)
                {
                    using (var srcPdf = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(pdf)))
                    {
                        int n = srcPdf.GetNumberOfPages();
                        srcPdf.CopyPagesTo(1, n, pdfDoc);
                    }
                }
            }
            Console.WriteLine($"[MERGE] (iText7) Saved merged PDF: {outputPdf}");
        }

        // Adds a single-page PDF with a message
        private void AddPlaceholderPdf(string pdfPath, string message, string imagePath = null)
        {
            using (var doc = new PdfSharp.Pdf.PdfDocument())
            {
                var page = doc.AddPage();
                using (var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                {
                    if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                    {
                        try
                        {
                            Console.WriteLine($"[IMG2PDF] Attempting to load image: {imagePath}");
                            var img = PdfSharp.Drawing.XImage.FromFile(imagePath);
                            Console.WriteLine($"[IMG2PDF] Loaded image: {imagePath}");
                            double maxWidth = page.Width.Point - 80;
                            double maxHeight = page.Height.Point - 300;
                            double scale = Math.Min(maxWidth / img.PixelWidth * 72.0 / img.HorizontalResolution, maxHeight / img.PixelHeight * 72.0 / img.VerticalResolution);
                            double imgWidth = img.PixelWidth * 72.0 / img.HorizontalResolution * scale;
                            double imgHeight = img.PixelHeight * 72.0 / img.VerticalResolution * scale;
                            double x = (page.Width.Point - imgWidth) / 2;
                            double y = 100;
                            gfx.DrawImage(img, x, y, imgWidth, imgHeight);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[IMG2PDF] Failed to load image: {imagePath} - {ex.Message}");
                        }
                    }
                    var font = new PdfSharp.Drawing.XFont("Arial", 16);
                    gfx.DrawString(message, font, PdfSharp.Drawing.XBrushes.Black, new PdfSharp.Drawing.XRect(40, page.Height.Point - 200, page.Width.Point - 80, 100), PdfSharp.Drawing.XStringFormats.Center);
                }
                doc.Save(pdfPath);
            }
        }

        // Attempts to convert Office files to PDF using Office Interop (requires Office installed)
        private bool TryConvertOfficeToPdf(string inputPath, string outputPdf)
        {
            string ext = Path.GetExtension(inputPath).ToLowerInvariant();
            bool result = false;
            Exception threadEx = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    if (ext == ".doc" || ext == ".docx")
                    {
                        var wordApp = new Microsoft.Office.Interop.Word.Application();
                        var doc = wordApp.Documents.Open(inputPath);
                        doc.ExportAsFixedFormat(outputPdf, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                        doc.Close();
                        wordApp.Quit();
                        result = true;
                    }
                    else if (ext == ".xls" || ext == ".xlsx")
                    {
                        var excelApp = new Microsoft.Office.Interop.Excel.Application();
                        var wb = excelApp.Workbooks.Open(inputPath);
                        wb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPdf);
                        wb.Close();
                        excelApp.Quit();
                        result = true;
                    }
                }
                catch (Exception ex)
                {
                    threadEx = ex;
                }
            });
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            if (threadEx != null)
            {
                // Optionally log threadEx
                return false;
            }
            return result;
        }

        private void FilesListBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Delete && FilesListBox.SelectedItems.Count > 0)
            {
                var itemsToRemove = new List<string>();
                foreach (var item in FilesListBox.SelectedItems)
                {
                    itemsToRemove.Add(item as string);
                }
                foreach (var item in itemsToRemove)
                {
                    FilesListBox.Items.Remove(item);
                    selectedFiles.Remove(item);
                }
                ConvertButton.IsEnabled = FilesListBox.Items.Count > 0;
            }
        }
    }
}