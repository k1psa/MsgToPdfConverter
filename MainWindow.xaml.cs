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
using System.Threading.Tasks;
using System.Linq;
using System.Runtime.InteropServices;

namespace MsgToPdfConverter
{
    public partial class MainWindow : Window
    {
        private List<string> selectedFiles = new List<string>();
        private int convertButtonClickCount = 0;
        private bool isConverting = false;

        public MainWindow()
        {
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

        private void KillWkhtmltopdfProcesses()
        {
            try
            {
                var procs = System.Diagnostics.Process.GetProcessesByName("wkhtmltopdf");
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

        private void RunDinkToPdfConversion(HtmlToPdfDocument doc)
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
                    var converter = new SynchronizedConverter(new PdfTools());
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

        private async void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("[DEBUG] Entered ConvertButton_Click");
            if (isConverting)
            {
                Console.WriteLine("[DEBUG] Conversion already in progress, ignoring click.");
                return;
            }
            isConverting = true;
            convertButtonClickCount++;
            Console.WriteLine($"[DEBUG] Convert button pressed {convertButtonClickCount} time(s)");
            try
            {
                Console.WriteLine("[DEBUG] Disabling ConvertButton and showing ProgressBar");
                ConvertButton.IsEnabled = false;
                ProgressBar.Visibility = Visibility.Visible;
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = selectedFiles.Count;
                ProgressBar.Value = 0;
                int success = 0, fail = 0;
                bool appendAttachments = AppendAttachmentsCheckBox.IsChecked == true;
                Console.WriteLine($"[DEBUG] appendAttachments: {appendAttachments}");
                await Task.Run(() =>
                {
                    Console.WriteLine("[DEBUG] Starting batch loop");
                    for (int i = 0; i < selectedFiles.Count; i++)
                    {
                        Console.WriteLine($"[DEBUG] Loop index: {i}");
                        try
                        {
                            Console.WriteLine($"[TASK] Processing file {i + 1} of {selectedFiles.Count}: {selectedFiles[i]}");
                            var msgFilePath = selectedFiles[i];
                            Console.WriteLine($"[TASK] Loading MSG: {msgFilePath}");
                            var msg = new Storage.Message(msgFilePath);
                            Console.WriteLine($"[TASK] Loaded MSG: {msgFilePath}");
                            string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
                            string baseName = Path.GetFileNameWithoutExtension(msgFilePath);
                            string dir = Path.GetDirectoryName(msgFilePath);
                            string pdfFilePath = Path.Combine(dir, $"{baseName} - {datePart}.pdf");
                            if (File.Exists(pdfFilePath))
                            {
                                try { File.Delete(pdfFilePath); Console.WriteLine($"[DEBUG] Deleted old PDF: {pdfFilePath}"); } catch (Exception ex) { Console.WriteLine($"[DEBUG] Could not delete old PDF: {ex.Message}"); }
                            }
                            Console.WriteLine($"[TASK] Output PDF path: {pdfFilePath}");
                            string htmlWithHeader = BuildEmailHtml(msg);
                            Console.WriteLine($"[TASK] Built HTML for: {msgFilePath}");
                            Console.WriteLine($"[TASK] HTML length: {htmlWithHeader?.Length ?? 0}");
                            // Write HTML to a temp file
                            string tempHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".html");
                            File.WriteAllText(tempHtmlPath, htmlWithHeader);
                            Console.WriteLine($"[DEBUG] Written HTML to temp file: {tempHtmlPath}");
                            // Find all inline ContentIds
                            var inlineContentIds = GetInlineContentIds(htmlWithHeader);
                            // Launch a new process for each conversion
                            var psi = new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName,
                                Arguments = $"--html2pdf \"{tempHtmlPath}\" \"{pdfFilePath}\"",
                                UseShellExecute = false,
                                CreateNoWindow = true,
                                RedirectStandardOutput = true,
                                RedirectStandardError = true
                            };
                            Console.WriteLine($"[DEBUG] Starting HtmlToPdfWorker process: {psi.FileName} {psi.Arguments}");
                            var proc = System.Diagnostics.Process.Start(psi);
                            string stdOut = proc.StandardOutput.ReadToEnd();
                            string stdErr = proc.StandardError.ReadToEnd();
                            proc.WaitForExit();
                            Console.WriteLine($"[DEBUG] HtmlToPdfWorker exit code: {proc.ExitCode}");
                            if (proc.ExitCode != 0)
                            {
                                Console.WriteLine($"[DEBUG] HtmlToPdfWorker failed. StdOut: {stdOut} StdErr: {stdErr}");
                                throw new Exception($"HtmlToPdfWorker failed: {stdErr}");
                            }
                            File.Delete(tempHtmlPath);
                            Console.WriteLine($"[TASK] Email PDF created: {pdfFilePath}");
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            if (appendAttachments && msg.Attachments != null && msg.Attachments.Count > 0)
                            {
                                Console.WriteLine($"[DEBUG] Processing attachments for {pdfFilePath}");
                                var typedAttachments = new List<Storage.Attachment>();
                                foreach (var att in msg.Attachments)
                                {
                                    // Exclude inline images (signature, etc.)
                                    if (att is Storage.Attachment a)
                                    {
                                        if ((a.IsInline == true) || (!string.IsNullOrEmpty(a.ContentId) && inlineContentIds.Contains(a.ContentId.Trim('<', '>', '\"', '\'', ' '))))
                                        {
                                            Console.WriteLine($"[DEBUG] Skipping inline attachment: {a.FileName} (ContentId: {a.ContentId}, IsInline: {a.IsInline})");
                                            continue;
                                        }
                                        typedAttachments.Add(a);
                                    }
                                }
                                var allPdfFiles = new List<string> { pdfFilePath };
                                var allTempFiles = new List<string>();
                                string tempDir = Path.GetDirectoryName(pdfFilePath);
                                int totalAttachments = typedAttachments.Count;
                                for (int attIndex = 0; attIndex < typedAttachments.Count; attIndex++)
                                {
                                    var att = typedAttachments[attIndex];
                                    string attName = att.FileName ?? "attachment";
                                    string ext = Path.GetExtension(attName).ToLowerInvariant();
                                    string attPath = Path.Combine(tempDir, attName);
                                    string attPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName) + ".pdf");
                                    string headerText = $"Attachment: {attIndex + 1}/{totalAttachments}";
                                    try
                                    {
                                        File.WriteAllBytes(attPath, att.Data);
                                        allTempFiles.Add(attPath);
                                        string finalAttachmentPdf = null;
                                        if (ext == ".pdf")
                                        {
                                            string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                                            AddHeaderPdf(headerPdf, headerText);
                                            finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                                            PdfAppendTest.AppendPdfs(new List<string> { headerPdf, attPath }, finalAttachmentPdf);
                                            allTempFiles.Add(headerPdf);
                                            allTempFiles.Add(finalAttachmentPdf);
                                        }
                                        else if (ext == ".jpg" || ext == ".jpeg")
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
                                            finalAttachmentPdf = attPdf;
                                            allTempFiles.Add(attPdf);
                                        }
                                        else if (ext == ".doc" || ext == ".docx" || ext == ".xls" || ext == ".xlsx")
                                        {
                                            if (TryConvertOfficeToPdf(attPath, attPdf))
                                            {
                                                string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                                                AddHeaderPdf(headerPdf, headerText);
                                                finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                                                PdfAppendTest.AppendPdfs(new List<string> { headerPdf, attPdf }, finalAttachmentPdf);
                                                allTempFiles.Add(headerPdf);
                                                allTempFiles.Add(attPdf); // Add the Office-generated PDF to cleanup list
                                                allTempFiles.Add(finalAttachmentPdf);
                                            }
                                            else
                                            {
                                                // Conversion failed, add placeholder
                                                finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                                                AddHeaderPdf(finalAttachmentPdf, headerText + "\n(Conversion failed)");
                                                allTempFiles.Add(finalAttachmentPdf);
                                            }
                                        }
                                        else if (ext == ".zip")
                                        {
                                            string extractDir = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName));
                                            System.IO.Compression.ZipFile.ExtractToDirectory(attPath, extractDir);
                                            allTempFiles.Add(attPath);
                                            var zipFiles = Directory.GetFiles(extractDir, "*.*", SearchOption.AllDirectories);
                                            int zipFileIndex = 0;
                                            foreach (var zf in zipFiles)
                                            {
                                                zipFileIndex++;
                                                string zfPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(zf) + ".pdf");
                                                string zfExt = Path.GetExtension(zf).ToLowerInvariant();
                                                string zipHeader = $"Attachment: {attIndex + 1}/{totalAttachments} (ZIP {zipFileIndex}/{zipFiles.Length})";
                                                string finalZipPdf = null;
                                                if (zfExt == ".pdf")
                                                {
                                                    string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                                                    AddHeaderPdf(headerPdf, zipHeader);
                                                    string mergedZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                                                    PdfAppendTest.AppendPdfs(new List<string> { headerPdf, zf }, mergedZipPdf);
                                                    allPdfFiles.Add(mergedZipPdf);
                                                    allTempFiles.Add(headerPdf);
                                                    allTempFiles.Add(mergedZipPdf);
                                                }
                                                else if (zfExt == ".doc" || zfExt == ".docx" || zfExt == ".xls" || zfExt == ".xlsx")
                                                {
                                                    if (TryConvertOfficeToPdf(zf, zfPdf))
                                                    {
                                                        string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                                                        AddHeaderPdf(headerPdf, zipHeader);
                                                        finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                                                        PdfAppendTest.AppendPdfs(new List<string> { headerPdf, zfPdf }, finalZipPdf);
                                                        allTempFiles.Add(headerPdf);
                                                        allTempFiles.Add(zfPdf); // Add the Office-generated PDF to cleanup list
                                                        allTempFiles.Add(finalZipPdf);
                                                    }
                                                    else
                                                    {
                                                        finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                                                        AddHeaderPdf(finalZipPdf, zipHeader + "\n(Conversion failed)");
                                                        allTempFiles.Add(finalZipPdf);
                                                    }
                                                }
                                                else
                                                {
                                                    finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                                                    AddHeaderPdf(finalZipPdf, zipHeader + "\n(Unsupported type)");
                                                    allTempFiles.Add(finalZipPdf);
                                                }
                                                if (finalZipPdf != null)
                                                    allPdfFiles.Add(finalZipPdf);
                                            }
                                            allTempFiles.Add(extractDir);
                                        }
                                        else
                                        {
                                            // Unsupported type, add placeholder
                                            finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                                            AddHeaderPdf(finalAttachmentPdf, headerText + "\n(Unsupported type)");
                                            allTempFiles.Add(finalAttachmentPdf);
                                        }
                                        if (finalAttachmentPdf != null)
                                            allPdfFiles.Add(finalAttachmentPdf);
                                    }
                                    catch (Exception ex)
                                    {
                                        // On error, add placeholder
                                        string errorPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                                        AddHeaderPdf(errorPdf, headerText + $"\n(Error: {ex.Message})");
                                        allPdfFiles.Add(errorPdf);
                                        allTempFiles.Add(errorPdf);
                                        Console.WriteLine($"[ATTACH] Exception: {attName} - {ex}");
                                    }
                                }
                                string mergedPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(pdfFilePath) + "_merged.pdf");
                                Console.WriteLine($"[DEBUG] Before PDF merge: {string.Join(", ", allPdfFiles)} -> {mergedPdf}");
                                PdfAppendTest.AppendPdfs(allPdfFiles, mergedPdf);
                                Console.WriteLine("[DEBUG] After PDF merge");
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                System.Threading.Thread.Sleep(200);
                                foreach (var f in allTempFiles)
                                {
                                    try
                                    {
                                        if (File.Exists(f) && !string.Equals(f, mergedPdf, StringComparison.OrdinalIgnoreCase) && !string.Equals(f, pdfFilePath, StringComparison.OrdinalIgnoreCase))
                                        {
                                            RobustDeleteFile(f);
                                        }
                                        else if (Directory.Exists(f))
                                        {
                                            try
                                            {
                                                Directory.Delete(f, true);
                                                Console.WriteLine($"[CLEANUP] Deleted temp directory: {f}");
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"[CLEANUP] Failed to delete temp directory: {f} - {ex.Message}");
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"[CLEANUP] Unexpected error deleting temp file or directory: {f} - {ex.Message}");
                                    }
                                }
                                Console.WriteLine("[DEBUG] Finished temp file deletion");

                                if (File.Exists(mergedPdf))
                                {
                                    if (File.Exists(pdfFilePath))
                                        File.Delete(pdfFilePath);
                                    File.Move(mergedPdf, pdfFilePath);
                                }
                                Console.WriteLine($"Merged and replaced {pdfFilePath}");
                            }
                            success++;
                            Console.WriteLine($"[DEBUG] Success count: {success}");
                        }
                        catch (Exception ex)
                        {
                            fail++;
                            Console.WriteLine($"[ERROR] Failed to convert: {selectedFiles[i]}\nError: {ex}");
                        }
                        finally
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            KillWkhtmltopdfProcesses();
                            Dispatcher.Invoke(() => ProgressBar.Value = i + 1);
                            Console.WriteLine($"[DEBUG] Cleanup complete for file {i + 1}");
                        }
                    }
                    Console.WriteLine($"[DEBUG] Batch loop finished. Success: {success}, Fail: {fail}");
                });
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Exception in ConvertButton_Click outer: {ex}");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                isConverting = false;
                ProgressBar.Visibility = Visibility.Collapsed;
                ConvertButton.IsEnabled = true;
                Console.WriteLine("[DEBUG] ConvertButton_Click finished");
            }
        }

        // Appends attachments as PDFs to the main email PDF
        private void AppendAttachmentsToPdf(string mainPdfPath, List<Storage.Attachment> attachments, SynchronizedConverter converter)
        {
            Console.WriteLine($"Total attachments: {attachments.Count}");
            string allNames = string.Join(", ", attachments.ConvertAll(a => a.FileName ?? "(no name)"));
            Console.WriteLine($"Attachment names: {allNames}");
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
                            // Also add to cleanup list so the Office-generated PDF gets deleted
                            Console.WriteLine($"[ATTACH] Adding Office-generated PDF to cleanup list: {attPdf}");
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

        // Helper to create a single-page PDF with a header text at the top center using iText
        private void AddHeaderPdf(string pdfPath, string headerText)
        {
            using (var writer = new iText.Kernel.Pdf.PdfWriter(pdfPath))
            using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
            using (var doc = new iText.Layout.Document(pdf))
            {
                var p = new iText.Layout.Element.Paragraph(headerText)
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                    .SetFontSize(18);
                doc.Add(p);
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
                        Marshal.ReleaseComObject(doc);
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        result = true;
                    }
                    else if (ext == ".xls" || ext == ".xlsx")
                    {
                        var excelApp = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
                        Microsoft.Office.Interop.Excel.Workbook wb = null;
                        try
                        {
                            workbooks = excelApp.Workbooks;
                            wb = workbooks.Open(inputPath);
                            wb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPdf);
                        }
                        finally
                        {
                            if (wb != null) wb.Close(false);
                            if (wb != null) Marshal.ReleaseComObject(wb);
                            if (workbooks != null) Marshal.ReleaseComObject(workbooks);
                            if (excelApp != null) excelApp.Quit();
                            if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        result = true;
                    }
                }
                catch (Exception ex)
                {
                    threadEx = ex;
                }
            }); thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();            // Give Office extra time to release the generated PDF file
            if (result)
            {
                Console.WriteLine($"[Interop] Waiting for Office to release PDF file: {outputPdf}");

                // Wait longer and verify the PDF is not locked
                for (int i = 0; i < 10; i++)
                {
                    System.Threading.Thread.Sleep(500);

                    // Try to open the PDF file to verify it's not locked
                    try
                    {
                        using (var fs = new FileStream(outputPdf, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            // If we can open it, it's not locked
                            Console.WriteLine($"[Interop] PDF file ready after {(i + 1) * 500}ms: {outputPdf}");
                            break;
                        }
                    }
                    catch (IOException)
                    {
                        if (i == 9) // Last attempt
                        {
                            Console.WriteLine($"[Interop][WARNING] PDF file may still be locked after 5 seconds: {outputPdf}");
                        }
                    }
                }
            }

            if (threadEx != null)
            {
                Console.WriteLine($"[Interop] Office to PDF conversion failed: {threadEx.Message}");
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

        private void ResetApplication()
        {
            // Option 1: Restart the application programmatically
            System.Diagnostics.Process.Start(System.Windows.Application.ResourceAssembly.Location);
            System.Windows.Application.Current.Shutdown();
        }

        // Returns all ContentIds referenced as inline images in the HTML
        private HashSet<string> GetInlineContentIds(string html)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(html)) return set;
            // Match src='cid:...' or src="cid:..." and optional angle brackets
            var regex = new Regex("<img[^>]+src=['\"]cid:<?([^'\">]+)>?['\"]", RegexOptions.IgnoreCase);
            foreach (Match match in regex.Matches(html))
            {
                if (match.Groups.Count > 1)
                    set.Add(match.Groups[1].Value.Trim('<', '>', '\"', '\'', ' '));
            }
            return set;
        }        // Robust file deletion with retries and Excel-specific handling
        private void RobustDeleteFile(string filePath, int maxRetries = 5, int delayMs = 500)
        {
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);

                        // Wait a moment and verify the file is actually gone
                        System.Threading.Thread.Sleep(100);

                        if (!File.Exists(filePath))
                        {
                            Console.WriteLine($"[CLEANUP] Successfully deleted temp file: {filePath}");
                            return;
                        }
                        else
                        {
                            Console.WriteLine($"[CLEANUP] File.Delete() succeeded but file still exists (attempt {i + 1}): {filePath}");

                            // For Excel files, try killing Excel processes on the last few attempts
                            if (i >= maxRetries - 2 && (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) || filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)))
                            {
                                Console.WriteLine($"[CLEANUP] Excel file locked, attempting to kill Excel processes...");
                                KillExcelProcesses();
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[CLEANUP] File no longer exists: {filePath}");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[CLEANUP] Delete attempt {i + 1} failed for {filePath}: {ex.Message}");
                }

                System.Threading.Thread.Sleep(delayMs);
            }

            if (File.Exists(filePath))
            {
                Console.WriteLine($"[CLEANUP][ERROR] File still exists after all retries: {filePath}");
                Console.WriteLine($"[CLEANUP] This file may be locked by Excel.exe or another process.");
            }
        }

        // Kill Excel processes that may be holding files open
        private void KillExcelProcesses()
        {
            try
            {
                var excelProcesses = Process.GetProcessesByName("EXCEL");
                foreach (var process in excelProcesses)
                {
                    try
                    {
                        Console.WriteLine($"[CLEANUP] Killing Excel process PID: {process.Id}");
                        process.Kill();
                        process.WaitForExit(2000);
                        Console.WriteLine($"[CLEANUP] Successfully killed Excel process PID: {process.Id}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[CLEANUP] Failed to kill Excel process PID {process.Id}: {ex.Message}");
                    }
                }

                if (excelProcesses.Length == 0)
                {
                    Console.WriteLine("[CLEANUP] No Excel processes found to kill.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[CLEANUP] Failed to enumerate Excel processes: {ex.Message}");
            }
        }
    }
}