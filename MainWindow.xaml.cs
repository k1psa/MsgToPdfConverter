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
using Microsoft.VisualBasic.FileIO; // Add this at the top for FileSystem.DeleteFile
using MsgToPdfConverter.Services;

namespace MsgToPdfConverter
{
    public partial class MainWindow : Window
    {
        private List<string> selectedFiles = new List<string>();
        private int convertButtonClickCount = 0;
        private bool isConverting = false;
        private bool cancellationRequested = false;
        private string selectedOutputFolder = null;
        private bool extractOriginalOnly = false;
        private bool deleteMsgAfterConversion = false;
        private bool isPinned = false;
        private EmailConverterService _emailService = new EmailConverterService();
        private AttachmentService _attachmentService;
        private OutlookImportService _outlookImportService = new OutlookImportService();

        public MainWindow()
        {
            InitializeComponent();
            
            // Initialize AttachmentService with PDF service methods
            _attachmentService = new AttachmentService(
                (path, text, _) => PdfService.AddHeaderPdf(path, text), // Adapter for the different signature
                OfficeConversionService.TryConvertOfficeToPdf,
                PdfAppendTest.AppendPdfs,
                _emailService
            );
            
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
            var newFiles = FileDialogHelper.OpenMsgFileDialog();
            if (newFiles != null && newFiles.Count > 0)
            {
                foreach (var file in newFiles)
                {
                    // Only add if not already in the list
                    if (!selectedFiles.Contains(file))
                    {
                        selectedFiles.Add(file);
                        FilesListBox.Items.Add(file);
                    }
                }
            }
            UpdateFileCountAndButtons();
        }

        private void ClearListButton_Click(object sender, RoutedEventArgs e)
        {
            selectedFiles.Clear();
            FilesListBox.Items.Clear();
            UpdateFileCountAndButtons();
        }

        private void SelectOutputFolderButton_Click(object sender, RoutedEventArgs e)
        {
            var folder = FileDialogHelper.OpenFolderDialog();
            if (!string.IsNullOrEmpty(folder))
            {
                selectedOutputFolder = folder;
                OutputFolderLabel.Text = $"Output Folder: {selectedOutputFolder}";
            }
        }

        private void ClearOutputFolderButton_Click(object sender, RoutedEventArgs e)
        {
            selectedOutputFolder = null;
            OutputFolderLabel.Text = "(Default: Same as .msg file)";
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

        private void ConfigureDinkToPdfPath(PdfTools pdfTools)
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
                        DirectoryCopy(librariesArchPath, tempArchPath, true);
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

        private void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);
            if (!dir.Exists)
                throw new DirectoryNotFoundException($"Source directory does not exist or could not be found: {sourceDirName}");

            DirectoryInfo[] dirs = dir.GetDirectories();
            Directory.CreateDirectory(destDirName);

            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string tempPath = Path.Combine(destDirName, file.Name);
                file.CopyTo(tempPath, true);
            }

            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string tempPath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, tempPath, copySubDirs);
                }
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
            Console.WriteLine($"[DEBUG] Convert button pressed {convertButtonClickCount} time(s)"); try
            {
                Console.WriteLine("[DEBUG] Disabling UI and showing progress");
                SetProcessingState(true);
                ProgressBar.Minimum = 0;
                ProgressBar.Maximum = selectedFiles.Count;
                ProgressBar.Value = 0; int success = 0, fail = 0, processed = 0;
                // Use the field instead of the removed checkbox
                bool appendAttachments = AppendAttachmentsCheckBox.IsChecked == true;
                // Use the field for extract original only
                bool extractOriginalOnlyLocal = extractOriginalOnly;
                Console.WriteLine($"[DEBUG] appendAttachments: {appendAttachments}");
                await Task.Run(() =>
                {
                    Console.WriteLine("[DEBUG] Starting batch loop");
                    for (int i = 0; i < selectedFiles.Count; i++)
                    {
                        if (cancellationRequested)
                        {
                            Console.WriteLine("[DEBUG] Cancellation requested, breaking loop");
                            break;
                        }
                        processed++;
                        Console.WriteLine($"[DEBUG] Loop index: {i}");
                        string msgFilePath = selectedFiles[i];
                        Storage.Message msg = null;
                        try
                        {
                            // Update status on UI thread
                            Dispatcher.Invoke(() =>
                            {
                                ProcessingStatusLabel.Foreground = System.Windows.Media.Brushes.Blue;
                                ProcessingStatusLabel.Text = $"Processing file {processed}/{selectedFiles.Count}: {Path.GetFileName(selectedFiles[i])}";
                                // Progress bar will be updated when file is completed
                            });
                            Console.WriteLine($"[TASK] Processing file {i + 1} of {selectedFiles.Count}: {selectedFiles[i]}");
                            msg = new Storage.Message(msgFilePath);
                            Console.WriteLine($"[TASK] Loaded MSG: {msgFilePath}");
                            string datePart = msg.SentOn.HasValue ? msg.SentOn.Value.ToString("yyyy-MM-dd_HHmmss") : DateTime.Now.ToString("yyyy-MM-dd_HHmms");
                            string baseName = Path.GetFileNameWithoutExtension(msgFilePath);
                            string dir = !string.IsNullOrEmpty(selectedOutputFolder) ? selectedOutputFolder : Path.GetDirectoryName(msgFilePath);
                            string pdfFilePath = Path.Combine(dir, $"{baseName} - {datePart}.pdf");
                            if (File.Exists(pdfFilePath))
                            {
                                try { File.Delete(pdfFilePath); Console.WriteLine($"[DEBUG] Deleted old PDF: {pdfFilePath}"); } catch (Exception ex) { Console.WriteLine($"[DEBUG] Could not delete old PDF: {ex.Message}"); }
                            }
                            Console.WriteLine($"[TASK] Output PDF path: {pdfFilePath}");
                            // Replace direct BuildEmailHtml call with service
                            var htmlResult = _emailService.BuildEmailHtmlWithInlineImages(msg, extractOriginalOnly);
                            string htmlWithHeader = htmlResult.Html;
                            List<string> tempInlineFiles = htmlResult.TempFiles;
                            Console.WriteLine($"[TASK] Built HTML for: {msgFilePath}");
                            Console.WriteLine($"[TASK] HTML length: {htmlWithHeader?.Length ?? 0}");                            // Write HTML to a temp file
                            string tempHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".html");
                            File.WriteAllText(tempHtmlPath, htmlWithHeader, System.Text.Encoding.UTF8);
                            Console.WriteLine($"[DEBUG] Written HTML to temp file: {tempHtmlPath}");
                            // Find all inline ContentIds
                            var inlineContentIds = _emailService.GetInlineContentIds(htmlWithHeader);
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
                            GC.WaitForPendingFinalizers(); if (appendAttachments && msg.Attachments != null && msg.Attachments.Count > 0)
                            {
                                Console.WriteLine($"[DEBUG] Processing attachments for {pdfFilePath}");
                                Console.WriteLine($"[DEBUG] Total attachments found: {msg.Attachments.Count}"); var typedAttachments = new List<Storage.Attachment>();
                                var nestedMessages = new List<Storage.Message>();

                                foreach (var att in msg.Attachments)
                                {
                                    if (att is Storage.Attachment a)
                                    {
                                        Console.WriteLine($"[DEBUG] Examining attachment: {a.FileName} (IsInline: {a.IsInline}, ContentId: {a.ContentId})");

                                        if ((a.IsInline == true) || (!string.IsNullOrEmpty(a.ContentId) && inlineContentIds.Contains(a.ContentId.Trim('<', '>', '\"', '\'', ' '))))
                                        {
                                            Console.WriteLine($"[DEBUG] Skipping inline attachment: {a.FileName} (ContentId: {a.ContentId}, IsInline: {a.IsInline})");
                                            continue;
                                        }

                                        Console.WriteLine($"[DEBUG] Adding to processing list: {a.FileName}");
                                        typedAttachments.Add(a);
                                    }
                                    else if (att is Storage.Message nestedMsg)
                                    {
                                        Console.WriteLine($"[DEBUG] Found nested MSG: {nestedMsg.Subject ?? "No Subject"}");
                                        Console.WriteLine($"[DEBUG] Adding nested MSG to processing list");
                                        nestedMessages.Add(nestedMsg);
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[DEBUG] Unknown attachment type: {att?.GetType().Name}");
                                    }
                                }
                                Console.WriteLine($"[DEBUG] Attachments to process: {typedAttachments.Count}");
                                Console.WriteLine($"[DEBUG] Nested MSG files to process: {nestedMessages.Count}");
                                var allPdfFiles = new List<string> { pdfFilePath };
                                var allTempFiles = new List<string>();
                                string tempDir = Path.GetDirectoryName(pdfFilePath);

                                // Process regular attachments using the new helper method
                                int totalAttachments = typedAttachments.Count;
                                for (int attIndex = 0; attIndex < typedAttachments.Count; attIndex++)
                                {
                                    var att = typedAttachments[attIndex];
                                    string attName = att.FileName ?? "attachment";
                                    string attPath = Path.Combine(tempDir, attName);
                                    string headerText = $"Attachment {attIndex + 1}/{totalAttachments} - {attName}";
                                    try
                                    {
                                        File.WriteAllBytes(attPath, att.Data);
                                        allTempFiles.Add(attPath);
                                        string finalAttachmentPdf = _attachmentService.ProcessSingleAttachment(att, attPath, tempDir, headerText, allTempFiles);

                                        if (finalAttachmentPdf != null)
                                            allPdfFiles.Add(finalAttachmentPdf);
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"[ATTACH] Error processing attachment {attName}: {ex.Message}");
                                        string errorPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                                        PdfService.AddHeaderPdf(errorPdf, headerText + $"\n(Error: {ex.Message})");
                                        allPdfFiles.Add(errorPdf);
                                        allTempFiles.Add(errorPdf);
                                    }
                                }

                                // Process nested MSG files recursively (including their attachments)
                                Console.WriteLine($"[MSG] Processing {nestedMessages.Count} nested MSG files recursively");
                                foreach (var nestedMsg in nestedMessages)
                                {
                                    // This will recursively process the nested MSG and all its attachments
                                    _attachmentService.ProcessMsgAttachmentsRecursively(nestedMsg, allPdfFiles, allTempFiles, tempDir, extractOriginalOnly, 1);
                                }

                                string mergedPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(pdfFilePath) + "_merged.pdf");
                                Console.WriteLine($"[DEBUG] Before PDF merge: {string.Join(", ", allPdfFiles)} -> {mergedPdf}");
                                PdfAppendTest.AppendPdfs(allPdfFiles, mergedPdf); Console.WriteLine("[DEBUG] After PDF merge");
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                foreach (var f in allTempFiles)
                                {
                                    try
                                    {
                                        if (File.Exists(f) && !string.Equals(f, mergedPdf, StringComparison.OrdinalIgnoreCase) && !string.Equals(f, pdfFilePath, StringComparison.OrdinalIgnoreCase))
                                        {
                                            FileService.RobustDeleteFile(f);
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
                            if (msg != null && msg is IDisposable disposableMsg)
                            {
                                disposableMsg.Dispose();
                                Console.WriteLine($"[DEBUG] Disposed Storage.Message for: {msgFilePath}");
                            }
                            msg = null;
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            // Delete .msg file after conversion if the checkbox is checked
                            if (deleteMsgAfterConversion && File.Exists(msgFilePath))
                            {
                                try
                                {
                                    // Move to Recycle Bin instead of permanent delete
                                    FileService.MoveFileToRecycleBin(msgFilePath);
                                    Console.WriteLine($"[DELETE] Moved .msg file to Recycle Bin: {msgFilePath}");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"[DELETE] Could not move {msgFilePath} to Recycle Bin: {ex.Message}");
                                }
                            }
                            KillWkhtmltopdfProcesses();
                            Dispatcher.Invoke(() => ProgressBar.Value = i + 1);
                            Console.WriteLine($"[DEBUG] Cleanup complete for file {i + 1}");
                        }
                    }
                    Console.WriteLine($"[DEBUG] Batch loop finished. Success: {success}, Fail: {fail}, Processed: {processed}");
                });

                // Show final results
                string statusMessage;
                if (cancellationRequested)
                {
                    statusMessage = $"Processing cancelled. Processed {processed} files. Success: {success}, Failed: {fail}";
                }
                else
                {
                    statusMessage = $"Processing completed. Total files: {selectedFiles.Count}, Success: {success}, Failed: {fail}";
                }

                MessageBox.Show(statusMessage, "Processing Results", MessageBoxButton.OK,
                    fail > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DEBUG] Exception in ConvertButton_Click outer: {ex}");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetProcessingState(false);
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
                        if (OfficeConversionService.TryConvertOfficeToPdf(attPath, attPdf))
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
                        var zipFiles = Directory.GetFiles(extractDir, "*.*", System.IO.SearchOption.AllDirectories);
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
                                if (OfficeConversionService.TryConvertOfficeToPdf(zf, zfPdf))
                                {
                                    tempPdfFiles.Add(zfPdf);
                                    Console.WriteLine($"[ATTACH] ZIP Office converted: {zfPdf}");
                                }
                                else
                                {
                                    PdfService.AddPlaceholderPdf(zfPdf, $"Could not convert attachment: {Path.GetFileName(zf)}");
                                    Console.WriteLine($"[ATTACH] ZIP Office failed to convert: {zf}");
                                }
                            }
                            else
                            {
                                PdfService.AddPlaceholderPdf(zfPdf, $"Unsupported attachment: {Path.GetFileName(zf)}");
                                tempPdfFiles.Add(zfPdf);
                                Console.WriteLine($"[ATTACH] ZIP unsupported: {zf}");
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
                }
                catch (Exception ex)
                {
                    PdfService.AddPlaceholderPdf(attPdf, $"Error processing attachment: {ex.Message}");
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

        private void FilesListBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] FilesListBox_KeyDown triggered. Key: {e.Key}, SelectedItems: {FilesListBox.SelectedItems.Count}");
            if (e.Key == System.Windows.Input.Key.Delete && FilesListBox.SelectedItems.Count > 0)
            {
                var itemsToRemove = new List<string>();
                foreach (var item in FilesListBox.SelectedItems)
                {
                    itemsToRemove.Add(item as string);
                }

                // Remove selected files from the list (no confirmation required)
                foreach (var item in itemsToRemove)
                {
                    FilesListBox.Items.Remove(item);
                    selectedFiles.Remove(item);
                }
                UpdateFileCountAndButtons();
                e.Handled = true; // Suppress default behavior
            }
        }

        private void UpdateFileCountAndButtons()
        {
            int fileCount = FilesListBox.Items.Count;
            FileCountLabel.Text = $"Files selected: {fileCount}";
            ConvertButton.IsEnabled = fileCount > 0 && !isConverting;
        }

        private void SetProcessingState(bool processing)
        {
            isConverting = processing;            // Disable/enable main buttons
            SelectFilesButton.IsEnabled = !processing;
            ClearListButton.IsEnabled = !processing;
            ConvertButton.IsEnabled = !processing && FilesListBox.Items.Count > 0;
            AppendAttachmentsCheckBox.IsEnabled = !processing;
            FilesListBox.IsEnabled = !processing;

            // Show/hide cancel button
            CancelButton.Visibility = processing ? Visibility.Visible : Visibility.Collapsed;

            // Show/hide progress elements
            ProgressBar.Visibility = processing ? Visibility.Visible : Visibility.Collapsed;
            ProcessingStatusLabel.Visibility = processing ? Visibility.Visible : Visibility.Collapsed;

            if (!processing)
            {
                ProcessingStatusLabel.Text = "";
                cancellationRequested = false;
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            cancellationRequested = true;
            CancelButton.IsEnabled = false;
            ProcessingStatusLabel.Text = "Cancelling... Please wait.";
            ProcessingStatusLabel.Foreground = System.Windows.Media.Brushes.Red;
        }

        private void FilesListBox_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void FilesListBox_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private async void FilesListBox_Drop(object sender, DragEventArgs e)
        {
            // 1. Standard file/folder drop
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] droppedItems = (string[])e.Data.GetData(DataFormats.FileDrop);
                var newFiles = new List<string>();
                foreach (string item in droppedItems)
                {
                    if (File.Exists(item))
                    {
                        if (Path.GetExtension(item).ToLowerInvariant() == ".msg")
                        {
                            if (!selectedFiles.Contains(item))
                                newFiles.Add(item);
                        }
                    }
                    else if (Directory.Exists(item))
                    {
                        var msgFiles = Directory.GetFiles(item, "*.msg", System.IO.SearchOption.AllDirectories);
                        foreach (string msgFile in msgFiles)
                        {
                            if (!selectedFiles.Contains(msgFile))
                                newFiles.Add(msgFile);
                        }
                    }
                }
                foreach (string file in newFiles)
                {
                    selectedFiles.Add(file);
                    FilesListBox.Items.Add(file);
                }
                UpdateFileCountAndButtons();
                return;
            }

            // 2. Outlook email drag-and-drop support (improved robust approach)
            if (e.Data.GetDataPresent("FileGroupDescriptorW") || e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                try
                {
                    Console.WriteLine($"[DND] Starting Outlook email drag-and-drop processing");

                    // Log all available data formats for debugging
                    var availableFormats = e.Data.GetFormats();
                    Console.WriteLine($"[DND] Available data formats: {string.Join(", ", availableFormats)}");

                    // Ensure main window is properly activated and focused before showing dialog
                    this.Activate();
                    this.Focus();

                    // Temporarily set topmost to ensure dialog appears in front
                    bool wasTopmost = this.Topmost;
                    this.Topmost = true;

                    // Small delay to ensure window activation
                    await System.Threading.Tasks.Task.Delay(50);

                    // Restore original topmost state
                    this.Topmost = wasTopmost;

                    // Use selectedOutputFolder if set, otherwise prompt
                    string outputFolder = !string.IsNullOrEmpty(selectedOutputFolder)
                        ? selectedOutputFolder
                        : FileDialogHelper.OpenFolderDialog();
                    if (string.IsNullOrEmpty(outputFolder))
                        return;

                    // Debug: Show all available data formats
                    Console.WriteLine("[DND] Available data formats:");
                    foreach (string format in e.Data.GetFormats())
                    {
                        Console.WriteLine($"[DND]   - {format}");
                    }

                    string[] fileNames = null;
                    if (e.Data.GetDataPresent("FileGroupDescriptorW"))
                    {
                        using (var stream = (System.IO.MemoryStream)e.Data.GetData("FileGroupDescriptorW"))
                        {
                            fileNames = GetFileNamesFromFileGroupDescriptorW(stream);
                        }
                    }
                    else if (e.Data.GetDataPresent("FileGroupDescriptor"))
                    {
                        using (var stream = (System.IO.MemoryStream)e.Data.GetData("FileGroupDescriptor"))
                        {
                            fileNames = GetFileNamesFromFileGroupDescriptor(stream);
                        }
                    }

                    if (fileNames == null || fileNames.Length == 0)
                    {
                        Console.WriteLine("[DND] No file names found in drag data");
                        return;
                    }

                    Console.WriteLine($"[DND] Found {fileNames.Length} file(s) in drag data:");
                    for (int idx = 0; idx < fileNames.Length; idx++)
                    {
                        Console.WriteLine($"[DND]   {idx}: {fileNames[idx]}");
                    }

                    var tempFiles = new List<string>();
                    var skippedFiles = new List<string>();
                    var usedFilenames = new HashSet<string>(); // Track filenames used in this batch

                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        string fileName = fileNames[i];
                        if (!fileName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                            fileName += ".msg";

                        // Sanitize filename to remove illegal characters
                        fileName = FileService.SanitizeFileName(fileName);

                        // Make filename unique if it already exists or was used in this batch
                        string destPath = Path.Combine(outputFolder, fileName);
                        int counter = 1;
                        while (File.Exists(destPath) || usedFilenames.Contains(Path.GetFileName(destPath)))
                        {
                            string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                            string extension = Path.GetExtension(fileName);
                            string uniqueFileName = $"{nameWithoutExt}_{counter}{extension}";
                            destPath = Path.Combine(outputFolder, uniqueFileName);
                            counter++;
                        }

                        // Remember this filename for the current batch
                        usedFilenames.Add(Path.GetFileName(destPath));

                        // Try FileContents formats systematically
                        Console.WriteLine($"[DND] Processing file {i + 1}/{fileNames.Length}: {fileName}");

                        bool success = false;

                        // For multiple files, try indexed format first, then fallback
                        if (fileNames.Length > 1)
                        {
                            string indexedFormat = $"FileContents{i}";
                            Console.WriteLine($"[DND] Trying indexed format: {indexedFormat}");

                            if (e.Data.GetDataPresent(indexedFormat))
                            {
                                try
                                {
                                    using (var fileStream = (System.IO.MemoryStream)e.Data.GetData(indexedFormat))
                                    {
                                        if (fileStream != null && fileStream.Length > 0)
                                        {
                                            using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                                            {
                                                fileStream.Position = 0;
                                                fileStream.WriteTo(fs);
                                            }
                                            tempFiles.Add(destPath);
                                            success = true;
                                            Console.WriteLine($"[DND] Successfully saved {fileName} using {indexedFormat}");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"[DND] Empty stream for {indexedFormat}");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"[DND] Error writing {fileName} from {indexedFormat}: {ex.Message}");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"[DND] Format {indexedFormat} not present");
                            }
                        }

                        // Try non-indexed format as fallback
                        if (!success && e.Data.GetDataPresent("FileContents"))
                        {
                            Console.WriteLine($"[DND] Trying non-indexed format: FileContents");
                            try
                            {
                                using (var fileStream = (System.IO.MemoryStream)e.Data.GetData("FileContents"))
                                {
                                    if (fileStream != null && fileStream.Length > 0)
                                    {
                                        using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                                        {
                                            fileStream.Position = 0;
                                            fileStream.WriteTo(fs);
                                        }
                                        tempFiles.Add(destPath);
                                        success = true;
                                        Console.WriteLine($"[DND] Successfully saved {fileName} using FileContents");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[DND] Empty stream for FileContents");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[DND] Error writing {fileName} from FileContents: {ex.Message}");
                            }
                        }
                        if (!success)
                        {
                            // Try alternate Outlook formats (rare)
                            string[] altFormats = { "RenPrivateItem", "Attachment" };
                            foreach (var alt in altFormats)
                            {
                                if (e.Data.GetDataPresent(alt))
                                {
                                    try
                                    {
                                        using (var fileStream = (System.IO.MemoryStream)e.Data.GetData(alt))
                                        using (var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                                        {
                                            fileStream.WriteTo(fs);
                                        }
                                        tempFiles.Add(destPath);
                                        success = true;
                                        break;
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"[DND] Error writing {fileName} from {alt}: {ex.Message}");
                                    }
                                }
                            }
                        }
                        // Try FileDrop as a last resort (rare, but sometimes Outlook provides it)
                        if (!success && e.Data.GetDataPresent(DataFormats.FileDrop))
                        {
                            try
                            {
                                string[] dropped = (string[])e.Data.GetData(DataFormats.FileDrop);
                                foreach (var path in dropped)
                                {
                                    if (File.Exists(path) && Path.GetExtension(path).ToLowerInvariant() == ".msg")
                                    {
                                        File.Copy(path, destPath, true);
                                        tempFiles.Add(destPath);
                                        success = true;
                                        break;
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[DND] Error copying from FileDrop: {ex.Message}");
                            }
                        }
                        if (!success)
                        {
                            // Fallback: Try to use Outlook Interop to save the selected email(s) as .msg
                            Console.WriteLine($"[DND] Trying Outlook Interop fallback for file {i + 1}");
                            try
                            {
                                var outlookApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Outlook.Application") as Microsoft.Office.Interop.Outlook.Application;
                                if (outlookApp != null)
                                {
                                    var explorer = outlookApp.ActiveExplorer();
                                    if (explorer != null && explorer.Selection != null && explorer.Selection.Count > 0)
                                    {
                                        // For multiple files, try to get the specific selection item by index
                                        int selectionIndex = Math.Min(i + 1, explorer.Selection.Count);
                                        var mailItem = explorer.Selection[selectionIndex] as Microsoft.Office.Interop.Outlook.MailItem;
                                        if (mailItem != null)
                                        {
                                            // Use the same filename logic as above, but ensure uniqueness
                                            string safeSubject = FileService.SanitizeFileName(mailItem.Subject ?? "untitled");
                                            string interopFileName = safeSubject;
                                            if (!interopFileName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                                                interopFileName += ".msg";

                                            // Apply the same uniqueness logic
                                            string interopDestPath = Path.Combine(outputFolder, interopFileName);
                                            int interopCounter = 1;
                                            while (File.Exists(interopDestPath) || usedFilenames.Contains(Path.GetFileName(interopDestPath)))
                                            {
                                                string nameWithoutExt = Path.GetFileNameWithoutExtension(interopFileName);
                                                string extension = Path.GetExtension(interopFileName);
                                                string uniqueFileName = $"{nameWithoutExt}_{interopCounter}{extension}";
                                                interopDestPath = Path.Combine(outputFolder, uniqueFileName);
                                                interopCounter++;
                                            }

                                            // Update the used filenames tracking
                                            usedFilenames.Add(Path.GetFileName(interopDestPath));

                                            mailItem.SaveAs(interopDestPath, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG);
                                            tempFiles.Add(interopDestPath);
                                            success = true;
                                            Console.WriteLine($"[DND] Successfully saved via Interop: {interopDestPath}");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"[DND] Selection item {selectionIndex} is not a MailItem");
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"[DND] No active Outlook explorer or selection");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($"[DND] Could not get Outlook application object");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"[DND][Interop] Could not save selected Outlook email: {ex.Message}");
                            }
                        }
                        if (!success)
                        {
                            skippedFiles.Add(fileName);
                        }
                    }

                    foreach (var file in tempFiles)
                    {
                        if (!selectedFiles.Contains(file))
                        {
                            selectedFiles.Add(file);
                            FilesListBox.Items.Add(file);
                        }
                    }
                    UpdateFileCountAndButtons();

                    if (skippedFiles.Count > 0)
                    {
                        MessageBox.Show(
                            $"Some emails could not be added due to missing data from Outlook.\n\nPossible reasons:\n- The email is protected or encrypted\n- Outlook security settings\n- Outlook version limitations\n- The email is a meeting request or special item\n\nTry dragging the email(s) to a folder first, then add the .msg file.\n\nSkipped:\n{string.Join("\n", skippedFiles)}",
                            "Outlook Drag-and-Drop",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error processing Outlook email drop: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                return;
            }

            // 3. If all else fails, inform the user
            MessageBox.Show(
                "Could not extract email from Outlook drag-and-drop.\n\n" +
                "This may be due to Outlook version or security settings.\n" +
                "Try dragging the email to a folder first, then add the .msg file.",
                "Outlook Drag-and-Drop Not Supported",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        // Helper: Parse file names from FileGroupDescriptorW (Unicode)
        private string[] GetFileNamesFromFileGroupDescriptorW(Stream stream)
        {
            var fileNames = new List<string>();
            using (var reader = new BinaryReader(stream, System.Text.Encoding.Unicode))
            {
                stream.Position = 0;
                // FILEGROUPDESCRIPTORW starts with a 4-byte count
                int count = reader.ReadInt32();
                for (int i = 0; i < count; i++)
                {
                    // Skip 76 bytes to the file name (see FILEDESCRIPTORW struct)
                    stream.Position = 4 + i * 592 + 76;
                    // Read up to 520 bytes (260 WCHARs)
                    var nameBytes = reader.ReadBytes(520);
                    string name = System.Text.Encoding.Unicode.GetString(nameBytes).TrimEnd('\0');
                    fileNames.Add(name);
                }
            }
            return fileNames.ToArray();
        }

        // Helper: Parse file names from FileGroupDescriptor (ANSI)
        private string[] GetFileNamesFromFileGroupDescriptor(Stream stream)
        {
            var fileNames = new List<string>();
            using (var reader = new BinaryReader(stream, System.Text.Encoding.Default))
            {
                stream.Position = 0;
                // FILEGROUPDESCRIPTOR starts with a 4-byte count
                int count = reader.ReadInt32();
                for (int i = 0; i < count; i++)
                {
                    // Skip 76 bytes to the file name (see FILEDESCRIPTOR struct)
                    stream.Position = 4 + i * 592 + 76;
                    // Read up to 260 bytes (MAX_PATH)
                    var nameBytes = reader.ReadBytes(260);
                    string name = System.Text.Encoding.Default.GetString(nameBytes).TrimEnd('\0');
                    fileNames.Add(name);
                }
            }
            return fileNames.ToArray();
        }

        private void ProcessMsgAttachmentsRecursively(Storage.Message msg, List<string> allPdfFiles, List<string> allTempFiles, string tempDir, bool extractOriginalOnly, int depth = 0, int maxDepth = 5)
        {
            // Prevent infinite recursion
            if (depth > maxDepth)
            {
                Console.WriteLine($"[MSG] Stopping recursion at depth {depth} - max depth reached");
                return;
            }

            Console.WriteLine($"[MSG] Processing message recursively at depth {depth}: {msg.Subject ?? "No Subject"}");

            // First, create PDF for the nested MSG body content (if we're processing a nested message)
            if (depth > 0)
            {
                string msgSubject = msg.Subject ?? $"nested_msg_depth_{depth}";
                string headerText = $"Nested Email (Depth {depth}): {msgSubject}";

                try
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Creating PDF for nested message body: {msgSubject}");

                    // Use new inline image logic for nested emails
                    var (nestedHtml, tempFiles) = _emailService.BuildEmailHtmlWithInlineImages(msg, extractOriginalOnly);
                    allTempFiles.AddRange(tempFiles); // Track temp image files for cleanup

                    string nestedPdf = Path.Combine(tempDir, $"depth{depth}_{Guid.NewGuid()}_nested_msg.pdf");
                    string nestedHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".html");
                    File.WriteAllText(nestedHtmlPath, nestedHtml, System.Text.Encoding.UTF8);

                    var startInfo = new ProcessStartInfo
                    {
                        FileName = System.Reflection.Assembly.GetExecutingAssembly().Location,
                        Arguments = $"--html2pdf \"{nestedHtmlPath}\" \"{nestedPdf}\"",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };

                    using (var process = Process.Start(startInfo))
                    {
                        process.WaitForExit();
                        if (process.ExitCode == 0)
                        {
                            // Add header and merge
                            string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                            PdfService.AddHeaderPdf(headerPdf, headerText);
                            string finalNestedPdf = Path.Combine(tempDir, Guid.NewGuid() + "_nested_merged.pdf");
                            PdfAppendTest.AppendPdfs(new List<string> { headerPdf, nestedPdf }, finalNestedPdf);

                            allPdfFiles.Add(finalNestedPdf);
                            allTempFiles.Add(headerPdf);
                            allTempFiles.Add(nestedPdf);
                            allTempFiles.Add(finalNestedPdf);

                            Console.WriteLine($"[MSG] Depth {depth} - Successfully created PDF for nested MSG body: {msgSubject}");
                        }
                        else
                        {
                            Console.WriteLine($"[MSG] Depth {depth} - Failed to convert nested MSG body to PDF: {msgSubject}");
                        }
                    }

                    File.Delete(nestedHtmlPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Error processing nested MSG body {msgSubject}: {ex.Message}");
                    string errorPdf = Path.Combine(tempDir, Guid.NewGuid() + "_msg_error.pdf");
                    PdfService.AddHeaderPdf(errorPdf, headerText + $"\n(Error: {ex.Message})");
                    allPdfFiles.Add(errorPdf);
                    allTempFiles.Add(errorPdf);
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

                    if ((a.IsInline == true) || (!string.IsNullOrEmpty(a.ContentId) && inlineContentIds.Contains(a.ContentId.Trim('<', '>', '\"', '\'', ' '))))
                    {
                        Console.WriteLine($"[MSG] Depth {depth} - Skipping inline attachment: {a.FileName}");
                        continue;
                    }

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
                var att = typedAttachments[attIndex];
                string attName = att.FileName ?? "attachment";
                string attPath = Path.Combine(tempDir, $"depth{depth}_{attName}");
                string headerText = $"Attachment (Depth {depth}): {attIndex + 1}/{totalAttachments} - {attName}";

                try
                {
                    File.WriteAllBytes(attPath, att.Data);
                    allTempFiles.Add(attPath);
                    string finalAttachmentPdf = ProcessSingleAttachment(att, attPath, tempDir, headerText, allTempFiles);

                    if (finalAttachmentPdf != null)
                        allPdfFiles.Add(finalAttachmentPdf);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[MSG] Depth {depth} - Error processing attachment {attName}: {ex.Message}");
                    string errorPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                    PdfService.AddHeaderPdf(errorPdf, headerText + $"\n(Error: {ex.Message})");
                    allPdfFiles.Add(errorPdf);
                    allTempFiles.Add(errorPdf);
                }
            }

            // Process nested MSG files recursively (this will handle both their body content and attachments)
            for (int msgIndex = 0; msgIndex < nestedMessages.Count; msgIndex++)
            {
                var nestedMsg = nestedMessages[msgIndex];
                Console.WriteLine($"[MSG] Depth {depth} - Recursively processing nested message {msgIndex + 1}/{nestedMessages.Count}: {nestedMsg.Subject ?? "No Subject"}");
                ProcessMsgAttachmentsRecursively(nestedMsg, allPdfFiles, allTempFiles, tempDir, extractOriginalOnly, depth + 1, maxDepth);
            }
        }

        private string ProcessSingleAttachment(Storage.Attachment att, string attPath, string tempDir, string headerText, List<string> allTempFiles)
        {
            string attName = att.FileName ?? "attachment";
            string ext = Path.GetExtension(attName).ToLowerInvariant();
            string attPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attName) + ".pdf");
            string finalAttachmentPdf = null;

            try
            {
                if (ext == ".pdf")
                {
                    string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                    PdfService.AddHeaderPdf(headerPdf, headerText);
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
                    if (OfficeConversionService.TryConvertOfficeToPdf(attPath, attPdf))
                    {
                        string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                        PdfService.AddHeaderPdf(headerPdf, headerText);
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                        PdfAppendTest.AppendPdfs(new List<string> { headerPdf, attPdf }, finalAttachmentPdf);
                        allTempFiles.Add(headerPdf);
                        allTempFiles.Add(attPdf);
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                    else
                    {
                        finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                        PdfService.AddHeaderPdf(finalAttachmentPdf, headerText + "\n(Conversion failed)");
                        allTempFiles.Add(finalAttachmentPdf);
                    }
                }
                else if (ext == ".zip")
                {
                    finalAttachmentPdf = _attachmentService.ProcessZipAttachment(attPath, tempDir, headerText, allTempFiles);
                }
                else
                {
                    finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                    PdfService.AddHeaderPdf(finalAttachmentPdf, headerText + "\n(Unsupported type)");
                    allTempFiles.Add(finalAttachmentPdf);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ATTACH] Error processing attachment {attName}: {ex.Message}");
                finalAttachmentPdf = Path.Combine(tempDir, Guid.NewGuid() + "_error.pdf");
                PdfService.AddHeaderPdf(finalAttachmentPdf, headerText + $"\n(Error: {ex.Message})");
                allTempFiles.Add(finalAttachmentPdf);
            }

            return finalAttachmentPdf;
        }

        private string ProcessZipAttachment(string attPath, string tempDir, string headerText, List<string> allTempFiles)
        {
            string extractDir = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(attPath));
            System.IO.Compression.ZipFile.ExtractToDirectory(attPath, extractDir);
            allTempFiles.Add(extractDir);

            var zipFiles = Directory.GetFiles(extractDir, "*.*", System.IO.SearchOption.AllDirectories);
            var zipPdfFiles = new List<string>();
            int zipFileIndex = 0;

            foreach (var zf in zipFiles)
            {
                zipFileIndex++;
                string zfPdf = Path.Combine(tempDir, Path.GetFileNameWithoutExtension(zf) + ".pdf");
                string zfExt = Path.GetExtension(zf).ToLowerInvariant();
                string zipHeader = $"{headerText} (ZIP {zipFileIndex}/{zipFiles.Length})";
                string finalZipPdf = null;

                if (zfExt == ".pdf")
                {
                    string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                    PdfService.AddHeaderPdf(headerPdf, zipHeader);
                    finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                    PdfAppendTest.AppendPdfs(new List<string> { headerPdf, zf }, finalZipPdf);
                    allTempFiles.Add(headerPdf);
                    allTempFiles.Add(finalZipPdf);
                }
                else if (zfExt == ".doc" || zfExt == ".docx" || zfExt == ".xls" || zfExt == ".xlsx")
                {
                    if (OfficeConversionService.TryConvertOfficeToPdf(zf, zfPdf))
                    {
                        string headerPdf = Path.Combine(tempDir, Guid.NewGuid() + "_header.pdf");
                        PdfService.AddHeaderPdf(headerPdf, zipHeader);
                        finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_merged.pdf");
                        PdfAppendTest.AppendPdfs(new List<string> { headerPdf, zfPdf }, finalZipPdf);
                        allTempFiles.Add(headerPdf);
                        allTempFiles.Add(zfPdf);
                        allTempFiles.Add(finalZipPdf);
                    }
                    else
                    {
                        PdfService.AddHeaderPdf(finalZipPdf, zipHeader + "\n(Conversion failed)");
                        allTempFiles.Add(finalZipPdf);
                    }
                }
                else
                {
                    finalZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_placeholder.pdf");
                    PdfService.AddHeaderPdf(finalZipPdf, zipHeader + "\n(Unsupported type)");
                    allTempFiles.Add(finalZipPdf);
                }
                if (finalZipPdf != null)
                    zipPdfFiles.Add(finalZipPdf);
            }

            // Merge all files from the ZIP into a single PDF
            if (zipPdfFiles.Count == 0)
            {
                // No processable files found in ZIP
                string placeholderPdf = Path.Combine(tempDir, Guid.NewGuid() + "_empty_zip.pdf");
                PdfService.AddHeaderPdf(placeholderPdf, headerText + "\n(Empty or no processable files in ZIP)");
                allTempFiles.Add(placeholderPdf);
                return placeholderPdf;
            }
            else if (zipPdfFiles.Count == 1)
            {
                // Only one file, return it directly
                return zipPdfFiles[0];
            }
            else
            {
                // Multiple files - merge them all into one comprehensive PDF
                string mergedZipPdf = Path.Combine(tempDir, Guid.NewGuid() + "_zip_merged.pdf");
                Console.WriteLine($"[ZIP] Merging {zipPdfFiles.Count} files from ZIP into single PDF: {mergedZipPdf}");
                PdfAppendTest.AppendPdfs(zipPdfFiles, mergedZipPdf);
                allTempFiles.Add(mergedZipPdf);
                return mergedZipPdf;
            }
        }

        private void OptionsButton_Click(object sender, RoutedEventArgs e)
        {
            var optionsWindow = new OptionsWindow(extractOriginalOnly, deleteMsgAfterConversion)
            {
                Owner = this
            };
            if (optionsWindow.ShowDialog() == true)
            {
                extractOriginalOnly = optionsWindow.ExtractOriginalOnly;
                deleteMsgAfterConversion = optionsWindow.DeleteMsgAfterConversion;
            }
        }

        private void AlwaysOnTopCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            this.Topmost = true;
        }

        private void AlwaysOnTopCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            this.Topmost = false;
        }

        private void PinButton_Click(object sender, RoutedEventArgs e)
        {
            isPinned = !isPinned;
            this.Topmost = isPinned;
            PinButton.Foreground = isPinned ? System.Windows.Media.Brushes.Red : System.Windows.Media.Brushes.Black;
            PinButton.Opacity = isPinned ? 1.0 : 0.7;
        }
    }
}