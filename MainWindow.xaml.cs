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
        private FileListService _fileListService = new FileListService();

        public MainWindow()
        {
            InitializeComponent();
            // Initialize AttachmentService with PDF service methods
            _attachmentService = new AttachmentService(
                (path, text, _) => PdfService.AddHeaderPdf(path, text),
                OfficeConversionService.TryConvertOfficeToPdf,
                PdfAppendTest.AppendPdfs,
                _emailService
            );
            EnvironmentService.CheckDotNetRuntime(this);
        }

        private void SelectFilesButton_Click(object sender, RoutedEventArgs e)
        {
            var newFiles = FileDialogHelper.OpenMsgFileDialog();
            if (newFiles != null && newFiles.Count > 0)
            {
                selectedFiles = _fileListService.AddFiles(selectedFiles, newFiles);
            }
            SyncFileListUI();
        }

        private void ClearListButton_Click(object sender, RoutedEventArgs e)
        {
            selectedFiles = _fileListService.ClearFiles();
            SyncFileListUI();
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
            PdfConversionService.KillWkhtmltopdfProcesses();
        }

        private void ConfigureDinkToPdfPath(PdfTools pdfTools)
        {
            PdfConversionService.ConfigureDinkToPdfPath(pdfTools);
        }

        private void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            FileService.DirectoryCopy(sourceDirName, destDirName, copySubDirs);
        }

        private void RunDinkToPdfConversion(HtmlToPdfDocument doc)
        {
            PdfConversionService.RunDinkToPdfConversion(doc);
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

        // Helper: Parse file names from FileGroupDescriptor (ANSI)
        // (Moved to OutlookImportService)
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
                selectedFiles = _fileListService.RemoveFiles(selectedFiles, itemsToRemove);
                SyncFileListUI();
                e.Handled = true; // Suppress default behavior
            }
        }

        // Syncs the FilesListBox UI with the selectedFiles list
        private void SyncFileListUI()
        {
            FilesListBox.Items.Clear();
            foreach (var file in selectedFiles)
            {
                FilesListBox.Items.Add(file);
            }
            UpdateFileCountAndButtons();
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

        private void FilesListBox_Drop(object sender, DragEventArgs e)
        {
            // 1. Standard file/folder drop
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] droppedItems = (string[])e.Data.GetData(DataFormats.FileDrop);
                foreach (string item in droppedItems)
                {
                    if (File.Exists(item))
                    {
                        selectedFiles = _fileListService.AddFiles(selectedFiles, new[] { item });
                    }
                    else if (Directory.Exists(item))
                    {
                        selectedFiles = _fileListService.AddFilesFromDirectory(selectedFiles, item);
                    }
                }
                SyncFileListUI();
                return;
            }

            // 2. Outlook email drag-and-drop support (all logic in service)
            if (e.Data.GetDataPresent("FileGroupDescriptorW") || e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                try
                {
                    string outputFolder = !string.IsNullOrEmpty(selectedOutputFolder)
                        ? selectedOutputFolder
                        : FileDialogHelper.OpenFolderDialog();
                    if (string.IsNullOrEmpty(outputFolder))
                        return;

                    var result = _outlookImportService.ExtractMsgFilesFromDragDrop(
                        e.Data,
                        outputFolder,
                        FileService.SanitizeFileName);

                    selectedFiles = _fileListService.AddFiles(selectedFiles, result.ExtractedFiles);
                    SyncFileListUI();

                    if (result.SkippedFiles.Count > 0)
                    {
                        MessageBox.Show(
                            $"Some emails could not be added due to missing data from Outlook.\n\nPossible reasons:\n- The email is protected or encrypted\n- Outlook security settings\n- Outlook version limitations\n- The email is a meeting request or special item\n\nTry dragging the email(s) to a folder first, then add the .msg file.\n\nSkipped:\n{string.Join("\n", result.SkippedFiles)}",
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

        private void PinButton_Click(object sender, RoutedEventArgs e)
        {
            isPinned = !isPinned;
            this.Topmost = isPinned;
            PinButton.Foreground = isPinned ? System.Windows.Media.Brushes.Red : System.Windows.Media.Brushes.Black;
            PinButton.Opacity = isPinned ? 1.0 : 0.7;
        }
    }
}