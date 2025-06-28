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
            bool wasTopmost = this.Topmost;
            this.Topmost = false;
            var folder = FileDialogHelper.OpenFolderDialog();
            this.Topmost = wasTopmost;
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

        private async void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if (isConverting)
                return;
            isConverting = true;
            SetProcessingState(true);
            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = selectedFiles.Count;
            ProgressBar.Value = 0;
            var conversionService = new ConversionService();
            try
            {
                // Use the field instead of the removed checkbox
                bool appendAttachments = AppendAttachmentsCheckBox.IsChecked == true;
                bool extractOriginalOnlyLocal = extractOriginalOnly;
                var result = await Task.Run(() =>
                    conversionService.ConvertMsgFilesWithAttachments(
                        selectedFiles,
                        selectedOutputFolder,
                        appendAttachments,
                        extractOriginalOnlyLocal,
                        deleteMsgAfterConversion,
                        _emailService,
                        _attachmentService,
                        (processed, total, progress, statusText) =>
                        {
                            Dispatcher.Invoke(() =>
                            {
                                ProcessingStatusLabel.Foreground = System.Windows.Media.Brushes.Blue;
                                ProcessingStatusLabel.Text = statusText;
                                ProgressBar.Value = processed;
                            });
                        },
                        () => cancellationRequested,
                        (msg) => Dispatcher.Invoke(() => MessageBox.Show(msg, "Processing Results", MessageBoxButton.OK, MessageBoxImage.Information))
                    )
                );
                string statusMessage = result.Cancelled
                    ? $"Processing cancelled. Processed {result.Processed} files. Success: {result.Success}, Failed: {result.Fail}"
                    : $"Processing completed. Total files: {selectedFiles.Count}, Success: {result.Success}, Failed: {result.Fail}";
                MessageBox.Show(statusMessage, "Processing Results", MessageBoxButton.OK,
                    result.Fail > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetProcessingState(false);
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
                        : null;
                    if (string.IsNullOrEmpty(outputFolder))
                    {
                        bool wasTopmost = this.Topmost;
                        this.Topmost = false;
                        outputFolder = FileDialogHelper.OpenFolderDialog();
                        this.Topmost = wasTopmost;
                    }
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