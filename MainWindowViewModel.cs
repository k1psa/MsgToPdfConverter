using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Collections.Generic;
using System.Linq;
using MsgToPdfConverter.Services;
using MsgToPdfConverter.Utils;

namespace MsgToPdfConverter
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        // State
        private ObservableCollection<string> _selectedFiles = new ObservableCollection<string>();
        private string _selectedOutputFolder;
        private bool _isConverting;
        private bool _cancellationRequested;
        private bool _deleteFilesAfterConversion;
        private bool _isPinned;
        private int _progressValue;
        private int _progressMax;
        private string _processingStatus;
        private string _fileCountText;
        private bool _appendAttachments;
        private bool _combineAllPdfs;
        private string _combinedPdfOutputPath;
        private string _lastConfirmedCombinedPdfPath = null;
        private bool _combinePdfOverwriteConfirmed = false;
        private bool _isProcessingFile;
        private int _fileProgressValue;
        private int _fileProgressMax = 100;

        // Services
        private readonly EmailConverterService _emailService = new EmailConverterService();
        private readonly AttachmentService _attachmentService;
        private readonly OutlookImportService _outlookImportService = new OutlookImportService();
        private readonly FileListService _fileListService = new FileListService();

        public MainWindowViewModel()
        {
            _attachmentService = new AttachmentService(
                (path, text, _) => PdfService.AddHeaderPdf(path, text),
                OfficeConversionService.TryConvertOfficeToPdf,
                PdfAppendTest.AppendPdfs,
                _emailService
            );
            SelectFilesCommand = new RelayCommand(SelectFiles);
            ClearListCommand = new RelayCommand(ClearList);
            SelectOutputFolderCommand = new RelayCommand(SelectOutputFolder);
            ClearOutputFolderCommand = new RelayCommand(ClearOutputFolder);
            ConvertCommand = new AsyncRelayCommand(ConvertAsync, (obj) => !_isConverting && _selectedFiles.Count > 0);
            CancelCommand = new RelayCommand(Cancel, (obj) => _isConverting);
            OptionsCommand = new RelayCommand(OpenOptions);
            PinCommand = new RelayCommand(TogglePin);
            RemoveSelectedFilesCommand = new RelayCommand(RemoveSelectedFiles);
            FileCountText = $"Files selected: {_selectedFiles.Count}";
            _selectedFiles.CollectionChanged += (s, e) =>
            {
                FileCountText = $"Files selected: {_selectedFiles.Count}";
                (ConvertCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            };
            Console.WriteLine("MainWindowViewModel initialized");
        }

        // Properties for binding
        public ObservableCollection<string> SelectedFiles { get => _selectedFiles; /* set removed to prevent replacement */ }
        public string SelectedOutputFolder { get => _selectedOutputFolder; set { _selectedOutputFolder = value; OnPropertyChanged(nameof(SelectedOutputFolder)); } }
        public bool IsConverting { get => _isConverting; set { _isConverting = value; OnPropertyChanged(nameof(IsConverting)); (ConvertCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged(); } }
        public bool CancellationRequested { get => _cancellationRequested; set { _cancellationRequested = value; OnPropertyChanged(nameof(CancellationRequested)); } }
        public bool DeleteFilesAfterConversion { get => _deleteFilesAfterConversion; set { _deleteFilesAfterConversion = value; OnPropertyChanged(nameof(DeleteFilesAfterConversion)); } }
        public bool IsPinned { get => _isPinned; set { _isPinned = value; OnPropertyChanged(nameof(IsPinned)); } }
        public int ProgressValue { get => _progressValue; set { _progressValue = value; OnPropertyChanged(nameof(ProgressValue)); } }
        public int ProgressMax { get => _progressMax; set { _progressMax = value; OnPropertyChanged(nameof(ProgressMax)); } }
        public string ProcessingStatus { get => _processingStatus; set { _processingStatus = value; OnPropertyChanged(nameof(ProcessingStatus)); } }
        public string FileCountText { get => _fileCountText; set { _fileCountText = value; OnPropertyChanged(nameof(FileCountText)); } }
        public bool AppendAttachments { get => _appendAttachments; set { _appendAttachments = value; OnPropertyChanged(nameof(AppendAttachments)); } }
        public bool CombineAllPdfs
        {
            get => _combineAllPdfs;
            set
            {
                if (_combineAllPdfs != value)
                {
                    _combineAllPdfs = value;
                    OnPropertyChanged(nameof(CombineAllPdfs));
                    if (_combineAllPdfs)
                    {
                        // Always show file save dialog when checked
                        string path = FileDialogHelper.SavePdfFileDialog("Binder1.pdf");
                        if (!string.IsNullOrEmpty(path))
                        {
                            CombinedPdfOutputPath = path;
                            _combinePdfOverwriteConfirmed = true;
                            _lastConfirmedCombinedPdfPath = path;
                        }
                        else
                        {
                            // If user cancels, uncheck
                            _combineAllPdfs = false;
                            OnPropertyChanged(nameof(CombineAllPdfs));
                            _combinePdfOverwriteConfirmed = false;
                            _lastConfirmedCombinedPdfPath = null;
                        }
                    }
                    else
                    {
                        CombinedPdfOutputPath = null;
                        _combinePdfOverwriteConfirmed = false;
                        _lastConfirmedCombinedPdfPath = null;
                    }
                }
            }
        }
        public string CombinedPdfOutputPath
        {
            get => _combinedPdfOutputPath;
            set { _combinedPdfOutputPath = value; OnPropertyChanged(nameof(CombinedPdfOutputPath)); }
        }
        public bool IsProcessingFile { get => _isProcessingFile; set { _isProcessingFile = value; OnPropertyChanged(nameof(IsProcessingFile)); } }
        public int FileProgressValue { get => _fileProgressValue; set { _fileProgressValue = value; OnPropertyChanged(nameof(FileProgressValue)); OnPropertyChanged(nameof(FileProgressPercentage)); OnPropertyChanged(nameof(FileProgressRatio)); } }
        public int FileProgressMax { get => _fileProgressMax; set { _fileProgressMax = value; OnPropertyChanged(nameof(FileProgressMax)); OnPropertyChanged(nameof(FileProgressPercentage)); OnPropertyChanged(nameof(FileProgressRatio)); } }
        public double FileProgressPercentage => _fileProgressMax > 0 ? (double)_fileProgressValue / _fileProgressMax * 100 : 0;
        public double FileProgressRatio => _fileProgressMax > 0 ? (double)_fileProgressValue / _fileProgressMax : 0;

        // Commands
        public ICommand SelectFilesCommand { get; }
        public ICommand ClearListCommand { get; }
        public ICommand SelectOutputFolderCommand { get; }
        public ICommand ClearOutputFolderCommand { get; }
        public ICommand ConvertCommand { get; }
        public ICommand CancelCommand { get; }
        public ICommand OptionsCommand { get; }
        public ICommand PinCommand { get; }
        public ICommand RemoveSelectedFilesCommand { get; }

        // Methods for commands
        private void SelectFiles(object parameter)
        {
            var mainWindow = Application.Current.MainWindow;
            bool wasTopmost = mainWindow != null && mainWindow.Topmost;
            if (IsPinned && mainWindow != null) mainWindow.Topmost = false;
            var newFiles = FileDialogHelper.OpenMsgFileDialog();
            if (IsPinned && mainWindow != null) mainWindow.Topmost = true;
            if (newFiles != null && newFiles.Count > 0)
            {
                var updated = _fileListService.AddFiles(new System.Collections.Generic.List<string>(_selectedFiles), newFiles);
                _selectedFiles.Clear();
                foreach (var file in updated)
                    _selectedFiles.Add(file);
            }
        }

        private void ClearList(object parameter)
        {
            _selectedFiles.Clear();
        }

        private void SelectOutputFolder(object parameter)
        {
            var mainWindow = Application.Current.MainWindow;
            bool wasTopmost = mainWindow != null && mainWindow.Topmost;
            if (IsPinned && mainWindow != null) mainWindow.Topmost = false;
            var folder = FileDialogHelper.OpenFolderDialog();
            if (IsPinned && mainWindow != null) mainWindow.Topmost = true;
            if (!string.IsNullOrEmpty(folder))
            {
                SelectedOutputFolder = folder;
            }
        }

        private void ClearOutputFolder(object parameter)
        {
            SelectedOutputFolder = null;
        }

        private async Task ConvertAsync(object parameter)
        {
            if (IsConverting) return;
            // Pre-check for combined output overwrite BEFORE starting conversion
            if (CombineAllPdfs)
            {
                var mainWindow = Application.Current?.MainWindow;
                bool wasTopmost = mainWindow != null && mainWindow.Topmost;
                // Only skip dialog if user just confirmed overwrite for the same file
                while (File.Exists(CombinedPdfOutputPath) && (!_combinePdfOverwriteConfirmed || CombinedPdfOutputPath != _lastConfirmedCombinedPdfPath))
                {
                    if (IsPinned && mainWindow != null) mainWindow.Topmost = false;
                    string path = FileDialogHelper.SavePdfFileDialog(Path.GetFileName(CombinedPdfOutputPath));
                    if (IsPinned && mainWindow != null) mainWindow.Topmost = wasTopmost;
                    if (string.IsNullOrEmpty(path))
                    {
                        // User cancelled
                        return;
                    }
                    CombinedPdfOutputPath = path;
                    _combinePdfOverwriteConfirmed = true;
                    _lastConfirmedCombinedPdfPath = path;
                }
            }
            Console.WriteLine($"Starting conversion for {SelectedFiles.Count} files. Output folder: {SelectedOutputFolder}");
            Console.WriteLine($"[DEBUG] Passing DeleteMsgAfterConversion: {DeleteFilesAfterConversion}");
            IsConverting = true;
            ProgressValue = 0;
            ProgressMax = SelectedFiles.Count;
            ProcessingStatus = "";
            var conversionService = new ConversionService();
            try
            {
                List<string> generatedPdfs = new List<string>();
                var result = await Task.Run(() =>
                {
                    var res = conversionService.ConvertFilesWithAttachments(
                        new System.Collections.Generic.List<string>(SelectedFiles),
                        SelectedOutputFolder,
                        AppendAttachments,
                        false, // always ignore extractOriginalOnly
                        DeleteFilesAfterConversion,
                        CombineAllPdfs, // <--- pass combineAllPdfs
                        _emailService,
                        _attachmentService,
                        (processed, total, progress, statusText) =>
                        {
                            ProcessingStatus = statusText;
                            ProgressValue = processed;
                            IsProcessingFile = !string.IsNullOrEmpty(statusText) && statusText.Contains("Processing file");
                            Console.WriteLine($"Progress: {processed}/{total} - {statusText}");
                        },
                        (current, max) =>
                        {
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                FileProgressValue = current;
                                FileProgressMax = max;
                                IsProcessingFile = max > 0;
                                Console.WriteLine($"[FILE-PROGRESS] {current}/{max} = {FileProgressRatio:F2} - IsProcessingFile: {IsProcessingFile}");
                            });
                        },
                        () => CancellationRequested,
                        null, // no messagebox during conversion
                        generatedPdfs // <-- pass list to collect generated PDFs
                    );
                    return res;
                });
                string statusMessage;
                if (CombineAllPdfs && !string.IsNullOrEmpty(CombinedPdfOutputPath) && generatedPdfs.Count > 0)
                {
                    PdfAppendTest.AppendPdfs(generatedPdfs, CombinedPdfOutputPath);
                    // Always delete intermediate PDFs after combining
                    foreach (var pdf in generatedPdfs)
                    {
                        bool isOriginal = SelectedFiles.Any(f => string.Equals(f, pdf, StringComparison.OrdinalIgnoreCase));
                        if (!isOriginal)
                        {
                            try { if (File.Exists(pdf)) File.Delete(pdf); } catch { }
                        }
                    }
                    // Optionally delete original source files if requested
                    if (DeleteFilesAfterConversion)
                    {
                        foreach (var src in SelectedFiles)
                        {
                            if (!string.Equals(src, CombinedPdfOutputPath, StringComparison.OrdinalIgnoreCase))
                            {
                                try { if (File.Exists(src)) FileService.MoveFileToRecycleBin(src); } catch { }
                            }
                        }
                    }
                    statusMessage = $"{generatedPdfs.Count} file(s) have been combined into {System.IO.Path.GetFileName(CombinedPdfOutputPath)}";
                }
                else
                {
                    statusMessage = result.Cancelled
                        ? $"Processing cancelled. Processed {result.Processed} files. Success: {result.Success}, Failed: {result.Fail}"
                        : $"Processing completed. Total files: {SelectedFiles.Count}, Success: {result.Success}, Failed: {result.Fail}";
                }
                Console.WriteLine(statusMessage);
                // Ensure MessageBox is centered and on top, even if window is pinned
                var mainWindow2 = Application.Current?.MainWindow;
                bool wasTopmost2 = mainWindow2 != null && mainWindow2.Topmost;
                if (IsPinned && mainWindow2 != null) mainWindow2.Topmost = false;
                // Use helper to center MessageBox within main window
                MsgToPdfConverter.Utils.MessageBoxHelper.ShowCentered(mainWindow2, statusMessage, "Processing Results", MessageBoxButton.OK,
                    (result.Fail > 0 && !CombineAllPdfs) ? MessageBoxImage.Warning : MessageBoxImage.Information);
                if (IsPinned && mainWindow2 != null) mainWindow2.Topmost = wasTopmost2;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                Console.WriteLine("Conversion finished.");
                IsConverting = false;
                CancellationRequested = false;
                ProcessingStatus = "";
                
                // Delay resetting file progress to let user see 100% completion
#pragma warning disable CS4014 // Because this call is not awaited, execution continues before call is completed
                Task.Delay(1500).ContinueWith(_ =>
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        IsProcessingFile = false;
                        FileProgressValue = 0;
                        FileProgressMax = 0;
                    });
                });
#pragma warning restore CS4014
                
                // After conversion, reset confirmation so dialog will show again if needed next time
                _combinePdfOverwriteConfirmed = false;
                _lastConfirmedCombinedPdfPath = CombinedPdfOutputPath;
            }
        }

        private void Cancel(object parameter)
        {
            CancellationRequested = true;
            ProcessingStatus = "Cancelling... Please wait.";
        }

        private void OpenOptions(object parameter)
        {
            var optionsWindow = new OptionsWindow(DeleteFilesAfterConversion, Properties.Settings.Default.CloseButtonBehavior ?? "Ask")
            {
                Owner = Application.Current.MainWindow
            };
            if (optionsWindow.ShowDialog() == true)
            {
                DeleteFilesAfterConversion = optionsWindow.DeleteFilesAfterConversion;
                Console.WriteLine($"[DEBUG] DeleteFilesAfterConversion set to: {DeleteFilesAfterConversion}");
            }
        }

        private void TogglePin(object parameter)
        {
            IsPinned = !IsPinned;
            if (Application.Current.MainWindow != null)
                Application.Current.MainWindow.Topmost = IsPinned;
        }

        private void RemoveSelectedFiles(object selectedItems)
        {
            if (selectedItems is System.Collections.IList items && items.Count > 0)
            {
                var toRemove = new System.Collections.Generic.List<string>();
                foreach (var item in items)
                {
                    if (item is string s)
                        toRemove.Add(s);
                }
                var updated = _fileListService.RemoveFiles(new System.Collections.Generic.List<string>(_selectedFiles), toRemove);
                _selectedFiles.Clear();
                foreach (var file in updated)
                    _selectedFiles.Add(file);
            }
        }

        // Move a file in the SelectedFiles collection from oldIndex to newIndex
        public void MoveFile(int oldIndex, int newIndex)
        {
            if (oldIndex < 0 || newIndex < 0 || oldIndex == newIndex || oldIndex >= _selectedFiles.Count || newIndex >= _selectedFiles.Count)
                return;
            var item = _selectedFiles[oldIndex];
            _selectedFiles.RemoveAt(oldIndex);
            _selectedFiles.Insert(newIndex, item);
        }

        // Drag-and-drop support for ListBox
        public void HandleDrop(IDataObject data)
        {
            Console.WriteLine("[DEBUG] HandleDrop called!");
            Console.WriteLine($"[DEBUG] Available data formats: {string.Join(", ", data.GetFormats())}");
            
            // 1. Standard file/folder drop
            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] droppedItems = (string[])data.GetData(DataFormats.FileDrop);
                var updated = new System.Collections.Generic.List<string>(_selectedFiles);
                foreach (string item in droppedItems)
                {
                    if (File.Exists(item))
                    {
                        updated = _fileListService.AddFiles(updated, new[] { item });
                    }
                    else if (Directory.Exists(item))
                    {
                        updated = _fileListService.AddFilesFromDirectory(updated, item);
                    }
                }
                _selectedFiles.Clear();
                foreach (var file in updated)
                    _selectedFiles.Add(file);
                return;
            }

            // 2. Outlook drag-and-drop: distinguish between attachment and email
            if (data.GetDataPresent("FileGroupDescriptorW") || data.GetDataPresent("FileGroupDescriptor"))
            {
                Console.WriteLine("[DEBUG] Outlook drag-and-drop detected!");
                try
                {
                    // Try to get the filename from the drop (for attachments)
                    string[] formats = data.GetFormats();
                    bool isAttachment = false;
                    string attachmentName = null;
                    if (data.GetDataPresent("FileGroupDescriptorW"))
                    {
                        // Try to extract the filename from the FileGroupDescriptorW stream
                        var stream = (System.IO.MemoryStream)data.GetData("FileGroupDescriptorW");
                        if (stream != null)
                        {
                            byte[] fileGroupDescriptor = new byte[stream.Length];
                            stream.Read(fileGroupDescriptor, 0, fileGroupDescriptor.Length);
                            // The filename is a Unicode string starting at offset 76
                            int nameStart = 76;
                            int nameLength = fileGroupDescriptor.Length - nameStart;
                            string name = System.Text.Encoding.Unicode.GetString(fileGroupDescriptor, nameStart, nameLength);
                            int nullIndex = name.IndexOf('\0');
                            if (nullIndex > 0)
                                name = name.Substring(0, nullIndex);
                            attachmentName = name;
                            // If the filename is not .msg, treat as attachment
                            if (!string.IsNullOrEmpty(name) && !name.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                                isAttachment = true;
                        }
                    }
                    if (isAttachment && !string.IsNullOrEmpty(attachmentName))
                    {
                        Console.WriteLine($"[DEBUG] Detected Outlook attachment drop: {attachmentName}");
                        string outputFolder = !string.IsNullOrEmpty(SelectedOutputFolder) ? SelectedOutputFolder : null;
                        var mainWindow = Application.Current.MainWindow;
                        if (string.IsNullOrEmpty(outputFolder))
                        {
                            if (IsPinned && mainWindow != null) mainWindow.Topmost = false;
                            outputFolder = FileDialogHelper.OpenFolderDialog();
                            if (IsPinned && mainWindow != null) mainWindow.Topmost = true;
                        }
                        if (string.IsNullOrEmpty(outputFolder))
                            return;
                        // Save the attachment file
                        var result = _outlookImportService.ExtractAttachmentsFromDragDrop(data, outputFolder, FileService.SanitizeFileName);
                        var updated = _fileListService.AddFiles(new System.Collections.Generic.List<string>(_selectedFiles), result.ExtractedFiles);
                        _selectedFiles.Clear();
                        foreach (var file in updated)
                            _selectedFiles.Add(file);
                        if (result.ExtractedFiles.Count > 0)
                        {
                            Console.WriteLine($"[DEBUG] Successfully added {result.ExtractedFiles.Count} attachment(s) to the list:");
                            foreach (var file in result.ExtractedFiles)
                                Console.WriteLine($"[DEBUG] - {Path.GetFileName(file)}");
                        }
                        if (result.SkippedFiles.Count > 0)
                        {
                            Console.WriteLine($"[DEBUG] Skipped files: {string.Join(", ", result.SkippedFiles)}");
                        }
                        return;
                    }
                    // Otherwise, treat as email
                    Console.WriteLine("[DEBUG] Detected Outlook email drop (saving as .msg)");
                    string outputFolderEmail = !string.IsNullOrEmpty(SelectedOutputFolder) ? SelectedOutputFolder : null;
                    var mainWindowEmail = Application.Current.MainWindow;
                    if (string.IsNullOrEmpty(outputFolderEmail))
                    {
                        if (IsPinned && mainWindowEmail != null) mainWindowEmail.Topmost = false;
                        outputFolderEmail = FileDialogHelper.OpenFolderDialog();
                        if (IsPinned && mainWindowEmail != null) mainWindowEmail.Topmost = true;
                    }
                    if (string.IsNullOrEmpty(outputFolderEmail))
                        return;
                    var resultEmail = _outlookImportService.ExtractMsgFilesFromDragDrop(
                        data,
                        outputFolderEmail,
                        FileService.SanitizeFileName);
                    var updatedEmail = _fileListService.AddFiles(new System.Collections.Generic.List<string>(_selectedFiles), resultEmail.ExtractedFiles);
                    _selectedFiles.Clear();
                    foreach (var file in updatedEmail)
                        _selectedFiles.Add(file);
                    if (resultEmail.ExtractedFiles.Count > 0)
                    {
                        Console.WriteLine($"[DEBUG] Successfully added {resultEmail.ExtractedFiles.Count} email(s) to the list:");
                        foreach (var file in resultEmail.ExtractedFiles)
                            Console.WriteLine($"[DEBUG] - {Path.GetFileName(file)}");
                    }
                    if (resultEmail.SkippedFiles.Count > 0)
                    {
                        Console.WriteLine($"[DEBUG] Skipped files: {string.Join(", ", resultEmail.SkippedFiles)}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[DEBUG] Error processing Outlook drop: {ex.Message}");
                }
                return;
            }

            // 3. If all else fails, inform the user
            Console.WriteLine("[DEBUG] Could not extract email or attachment from Outlook drag-and-drop.");
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
