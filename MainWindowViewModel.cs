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
                        // Show file save dialog when checked
                        string path = FileDialogHelper.SavePdfFileDialog("Binder1.pdf");
                        if (!string.IsNullOrEmpty(path))
                        {
                            CombinedPdfOutputPath = path;
                        }
                        else
                        {
                            // If user cancels, uncheck
                            _combineAllPdfs = false;
                            OnPropertyChanged(nameof(CombineAllPdfs));
                        }
                    }
                    else
                    {
                        CombinedPdfOutputPath = null;
                    }
                }
            }
        }
        public string CombinedPdfOutputPath
        {
            get => _combinedPdfOutputPath;
            set { _combinedPdfOutputPath = value; OnPropertyChanged(nameof(CombinedPdfOutputPath)); }
        }

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
                            Console.WriteLine($"Progress: {processed}/{total} - {statusText}");
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
                MessageBox.Show(statusMessage, "Processing Results", MessageBoxButton.OK,
                    (result.Fail > 0 && !CombineAllPdfs) ? MessageBoxImage.Warning : MessageBoxImage.Information);
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

            // 2. Outlook email drag-and-drop support
            if (data.GetDataPresent("FileGroupDescriptorW") || data.GetDataPresent("FileGroupDescriptor"))
            {
                Console.WriteLine("[DEBUG] Outlook email drop detected!");
                try
                {
                    string outputFolder = !string.IsNullOrEmpty(SelectedOutputFolder)
                        ? SelectedOutputFolder
                        : null;
                    var mainWindow = Application.Current.MainWindow;
                    if (string.IsNullOrEmpty(outputFolder))
                    {
                        if (IsPinned && mainWindow != null) mainWindow.Topmost = false;
                        outputFolder = FileDialogHelper.OpenFolderDialog();
                        if (IsPinned && mainWindow != null) mainWindow.Topmost = true;
                    }
                    if (string.IsNullOrEmpty(outputFolder))
                        return;

                    var result = _outlookImportService.ExtractMsgFilesFromDragDrop(
                        data,
                        outputFolder,
                        FileService.SanitizeFileName);

                    var updated = _fileListService.AddFiles(new System.Collections.Generic.List<string>(_selectedFiles), result.ExtractedFiles);
                    _selectedFiles.Clear();
                    foreach (var file in updated)
                        _selectedFiles.Add(file);

                    // Log success if files were extracted
                    if (result.ExtractedFiles.Count > 0)
                    {
                        Console.WriteLine($"[DEBUG] Successfully added {result.ExtractedFiles.Count} email(s) to the list:");
                        foreach (var file in result.ExtractedFiles)
                        {
                            Console.WriteLine($"[DEBUG] - {Path.GetFileName(file)}");
                        }
                    }

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

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
