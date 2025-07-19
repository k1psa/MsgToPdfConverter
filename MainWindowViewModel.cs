using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Threading;
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
        private DispatcherTimer _progressRingDelayTimer;
        private bool _pendingShowProgressRing;

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
            #if DEBUG
                DebugLogger.Log("MainWindowViewModel initialized");
            #endif
            _progressRingDelayTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
            _progressRingDelayTimer.Tick += (s, e) =>
            {
                _progressRingDelayTimer.Stop();
                if (_pendingShowProgressRing)
                {
                    IsProcessingFile = true;
                }
            };
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
            #if DEBUG
            DebugLogger.Log($"Starting conversion for {SelectedFiles.Count} files. Output folder: {SelectedOutputFolder}");
            DebugLogger.Log($"[DEBUG] Passing DeleteFilesAfterConversion: {DeleteFilesAfterConversion}");
            if (SelectedFiles == null)
                DebugLogger.Log("[DEBUG] SelectedFiles is null!");
            else
                DebugLogger.Log($"[DEBUG] SelectedFiles: [{string.Join(", ", SelectedFiles)}]");
            #endif
            IsConverting = true;
            ProgressValue = 0;
            ProgressMax = SelectedFiles.Count;
            ProcessingStatus = "";
            #if DEBUG
            if (_emailService == null) DebugLogger.Log("[DEBUG] _emailService is null!");
            if (_attachmentService == null) DebugLogger.Log("[DEBUG] _attachmentService is null!");
            if (_outlookImportService == null) DebugLogger.Log("[DEBUG] _outlookImportService is null!");
            if (_fileListService == null) DebugLogger.Log("[DEBUG] _fileListService is null!");
            #endif
            var conversionService = new ConversionService();
            try
            {
                List<string> generatedPdfs = new List<string>();
                var (success, fail, processed, cancelled) = await Task.Run(() =>
                {
                    if (conversionService == null)
                    {
                        #if DEBUG
                        DebugLogger.Log("[DEBUG] conversionService is null!");
                        #endif
                        // Return default tuple
                        return (0, 0, 0, false);
                    }
                    var res = conversionService.ConvertFilesWithAttachments(
                        new System.Collections.Generic.List<string>(SelectedFiles),
                        SelectedOutputFolder,
                        AppendAttachments,
                        false, // always ignore extractOriginalOnly
                        DeleteFilesAfterConversion,
                        CombineAllPdfs, // <--- pass combineAllPdfs
                        _emailService,
                        _attachmentService,
                        (fileProcessed, total, progress, statusText) =>
                        {
                            ProcessingStatus = statusText;
                            ProgressValue = fileProcessed;
                            // Remove direct IsProcessingFile set here
                            #if DEBUG
                            DebugLogger.Log($"Progress: {fileProcessed}/{total} - {statusText}");
                            #endif
                        },
                        (current, max) =>
                        {
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                FileProgressValue = current;
                                FileProgressMax = max;
                                // Progress ring delay logic
                                if (max > 0 && current < max)
                                {
                                    if (!_progressRingDelayTimer.IsEnabled && !IsProcessingFile)
                                    {
                                        _pendingShowProgressRing = true;
                                        _progressRingDelayTimer.Start();
                                    }
                                }
                                else
                                {
                                    _pendingShowProgressRing = false;
                                    _progressRingDelayTimer.Stop();
                                    IsProcessingFile = false;
                                }
                                #if DEBUG
                                DebugLogger.Log($"[FILE-PROGRESS] {current}/{max} = {FileProgressRatio:F2} - IsProcessingFile: {IsProcessingFile}");
                                #endif
                            });
                        },
                        () => CancellationRequested,
                        null, // no messagebox during conversion
                        generatedPdfs // <-- pass list to collect generated PDFs
                    );
                    // Value tuple cannot be null, but check for default values
                    if (res.Item1 == 0 && res.Item2 == 0 && res.Item3 == 0 && !res.Item4)
                    {
                        #if DEBUG
                        DebugLogger.Log("[DEBUG] ConvertFilesWithAttachments returned default result! Possible error in conversion.");
                        #endif
                    }
                    return res;
                });
                string statusMessage;
                if (CombineAllPdfs && !string.IsNullOrEmpty(CombinedPdfOutputPath) && generatedPdfs.Count > 0)
                {
                    #if DEBUG
                    DebugLogger.Log($"[DEBUG] Calling PdfAppendTest.AppendPdfs. generatedPdfs.Count={generatedPdfs.Count}, CombinedPdfOutputPath={CombinedPdfOutputPath}");
                    #endif
                    try
                    {
                        #if DEBUG
                        if (generatedPdfs == null)
                        {
                            DebugLogger.Log("[DEBUG] generatedPdfs is null!");
                        }
                        else if (generatedPdfs.Count == 0)
                        {
                            DebugLogger.Log("[DEBUG] generatedPdfs is empty!");
                        }
                        if (string.IsNullOrEmpty(CombinedPdfOutputPath))
                        {
                            DebugLogger.Log("[DEBUG] CombinedPdfOutputPath is null or empty!");
                        }
                        #endif
                        PdfAppendTest.AppendPdfs(generatedPdfs, CombinedPdfOutputPath);
                    }
                    catch (Exception ex)
                    {
                        #if DEBUG
                        DebugLogger.Log($"[ERROR] Exception in PdfAppendTest.AppendPdfs: {ex.Message}");
                        #endif
                        throw;
                    }
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
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] Entering file deletion loop. DeleteFilesAfterConversion={DeleteFilesAfterConversion}, AppendAttachments={AppendAttachments}");
                        #endif
                        if (!DeleteFilesAfterConversion)
                        {
                            #if DEBUG
                            DebugLogger.Log("[DELETE] Skipping deletion because DeleteFilesAfterConversion is false.");
                            #endif
                        }
                        else if (SelectedFiles == null)
                        {
                            #if DEBUG
                            DebugLogger.Log("[DELETE] SelectedFiles is null!");
                            #endif
                        }
                        else
                        {
                            foreach (var src in SelectedFiles)
                            {
                                if (src == null)
                                {
                                    #if DEBUG
                                    DebugLogger.Log("[DELETE] Skipped null file reference in SelectedFiles.");
                                    #endif
                                    continue;
                                }
                                if (!string.IsNullOrEmpty(src) && !string.Equals(src, CombinedPdfOutputPath, StringComparison.OrdinalIgnoreCase))
                                {
                                    try
                                    {
                                        if (File.Exists(src))
                                        {
                                            #if DEBUG
                                            DebugLogger.Log($"[DELETE] Deleting file: {src}");
                                            #endif
                                            FileService.MoveFileToRecycleBin(src);
                                        }
                                        else
                                        {
                                            #if DEBUG
                                            DebugLogger.Log($"[DELETE] File not found or already deleted: {src}");
                                            #endif
                                        }
                                    }
                                    catch (NullReferenceException nre)
                                    {
                                        #if DEBUG
                                        DebugLogger.Log($"[DELETE] NullReferenceException deleting file '{src}': {nre.Message}");
                                        if (src == null) DebugLogger.Log("[DELETE] src is null");
                                        #endif
                                    }
                                    catch (Exception ex)
                                    {
                                        #if DEBUG
                                        DebugLogger.Log($"[DELETE] Exception deleting file '{src}': {ex.Message}");
                                        #endif
                                    }
                                }
                                else
                                {
                                    #if DEBUG
                                    DebugLogger.Log($"[DELETE] Skipped empty or output file reference: {src}");
                                    #endif
                                }
                            }
                        }
                    }
                    statusMessage = $"{generatedPdfs.Count} file(s) have been combined into {System.IO.Path.GetFileName(CombinedPdfOutputPath)}";
                }
                else
                {
                    statusMessage = cancelled
                        ? $"Processing cancelled. Processed {processed} files. Success: {success}, Failed: {fail}"
                        : $"Processing completed. Total files: {SelectedFiles.Count}, Success: {success}, Failed: {fail}";
                }
                #if DEBUG
                DebugLogger.Log(statusMessage);
                DebugLogger.Log("[DEBUG] About to show completion dialog.");
                #endif
                // Ensure MessageBox is centered and on top, even if window is pinned
                var mainWindow2 = Application.Current?.MainWindow;
                bool wasTopmost2 = mainWindow2 != null && mainWindow2.Topmost;
                // If minimized, restore before showing dialog
                if (mainWindow2 != null)
                {
                    if (mainWindow2.WindowState == WindowState.Minimized)
                    {
                       
                        mainWindow2.WindowState = WindowState.Normal;
                        mainWindow2.Show();
                        mainWindow2.Activate();
                    }
                    // If window is hidden (tray), show and activate
                    if (!mainWindow2.IsVisible)
                    {
                        #if DEBUG
                        DebugLogger.Log("[DEBUG] Main window is not visible (tray), showing before dialog.");
                        #endif
                        mainWindow2.Show();
                        mainWindow2.WindowState = WindowState.Normal;
                        mainWindow2.Activate();
                    }
                }
                if (IsPinned && mainWindow2 != null) mainWindow2.Topmost = false;
                // Use helper to center MessageBox within main window
                MsgToPdfConverter.Utils.MessageBoxHelper.ShowCentered(mainWindow2, statusMessage, "Processing Results", MessageBoxButton.OK,
                    (fail > 0 && !CombineAllPdfs) ? MessageBoxImage.Warning : MessageBoxImage.Information);
                #if DEBUG
                DebugLogger.Log("[DEBUG] Completion dialog closed.");
                #endif
                if (IsPinned && mainWindow2 != null) mainWindow2.Topmost = wasTopmost2;
            }
            catch (Exception ex)
            {
#if DEBUG
                DebugLogger.Log($"An error occurred: {ex.Message}");
#endif
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
#if DEBUG
                DebugLogger.Log("Conversion finished.");
#endif
                IsConverting = false;
                CancellationRequested = false;
                ProcessingStatus = "";
                
                // Delay resetting file progress to let user see 100% completion
#pragma warning disable CS4014 // Because this call is not awaited, execution continues before call is completed
                Task.Delay(2000).ContinueWith(_ =>
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        _pendingShowProgressRing = false;
                        _progressRingDelayTimer.Stop();
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
                #if DEBUG
                DebugLogger.Log($"[DEBUG] DeleteFilesAfterConversion set to: {DeleteFilesAfterConversion}");
                #endif
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
            #if DEBUG
            DebugLogger.Log("[DEBUG] HandleDrop called!");
            DebugLogger.Log($"[DEBUG] Available data formats: {string.Join(", ", data.GetFormats())}");
            #endif
            
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
                #if DEBUG
                DebugLogger.Log("[DEBUG] Outlook drag-and-drop detected!");
                #endif
                try
                {
                    // Check if this is a child MSG attachment by looking at the data formats
                    string[] formats = data.GetFormats();
                    #if DEBUG
                    DebugLogger.Log($"[DEBUG] Data formats: {string.Join(", ", formats)}");
                    #endif
                    
                    // Child MSG attachments have ZoneIdentifier but no RenPrivateMessages/RenPrivateLatestMessages
                    // Parent emails have RenPrivateMessages and RenPrivateLatestMessages
                    bool isChildMsgAttachment = data.GetDataPresent("ZoneIdentifier") && 
                                               !data.GetDataPresent("RenPrivateMessages") && 
                                               !data.GetDataPresent("RenPrivateLatestMessages");
                    
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
                        }
                    }
                    
                    #if DEBUG
                    DebugLogger.Log($"[DEBUG] Attachment name: {attachmentName}");
                    DebugLogger.Log($"[DEBUG] Is child MSG attachment: {isChildMsgAttachment}");
                    #endif
                    
                    // If it's a child MSG attachment, treat it as an email extraction
                    if (isChildMsgAttachment && !string.IsNullOrEmpty(attachmentName) && 
                        attachmentName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                    {
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] Detected child MSG attachment drop: {attachmentName} (treating as email)");
                        #endif
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
                        
                        // Extract child MSG attachments and warn if ambiguous
                        var resultChildMsg = _outlookImportService.ExtractChildMsgFromDragDrop(data, outputFolder, FileService.SanitizeFileName, attachmentName);
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] Child MSG extraction result: Extracted={resultChildMsg.ExtractedFiles.Count}, Skipped={resultChildMsg.SkippedFiles.Count}");
                        #endif
                        if (resultChildMsg.ExtractedFiles.Count > 1)
                        {
                            // Show warning to user about ambiguity
                            var msgBoxWindow = Application.Current?.MainWindow;
                            bool wasTopmost = msgBoxWindow != null && msgBoxWindow.Topmost;
                            if (IsPinned && msgBoxWindow != null) msgBoxWindow.Topmost = false;
                            MsgToPdfConverter.Utils.MessageBoxHelper.ShowCentered(msgBoxWindow,
                                $"Warning: Multiple MSG attachments with the name '{attachmentName}' were found. All of them were extracted.",
                                "Ambiguous MSG Extraction", MessageBoxButton.OK, MessageBoxImage.Warning);
                            if (IsPinned && msgBoxWindow != null) msgBoxWindow.Topmost = wasTopmost;
                        }
                        if (resultChildMsg.ExtractedFiles.Count > 0)
                        {
                            foreach (var extractedFile in resultChildMsg.ExtractedFiles)
                            {
                            #if DEBUG
                            DebugLogger.Log($"[DEBUG] Extracted file: {extractedFile}");
                            #endif
                                if (!string.IsNullOrEmpty(extractedFile) && File.Exists(extractedFile))
                                {
                                    if (!_selectedFiles.Contains(extractedFile))
                                    {
                                        _selectedFiles.Add(extractedFile);
                                        #if DEBUG
                                        DebugLogger.Log($"[DEBUG] Added child MSG to list: {Path.GetFileName(extractedFile)}");
                                        #endif
                                    }
                                    else
                                    {
                                        #if DEBUG
                                        DebugLogger.Log($"[DEBUG] Child MSG already in list: {Path.GetFileName(extractedFile)}");
                                        #endif
                                    }
                                }
                                else
                                {
                                    #if DEBUG
                                    DebugLogger.Log($"[DEBUG] Skipped non-existent or empty extracted file: {extractedFile}");
                                    #endif
                                }
                            }
                        }
                        else
                        {
                            #if DEBUG
                            DebugLogger.Log("[DEBUG] No files were extracted from child MSG");
                            #endif
                        }
                        if (resultChildMsg.SkippedFiles.Count > 0)
                        {
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] Skipped files: {string.Join(", ", resultChildMsg.SkippedFiles)}");
                        #endif
                        }
                        return;
                    }
                    // Check if it's a non-MSG attachment
                    else if (!string.IsNullOrEmpty(attachmentName) && !attachmentName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase))
                    {
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] Detected Outlook attachment drop: {attachmentName}");
                        #endif
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
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] Successfully added {result.ExtractedFiles.Count} attachment(s) to the list:");
                        #endif
                            foreach (var file in result.ExtractedFiles)
                            {
                                #if DEBUG
                                DebugLogger.Log($"[DEBUG] - {Path.GetFileName(file)}");
                                #endif
                            }
                        }
                        if (result.SkippedFiles.Count > 0)
                        {
                        #if DEBUG
                        DebugLogger.Log($"[DEBUG] Skipped files: {string.Join(", ", result.SkippedFiles)}");
                        #endif
                        }
                        return;
                    }
                    // Otherwise, treat as parent email
                    #if DEBUG
                    DebugLogger.Log("[DEBUG] Detected Outlook parent email drop (saving as .msg)");
                    #endif
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

                }
                catch (Exception ex)
                {
#if DEBUG
                    DebugLogger.Log($"[DEBUG] Error processing Outlook drop: {ex.Message}");
#endif
                }
                return;
            }

            // 3. If all else fails, inform the user
            #if DEBUG
            DebugLogger.Log("[DEBUG] Could not extract email or attachment from Outlook drag-and-drop.");
            #endif
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        // Add this method to the MainWindowViewModel class:
        public void SafeFileProgressTick()
        {
            if (FileProgressValue < FileProgressMax)
                FileProgressValue++;
            else
                FileProgressValue = FileProgressMax;
        }
    }
}
