using System;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace MsgToPdfConverter
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Command-line worker mode: --html2pdf <htmlPath> <pdfPath>
            if (e.Args != null && e.Args.Length == 3 && e.Args[0] == "--html2pdf")
            {
                string htmlPath = e.Args[1];
                string pdfPath = e.Args[2];
                string html = File.ReadAllText(htmlPath);
                int result = MsgToPdfConverter.HtmlToPdfWorker.Convert(html, pdfPath);
                Environment.Exit(result);
                return;
            }

            // Validation test mode: test page ordering logic
            if (e.Args != null && e.Args.Length >= 1 && e.Args[0] == "--validate")
            {
                ValidationTest.TestPageOrderingLogic();
                Environment.Exit(0);
                return;
            }
            // Test mode: test embedded extraction
            else if (e.Args != null && e.Args.Length >= 1 && e.Args[0] == "--test")
            {
                string testFile = e.Args.Length > 1 ? e.Args[1] : "a.docx";
                TestExtraction.TestEmbeddedExtraction(testFile);
                Environment.Exit(0);
                return;
            }

            // Check .NET Framework before starting WPF
            if (!Program.IsDotNetFrameworkInstalled())
            {
                System.Windows.MessageBox.Show(
                    ".NET Framework 4.8 or higher is required to run this application.\n\n" +
                    "Please install it from:\n" +
                    "https://dotnet.microsoft.com/download/dotnet-framework",
                    ".NET Framework Required",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                Environment.Exit(1);
                return;
            }
            // Check for Microsoft Office (Word and Excel) before starting WPF
            if (!Program.IsOfficeInstalled())
            {
                System.Windows.MessageBox.Show(
                    "Microsoft Office (Word and Excel) is required to convert Office documents.\n\n" +
                    "Please install Microsoft Office and try again.",
                    "Microsoft Office Not Found",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                Environment.Exit(1);
                return;
            }

            // Single-instance logic
            var singleInstance = new Utils.SingleInstanceManager();
            if (!singleInstance.IsFirstInstance)
            {
                // If not first instance and file/folder argument, send to running instance
                if (e.Args != null && e.Args.Length == 1)
                {
                    var arg = e.Args[0];
                    if (!string.IsNullOrWhiteSpace(arg))
                    {
                        try { singleInstance.SendFileToFirstInstance(arg); } catch { }
                    }
                }
                Environment.Exit(0);
                return;
            }

            // Diagnostic logging for single-instance issues
            // (Removed diagnostic logging)

            var pendingFiles = new System.Collections.Concurrent.ConcurrentQueue<string>();
            // Buffer initial file/folder argument if present (for first launch via context menu)
            if (e.Args != null && e.Args.Length == 1 && !string.IsNullOrWhiteSpace(e.Args[0]))
            {
                pendingFiles.Enqueue(e.Args[0]);
            }
            singleInstance.FileReceived += (file) =>
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    var mw = System.Windows.Application.Current.MainWindow as MainWindow;
                    if (mw != null)
                    {
                        var vm = mw.DataContext as MainWindowViewModel;
                        if (vm != null && !string.IsNullOrEmpty(file))
                        {
                            // Show and activate window if minimized or hidden
                            if (!mw.IsVisible)
                            {
                                mw.Show();
                                mw.WindowState = System.Windows.WindowState.Normal;
                            }
                            else if (mw.WindowState == System.Windows.WindowState.Minimized)
                            {
                                mw.WindowState = System.Windows.WindowState.Normal;
                                mw.Show();
                            }
                            mw.Activate();

                            var fileListService = new MsgToPdfConverter.Services.FileListService();
                            var updated = new System.Collections.Generic.List<string>(vm.SelectedFiles);
                            if (System.IO.Directory.Exists(file))
                            {
                                // Add all supported files from the folder (recursively)
                                updated = fileListService.AddFilesFromDirectory(updated, file);
                            }
                            else if (System.IO.File.Exists(file))
                            {
                                updated = fileListService.AddFiles(updated, new[] { file });
                            }
                            foreach (var f in updated)
                            {
                                if (!vm.SelectedFiles.Contains(f))
                                    vm.SelectedFiles.Add(f);
                            }
                        }
                        // Drain any pending files
                        while (pendingFiles.TryDequeue(out var pf))
                        {
                            if (!vm.SelectedFiles.Contains(pf))
                                vm.SelectedFiles.Add(pf);
                        }
                    }
                    else
                    {
                        // Buffer until main window is ready
                        if (!string.IsNullOrEmpty(file))
                            pendingFiles.Enqueue(file);
                    }
                });
            };
            this.Startup += (s, evt) =>
            {
                // On startup, try to drain any pending files, expanding folders to files
                System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    var mw = System.Windows.Application.Current.MainWindow as MainWindow;
                    if (mw != null)
                    {
                        var vm = mw.DataContext as MainWindowViewModel;
                        if (vm != null)
                        {
                            var fileListService = new MsgToPdfConverter.Services.FileListService();
                            var updated = new System.Collections.Generic.List<string>(vm.SelectedFiles);
                            while (pendingFiles.TryDequeue(out var pf))
                            {
                                if (System.IO.Directory.Exists(pf))
                                {
                                    updated = fileListService.AddFilesFromDirectory(updated, pf);
                                }
                                else if (System.IO.File.Exists(pf))
                                {
                                    updated = fileListService.AddFiles(updated, new[] { pf });
                                }
                            }
                            foreach (var f in updated)
                            {
                                if (!vm.SelectedFiles.Contains(f))
                                    vm.SelectedFiles.Add(f);
                            }
                        }
                    }
                });
            };

            base.OnStartup(e);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Console.WriteLine("[DEBUG] App.OnExit called. Application is exiting.");
            // Cleanup resources if necessary
            base.OnExit(e);
        }
    }
}