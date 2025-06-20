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
                }
                Shutdown();
                return;
            }
            base.OnStartup(e);
            // Additional startup logic can be added here
        }

        private bool IsDotNetDesktopRuntimeInstalled()
        {
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

        protected override void OnExit(ExitEventArgs e)
        {
            // Cleanup resources if necessary
            base.OnExit(e);
        }
    }
}