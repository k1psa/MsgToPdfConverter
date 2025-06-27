using System;
using System.Diagnostics;
using System.IO;
using System.Windows;

namespace MsgToPdfConverter.Services
{
    public class RuntimeService
    {
        public void CheckDotNetRuntime()
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
    }
}
