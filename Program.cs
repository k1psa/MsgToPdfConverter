using System;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Windows;

namespace MsgToPdfConverter
{
    public class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            // Validation test mode: test page ordering logic
            if (args != null && args.Length >= 1 && args[0] == "--validate")
            {
                ValidationTest.TestPageOrderingLogic();
                return;
            }
            // Test mode: test embedded extraction
            else if (args != null && args.Length >= 1 && args[0] == "--test")
            {
                string testFile = args.Length > 1 ? args[1] : "a.docx";
                TestExtraction.TestEmbeddedExtraction(testFile);
                return;
            }
            // Worker mode: do not start WPF
            else if (args != null && args.Length == 3 && args[0] == "--html2pdf")
            {
                string htmlPath = args[1];
                string pdfPath = args[2];
                string html = File.ReadAllText(htmlPath);
                int result = HtmlToPdfWorker.Convert(html, pdfPath);
                Environment.Exit(result);
            }
            // Check .NET Framework before starting WPF
            else if (!IsDotNetFrameworkInstalled())
            {
                MessageBox.Show(
                    ".NET Framework 4.8 or higher is required to run this application.\n\n" +
                    "Please install it from:\n" +
                    "https://dotnet.microsoft.com/download/dotnet-framework",
                    ".NET Framework Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Environment.Exit(1);
            }
            // Check for Microsoft Office (Word and Excel) before starting WPF
            else if (!IsOfficeInstalled())
            {
                MessageBox.Show(
                    "Microsoft Office (Word and Excel) is required to convert Office documents.\n\n" +
                    "Please install Microsoft Office and try again.",
                    "Microsoft Office Not Found",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Environment.Exit(1);
            }
            // Normal WPF mode
            else
            {
                var app = new App();
                app.Run();
            }
        }

        private static bool IsDotNetFrameworkInstalled()
        {
            try
            {
                // Check if .NET Framework 4.8 or higher is installed
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"))
                {
                    if (key != null)
                    {
                        var release = key.GetValue("Release");
                        if (release != null && (int)release >= 528040) // .NET Framework 4.8
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            catch
            {
                return false; // Assume not installed if we can't check
            }
        }

        private static bool IsOfficeInstalled()
        {
            // Check for Word and Excel registry keys
            try
            {
                using (var wordKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("Word.Application"))
                using (var excelKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("Excel.Application"))
                {
                    return wordKey != null && excelKey != null;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
