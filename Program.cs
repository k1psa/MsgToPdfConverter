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
            // Test mode: test embedded extraction
            if (args != null && args.Length >= 1 && args[0] == "--test")
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
            }            // Normal WPF mode
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
    }
}
