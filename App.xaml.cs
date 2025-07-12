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

            base.OnStartup(e);
            // Additional startup logic can be added here
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Console.WriteLine("[DEBUG] App.OnExit called. Application is exiting.");
            // Cleanup resources if necessary
            base.OnExit(e);
        }
    }
}