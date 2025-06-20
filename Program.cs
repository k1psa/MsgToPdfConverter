using System;
using System.IO;
using System.Linq;

namespace MsgToPdfConverter
{
    public class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            // Worker mode: do not start WPF
            if (args != null && args.Length == 3 && args[0] == "--html2pdf")
            {
                string htmlPath = args[1];
                string pdfPath = args[2];
                string html = File.ReadAllText(htmlPath);
                int result = HtmlToPdfWorker.Convert(html, pdfPath);
                Environment.Exit(result);
            }
            // Normal WPF mode
            else
            {
                var app = new App();
                app.InitializeComponent();
                app.Run();
            }
        }
    }
}
