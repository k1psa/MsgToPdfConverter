using System;
using DinkToPdf;
using DinkToPdf.Contracts;

namespace MsgToPdfConverter
{
    public class HtmlToPdfWorker
    {
        public static int Convert(string html, string outputPath)
        {
            try
            {
                var doc = new HtmlToPdfDocument()
                {
                    GlobalSettings = new GlobalSettings
                    {
                        ColorMode = ColorMode.Color,
                        Orientation = Orientation.Portrait,
                        PaperSize = PaperKind.A4,
                        Out = outputPath
                    },
                    Objects = { new ObjectSettings { HtmlContent = html } }
                };
                var converter = new SynchronizedConverter(new PdfTools());
                converter.Convert(doc);
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[WORKER ERROR] {ex}");
                return 1;
            }
        }
    }
}
