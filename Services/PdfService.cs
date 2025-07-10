using System;
using System.Collections.Generic;
using System.IO;

namespace MsgToPdfConverter.Services
{
    public static class PdfService
    {
        public static void AddPlaceholderPdf(string pdfPath, string message, string imagePath = null)
        {
            using (var doc = new PdfSharp.Pdf.PdfDocument())
            {
                var page = doc.AddPage();
                using (var gfx = PdfSharp.Drawing.XGraphics.FromPdfPage(page))
                {
                    if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                    {
                        try
                        {
                            var img = PdfSharp.Drawing.XImage.FromFile(imagePath);
                            double maxWidth = page.Width.Point - 80;
                            double maxHeight = page.Height.Point - 300;
                            double scale = Math.Min(maxWidth / img.PixelWidth * 72.0 / img.HorizontalResolution, maxHeight / img.PixelHeight * 72.0 / img.VerticalResolution);
                            double imgWidth = img.PixelWidth * 72.0 / img.HorizontalResolution * scale;
                            double imgHeight = img.PixelHeight * 72.0 / img.VerticalResolution * scale;
                            double x = (page.Width.Point - imgWidth) / 2;
                            double y = 100;
                            gfx.DrawImage(img, x, y, imgWidth, imgHeight);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"[IMG2PDF] Failed to load image: {imagePath} - {ex.Message}");
                        }
                    }
                    var font = new PdfSharp.Drawing.XFont("Arial", 16);
                    gfx.DrawString(message, font, PdfSharp.Drawing.XBrushes.Black, new PdfSharp.Drawing.XRect(40, page.Height.Point - 200, page.Width.Point - 80, 100), PdfSharp.Drawing.XStringFormats.Center);
                }
                doc.Save(pdfPath);
            }
        }

        public static void AddHeaderPdf(string pdfPath, string headerText)
        {
            using (var writer = new iText.Kernel.Pdf.PdfWriter(pdfPath))
            using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
            using (var doc = new iText.Layout.Document(pdf))
            {
                var p = new iText.Layout.Element.Paragraph(headerText)
                    .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                    .SetFontSize(18);
                doc.Add(p);
            }
        }

        public static void MergePdfs(string[] pdfFiles, string outputPdf)
        {
            var inputFiles = new List<string>();
            foreach (var f in pdfFiles)
            {
                if (!string.Equals(f, outputPdf, StringComparison.OrdinalIgnoreCase))
                    inputFiles.Add(f);
            }
            var validInputFiles = new List<string>();
            foreach (var pdf in inputFiles)
            {
                try
                {
                    using (var reader = new iText.Kernel.Pdf.PdfReader(pdf))
                    using (var doc = new iText.Kernel.Pdf.PdfDocument(reader))
                    {
                        int n = doc.GetNumberOfPages();
                        if (n > 0)
                        {
                            validInputFiles.Add(pdf);
                        }
                    }
                }
                catch { }
            }
            if (validInputFiles.Count == 0)
                return;
            using (var stream = new FileStream(outputPdf, FileMode.Create, FileAccess.Write))
            using (var pdfWriter = new iText.Kernel.Pdf.PdfWriter(stream))
            using (var pdfDoc = new iText.Kernel.Pdf.PdfDocument(pdfWriter))
            {
                foreach (var pdf in validInputFiles)
                {
                    using (var srcPdf = new iText.Kernel.Pdf.PdfDocument(new iText.Kernel.Pdf.PdfReader(pdf)))
                    {
                        int n = srcPdf.GetNumberOfPages();
                        srcPdf.CopyPagesTo(1, n, pdfDoc);
                    }
                }
            }
        }

        public static void AddImagePdf(string pdfPath, string imagePath, string headerText = null)
        {
            try
            {
                using (var writer = new iText.Kernel.Pdf.PdfWriter(pdfPath))
                using (var pdf = new iText.Kernel.Pdf.PdfDocument(writer))
                using (var doc = new iText.Layout.Document(pdf))
                {
                    // Only add header if non-empty
                    if (!string.IsNullOrEmpty(headerText))
                    {
                        var header = new iText.Layout.Element.Paragraph(headerText)
                            .SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                            .SetFontSize(16)
                            .SetBold()
                            .SetMarginBottom(20);
                        doc.Add(header);
                    }

                    // Add image if it exists
                    if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
                    {
                        var imageData = iText.IO.Image.ImageDataFactory.Create(imagePath);
                        var image = new iText.Layout.Element.Image(imageData);
                        // Center the image and scale it to fit the page (A4 minus margin)
                        image.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
                        float maxWidth = doc.GetPdfDocument().GetDefaultPageSize().GetWidth() - 40; // 20pt margin each side
                        float maxHeight = doc.GetPdfDocument().GetDefaultPageSize().GetHeight() - 40; // 20pt margin top/bottom
                        image.SetMaxWidth(maxWidth);
                        image.SetMaxHeight(maxHeight);
                        doc.Add(image);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating image PDF: {ex.Message}");
                // Fallback to text-only PDF
                AddHeaderPdf(pdfPath, headerText ?? "Attachment Hierarchy");
            }
        }
    }
}
