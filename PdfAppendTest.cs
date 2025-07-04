using System;
using System.Collections.Generic;
using System.IO;
using iText.Kernel.Pdf;

namespace MsgToPdfConverter
{
    public static class PdfAppendTest
    {
        // Appends multiple PDFs into a new output PDF using iText7
        public static void AppendPdfs(List<string> inputFiles, string outputFile)
        {
            using (var pdfWriter = new PdfWriter(outputFile))
            using (var pdfDoc = new PdfDocument(pdfWriter))
            {
                foreach (var file in inputFiles)
                {
                    using (var srcPdf = new PdfDocument(new PdfReader(file)))
                    {
                        srcPdf.CopyPagesTo(1, srcPdf.GetNumberOfPages(), pdfDoc);
                    }
                }
            }
        }

        // Simple test entry point
        public static void RunTest()
        {
            var inputFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PdfAppendTestInput.txt");
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }
            var lines = File.ReadAllLines(inputFile);
            // Filter out comments and blank lines
            var filtered = new List<string>();
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (!string.IsNullOrEmpty(trimmed) && !trimmed.StartsWith("#"))
                    filtered.Add(trimmed);
            }
            if (filtered.Count < 3)
            {
                Console.WriteLine("Please provide at least two input PDF paths and one output PDF path in PdfAppendTestInput.txt");
                return;
            }
            // If the last input before output is a JPG, convert it to PDF and use that in the merge
            var pdfsToMerge = new List<string>(filtered);
            var output = pdfsToMerge[pdfsToMerge.Count - 1];
            pdfsToMerge.RemoveAt(pdfsToMerge.Count - 1);
            // Check for JPG as last input
            string lastInput = pdfsToMerge[pdfsToMerge.Count - 1];
            if (lastInput.EndsWith(".jpg", StringComparison.OrdinalIgnoreCase) || lastInput.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase))
            {
                string jpgPdf = Path.ChangeExtension(lastInput, ".pdf");
                if (File.Exists(lastInput))
                {
                    CreatePdfFromJpg(lastInput, jpgPdf);
                    Console.WriteLine($"Created PDF from JPG: {jpgPdf}");
                    pdfsToMerge[pdfsToMerge.Count - 1] = jpgPdf;
                }
                else
                {
                    Console.WriteLine($"JPG file not found: {lastInput}");
                    pdfsToMerge.RemoveAt(pdfsToMerge.Count - 1);
                }
            }
            Console.WriteLine("Merging the following PDFs:");
            foreach (var f in pdfsToMerge) Console.WriteLine(f);
            Console.WriteLine($"Output: {output}");
            AppendPdfs(pdfsToMerge, output);
            Console.WriteLine($"Merged PDF created at: {output}");
        }

        // Helper to create a simple one-page PDF with text
        private static void CreateSimplePdf(string path, string text)
        {
            using (var writer = new PdfWriter(path))
            using (var pdf = new PdfDocument(writer))
            {
                var page = pdf.AddNewPage();
                var canvas = new iText.Kernel.Pdf.Canvas.PdfCanvas(page);
                var font = iText.Kernel.Font.PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA);
                canvas.BeginText().SetFontAndSize(font, 24).MoveText(50, 700).ShowText(text).EndText();
            }
        }

        // Helper to create a PDF from a JPG image
        public static void CreatePdfFromJpg(string jpgPath, string pdfPath)
        {
            using (var writer = new PdfWriter(pdfPath))
            using (var pdf = new PdfDocument(writer))
            {
                var doc = new iText.Layout.Document(pdf);
                var imgData = iText.IO.Image.ImageDataFactory.Create(jpgPath);
                var image = new iText.Layout.Element.Image(imgData);
                doc.Add(image);
                doc.Close();
            }
        }

        // Main method for standalone execution
        public static void Main(string[] args)
        {
            // Example: create a PDF from a JPG for testing
            string jpgPath = "test.jpg"; // Place a test.jpg in the output directory
            string pdfPath = "test_from_jpg.pdf";
            if (File.Exists(jpgPath))
            {
                CreatePdfFromJpg(jpgPath, pdfPath);
                Console.WriteLine($"Created PDF from JPG: {pdfPath}");
            }
            else
            {
                Console.WriteLine($"JPG file not found: {jpgPath}");
            }
            RunTest();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
