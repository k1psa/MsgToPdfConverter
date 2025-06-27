using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace MsgToPdfConverter.Services
{
    public static class OfficeConversionService
    {
        /// <summary>
        /// Attempts to convert Office files to PDF using Office Interop (requires Office installed)
        /// </summary>
        public static bool TryConvertOfficeToPdf(string inputPath, string outputPdf)
        {
            string ext = Path.GetExtension(inputPath).ToLowerInvariant();
            bool result = false;
            Exception threadEx = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    if (ext == ".doc" || ext == ".docx")
                    {
                        var wordApp = new Microsoft.Office.Interop.Word.Application();
                        var doc = wordApp.Documents.Open(inputPath);
                        doc.ExportAsFixedFormat(outputPdf, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        result = true;
                    }
                    else if (ext == ".xls" || ext == ".xlsx")
                    {
                        var excelApp = new Microsoft.Office.Interop.Excel.Application();
                        Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
                        Microsoft.Office.Interop.Excel.Workbook wb = null;
                        try
                        {
                            workbooks = excelApp.Workbooks;
                            wb = workbooks.Open(inputPath);
                            wb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPdf);
                        }
                        finally
                        {
                            if (wb != null)
                            {
                                wb.Close(false);
                                Marshal.ReleaseComObject(wb);
                            }
                            if (workbooks != null)
                            {
                                Marshal.ReleaseComObject(workbooks);
                            }
                            if (excelApp != null)
                            {
                                excelApp.Quit();
                                Marshal.ReleaseComObject(excelApp);
                            }
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        result = true;
                    }
                }
                catch (Exception ex)
                {
                    threadEx = ex;
                }
            });
            
            thread.SetApartmentState(System.Threading.ApartmentState.STA);
            thread.Start();
            thread.Join();
            
            // Give Office extra time to release the generated PDF file
            if (result)
            {
                Console.WriteLine($"[Interop] Waiting for Office to release PDF file: {outputPdf}");

                // Wait and verify the PDF is not locked (start with shorter delays)
                int[] delays = { 100, 200, 300, 500, 500, 500, 1000, 1000, 1000, 1000 };
                for (int i = 0; i < delays.Length; i++)
                {
                    System.Threading.Thread.Sleep(delays[i]);

                    // Try to open the PDF file to verify it's not locked
                    try
                    {
                        using (var fs = new FileStream(outputPdf, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            // If we can open it, it's not locked
                            Console.WriteLine($"[Interop] PDF file ready after {delays.Take(i + 1).Sum()}ms: {outputPdf}");
                            break;
                        }
                    }
                    catch (IOException)
                    {
                        if (i == delays.Length - 1) // Last attempt
                        {
                            Console.WriteLine($"[Interop][WARNING] PDF file may still be locked after {delays.Sum()}ms: {outputPdf}");
                        }
                    }
                }
            }

            if (threadEx != null)
            {
                Console.WriteLine($"[Interop] Office to PDF conversion failed: {threadEx.Message}");
                return false;
            }
            return result;
        }
    }
}
