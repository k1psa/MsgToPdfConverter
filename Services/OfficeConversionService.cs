using System;
using System.Collections.Generic;
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
            return TryConvertOfficeToPdf(inputPath, outputPdf, null);
        }

        /// <summary>
        /// Attempts to convert Office files to PDF using Office Interop with progress callback for embedding operations
        /// </summary>
        public static bool TryConvertOfficeToPdf(string inputPath, string outputPdf, Action progressTick)
        {
            string ext = Path.GetExtension(inputPath).ToLowerInvariant();
            bool result = false;
            Exception threadEx = null;
            var thread = new System.Threading.Thread(() =>
            {
                try
                {
                    // Extract embedded OLE objects before PDF export (for both Word and Excel)
                    var extractedObjects = new List<MsgToPdfConverter.Utils.InteropEmbeddedExtractor.ExtractedObjectInfo>();
                    string tempDir = null;
                    try
                    {
                        tempDir = Path.Combine(Path.GetTempPath(), "MsgToPdf_Embedded_" + Guid.NewGuid());
                        Directory.CreateDirectory(tempDir);
                        extractedObjects = MsgToPdfConverter.Utils.InteropEmbeddedExtractor.ExtractEmbeddedObjects(inputPath, tempDir);
#if DEBUG
                        DebugLogger.Log($"[InteropExtractor] Total extracted objects: {extractedObjects.Count}");
#endif
                    }
                    catch (Exception ex)
                    {
#if DEBUG
                        DebugLogger.Log($"[InteropExtractor] Extraction failed: {ex.Message}");
#endif
                    }
                    progressTick?.Invoke(); // Tick: after embedded extraction

                    string mainPdfPath = outputPdf;
                    if (extractedObjects.Count > 0)
                    {
                        // If we have embedded objects, create the main PDF in a temp location first
                        mainPdfPath = Path.Combine(Path.GetTempPath(), $"main_pdf_{Guid.NewGuid()}.pdf");
                    }

                    if (ext == ".doc" || ext == ".docx")
                    {
                        var wordApp = new Microsoft.Office.Interop.Word.Application();
                        var doc = wordApp.Documents.Open(inputPath);
                        progressTick?.Invoke(); // Tick: after opening document
                        doc.ExportAsFixedFormat(mainPdfPath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                        progressTick?.Invoke(); // Tick: after exporting to PDF
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        wordApp.Quit();
                        Marshal.ReleaseComObject(wordApp);
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
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
                            wb.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, mainPdfPath);
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
                    }

                    // Insert embedded files into the PDF if any were extracted (for both Word and Excel)
                    if (extractedObjects.Count > 0)
                    {
                        try
                        {
#if DEBUG
                            DebugLogger.Log($"[PDF-EMBED] Inserting {extractedObjects.Count} embedded files into PDF");
#endif
                            PdfEmbeddedInsertionService.InsertEmbeddedFiles(mainPdfPath, extractedObjects, outputPdf, progressTick);
                            progressTick?.Invoke(); // Tick: after embedding
                            // Clean up temp main PDF
                            if (File.Exists(mainPdfPath) && mainPdfPath != outputPdf)
                            {
                                File.Delete(mainPdfPath);
                            }
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            DebugLogger.Log($"[PDF-EMBED] Failed to insert embedded files: {ex.Message}");
#endif
                            // Fallback: copy the main PDF without embedded files
                            if (mainPdfPath != outputPdf)
                            {
                                File.Copy(mainPdfPath, outputPdf, true);
                                File.Delete(mainPdfPath);
                            }
                        }
                    }
                    // Clean up temp extraction directory
                    if (!string.IsNullOrEmpty(tempDir) && Directory.Exists(tempDir))
                    {
                        try
                        {
                            Directory.Delete(tempDir, true);
                        }
                        catch (Exception ex)
                        {
#if DEBUG
                            DebugLogger.Log($"[CLEANUP] Failed to delete temp directory {tempDir}: {ex.Message}");
#endif
                        }
                    }
                    result = true;
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

#if DEBUG
                DebugLogger.Log($"[Interop] Waiting for Office to release PDF file: {outputPdf}");
#endif
                int[] delays = { 100, 200, 300, 500, 500, 500, 1000, 1000, 1000, 1000 };
                for (int i = 0; i < delays.Length; i++)
                {
                    System.Threading.Thread.Sleep(delays[i]);
                    progressTick?.Invoke(); // Tick: show progress during wait
                    // Try to open the PDF file to verify it's not locked
                    try
                    {
                        using (var fs = new FileStream(outputPdf, FileMode.Open, FileAccess.Read, FileShare.Read))
                        {
                            // If we can open it, it's not locked

#if DEBUG
                            DebugLogger.Log($"[Interop] PDF file ready after {delays.Take(i + 1).Sum()}ms: {outputPdf}");
#endif
                            break;
                        }
                    }
                    catch (IOException)
                    {
                        if (i == delays.Length - 1) // Last attempt
                        {

#if DEBUG
                            DebugLogger.Log($"[Interop][WARNING] PDF file may still be locked after {delays.Sum()}ms: {outputPdf}");
#endif
                        }
                    }
                }
            }

            if (threadEx != null)
            {

#if DEBUG
                DebugLogger.Log($"[Interop] Office to PDF conversion failed: {threadEx.Message}");
#endif
                return false;
            }
            return result;
        }

        // Add a helper for safe progress ticking (to be used in the ViewModel or tick callback):
        // public void SafeProgressTick() {
        //     if (FileProgressValue < FileProgressMax)
        //         FileProgressValue++;
        //     else
        //         FileProgressValue = FileProgressMax;
        // }
        // Only call FileProgressValue = 0 and FileProgressMax = N at the start of a new file, not after finishing one.
        // In OfficeConversionService.cs, do not call progressTick after the file is done, and only reset at the start of a new file.
    }
}
