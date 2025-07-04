using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MsgToPdfConverter.Utils
{
    public class EmbeddedFileInfo
    {
        public string FileName { get; set; }
        public string ContentType { get; set; }
        public byte[] Data { get; set; }
        public int ParagraphIndex { get; set; } // Where in the doc the object appears
    }

    public static class DocxEmbeddedExtractor
    {
        /// <summary>
        /// Extracts embedded files (OLE objects, packages) from a .docx file.
        /// </summary>
        /// <param name="docxPath">Path to the .docx file</param>
        /// <returns>List of embedded files with their data and position</returns>
        public static List<EmbeddedFileInfo> ExtractEmbeddedFiles(string docxPath)
        {
            Console.WriteLine($"[DEBUG] ExtractEmbeddedFiles called for: {docxPath}");
            var result = new List<EmbeddedFileInfo>();
            using (var doc = WordprocessingDocument.Open(docxPath, false))
            {
                var mainPart = doc.MainDocumentPart;
                if (mainPart == null) { Console.WriteLine("[DEBUG] mainPart is null"); return result; }
                var body = mainPart.Document.Body;
                if (body == null) { Console.WriteLine("[DEBUG] body is null"); return result; }
                int paraIndex = 0;
                foreach (var para in body.Elements<Paragraph>())
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var obj in run.Elements().Where(e => e.LocalName == "object"))
                        {
                            var oleObj = obj.Elements().FirstOrDefault(e => e.LocalName == "OLEObject");
                            if (oleObj != null)
                            {
                                var relIdAttr = oleObj.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                                if (relIdAttr != null)
                                {
                                    string relId = relIdAttr.Value;
                                    var part = mainPart.GetPartById(relId);
                                    if (part is EmbeddedPackagePart pkgPart)
                                    {
                                        using (var ms = new MemoryStream())
                                        {
                                            pkgPart.GetStream().CopyTo(ms);
                                            // Try to extract real file from OLE package
                                            var oleInfo = OlePackageExtractor.ExtractPackage(ms.ToArray());
                                            if (oleInfo != null)
                                            {
                                                result.Add(new EmbeddedFileInfo
                                                {
                                                    FileName = oleInfo.FileName,
                                                    ContentType = oleInfo.ContentType,
                                                    Data = oleInfo.Data,
                                                    ParagraphIndex = paraIndex
                                                });
                                                Console.WriteLine($"[DEBUG] Extracted OLE-embedded file: {oleInfo.FileName}, ContentType: {oleInfo.ContentType}, ParagraphIndex: {paraIndex}");
                                            }
                                            else
                                            {
                                                result.Add(new EmbeddedFileInfo
                                                {
                                                    FileName = pkgPart.Uri.ToString(),
                                                    ContentType = pkgPart.ContentType,
                                                    Data = ms.ToArray(),
                                                    ParagraphIndex = paraIndex
                                                });
                                                Console.WriteLine($"[DEBUG] Extracted EmbeddedPackagePart (raw): {pkgPart.Uri}, ContentType: {pkgPart.ContentType}, ParagraphIndex: {paraIndex}");
                                            }
                                        }
                                    }
                                    else if (part is EmbeddedObjectPart objPart)
                                    {
                                        using (var ms = new MemoryStream())
                                        {
                                            objPart.GetStream().CopyTo(ms);
                                            // Try to extract real file from OLE package
                                            var oleInfo = OlePackageExtractor.ExtractPackage(ms.ToArray());
                                            if (oleInfo != null)
                                            {
                                                result.Add(new EmbeddedFileInfo
                                                {
                                                    FileName = oleInfo.FileName,
                                                    ContentType = oleInfo.ContentType,
                                                    Data = oleInfo.Data,
                                                    ParagraphIndex = paraIndex
                                                });
                                                Console.WriteLine($"[DEBUG] Extracted OLE-embedded file: {oleInfo.FileName}, ContentType: {oleInfo.ContentType}, ParagraphIndex: {paraIndex}");
                                            }
                                            else
                                            {
                                                result.Add(new EmbeddedFileInfo
                                                {
                                                    FileName = objPart.Uri.ToString(),
                                                    ContentType = objPart.ContentType,
                                                    Data = ms.ToArray(),
                                                    ParagraphIndex = paraIndex
                                                });
                                                Console.WriteLine($"[DEBUG] Extracted EmbeddedObjectPart (raw): {objPart.Uri}, ContentType: {objPart.ContentType}, ParagraphIndex: {paraIndex}");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        foreach (var vmlOle in run.Descendants().Where(e => e.LocalName == "oleObject"))
                        {
                            var relIdAttr = vmlOle.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                            if (relIdAttr != null)
                            {
                                string relId = relIdAttr.Value;
                                var part = mainPart.GetPartById(relId);
                                if (part is EmbeddedPackagePart pkgPart)
                                {
                                    using (var ms = new MemoryStream())
                                    {
                                        pkgPart.GetStream().CopyTo(ms);
                                        var oleInfo = OlePackageExtractor.ExtractPackage(ms.ToArray());
                                        if (oleInfo != null)
                                        {
                                            result.Add(new EmbeddedFileInfo
                                            {
                                                FileName = oleInfo.FileName,
                                                ContentType = oleInfo.ContentType,
                                                Data = oleInfo.Data,
                                                ParagraphIndex = paraIndex
                                            });
                                            Console.WriteLine($"[DEBUG] Extracted OLE-embedded file (VML): {oleInfo.FileName}, ContentType: {oleInfo.ContentType}, ParagraphIndex: {paraIndex}");
                                        }
                                        else
                                        {
                                            result.Add(new EmbeddedFileInfo
                                            {
                                                FileName = pkgPart.Uri.ToString(),
                                                ContentType = pkgPart.ContentType,
                                                Data = ms.ToArray(),
                                                ParagraphIndex = paraIndex
                                            });
                                            Console.WriteLine($"[DEBUG] Extracted EmbeddedPackagePart (VML, raw): {pkgPart.Uri}, ContentType: {pkgPart.ContentType}, ParagraphIndex: {paraIndex}");
                                        }
                                    }
                                }
                                else if (part is EmbeddedObjectPart objPart)
                                {
                                    using (var ms = new MemoryStream())
                                    {
                                        objPart.GetStream().CopyTo(ms);
                                        var oleInfo = OlePackageExtractor.ExtractPackage(ms.ToArray());
                                        if (oleInfo != null)
                                        {
                                            result.Add(new EmbeddedFileInfo
                                            {
                                                FileName = oleInfo.FileName,
                                                ContentType = oleInfo.ContentType,
                                                Data = oleInfo.Data,
                                                ParagraphIndex = paraIndex
                                            });
                                            Console.WriteLine($"[DEBUG] Extracted OLE-embedded file (VML): {oleInfo.FileName}, ContentType: {oleInfo.ContentType}, ParagraphIndex: {paraIndex}");
                                        }
                                        else
                                        {
                                            result.Add(new EmbeddedFileInfo
                                            {
                                                FileName = objPart.Uri.ToString(),
                                                ContentType = objPart.ContentType,
                                                Data = ms.ToArray(),
                                                ParagraphIndex = paraIndex
                                            });
                                            Console.WriteLine($"[DEBUG] Extracted EmbeddedObjectPart (VML, raw): {objPart.Uri}, ContentType: {objPart.ContentType}, ParagraphIndex: {paraIndex}");
                                        }
                                    }
                                }
                            }
                        }
                    }
                    paraIndex++;
                }
                var alreadyAdded = new HashSet<string>(result.Select(r => r.FileName));
                foreach (var pkgPart in mainPart.EmbeddedPackageParts)
                {
                    if (!alreadyAdded.Contains(pkgPart.Uri.ToString()))
                    {
                        using (var ms = new MemoryStream())
                        {
                            pkgPart.GetStream().CopyTo(ms);
                            var oleInfo = OlePackageExtractor.ExtractPackage(ms.ToArray());
                            if (oleInfo != null)
                            {
                                result.Add(new EmbeddedFileInfo
                                {
                                    FileName = oleInfo.FileName,
                                    ContentType = oleInfo.ContentType,
                                    Data = oleInfo.Data,
                                    ParagraphIndex = -1
                                });
                                Console.WriteLine($"[DEBUG] Extracted OLE-embedded file (unreferenced): {oleInfo.FileName}, ContentType: {oleInfo.ContentType}, ParagraphIndex: -1");
                            }
                            else
                            {
                                result.Add(new EmbeddedFileInfo
                                {
                                    FileName = pkgPart.Uri.ToString(),
                                    ContentType = pkgPart.ContentType,
                                    Data = ms.ToArray(),
                                    ParagraphIndex = -1
                                });
                                Console.WriteLine($"[DEBUG] Extracted EmbeddedPackagePart (unreferenced, raw): {pkgPart.Uri}, ContentType: {pkgPart.ContentType}, ParagraphIndex: -1");
                            }
                        }
                    }
                }
            }
            Console.WriteLine($"[DEBUG] ExtractEmbeddedFiles returning {result.Count} embedded files");
            return result;
        }

        /// <summary>
        /// Logs all part URIs and content types in the .docx for debugging.
        /// </summary>
        public static void LogAllParts(string docxPath)
        {
            using (var doc = WordprocessingDocument.Open(docxPath, false))
            {
                var mainPart = doc.MainDocumentPart;
                if (mainPart == null) return;
                Console.WriteLine($"[DEBUG] Parts in {docxPath}:");
                foreach (var part in mainPart.Parts)
                {
                    Console.WriteLine($"  [Part] URI: {part.OpenXmlPart.Uri}, ContentType: {part.OpenXmlPart.ContentType}");
                }
                // Also log EmbeddedPackageParts and EmbeddedObjectParts
                foreach (var pkg in mainPart.EmbeddedPackageParts)
                {
                    Console.WriteLine($"  [EmbeddedPackagePart] URI: {pkg.Uri}, ContentType: {pkg.ContentType}");
                }
                foreach (var obj in mainPart.EmbeddedObjectParts)
                {
                    Console.WriteLine($"  [EmbeddedObjectPart] URI: {obj.Uri}, ContentType: {obj.ContentType}");
                }
            }
        }
    }
}
