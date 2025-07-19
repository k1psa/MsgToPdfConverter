using System;
using System.IO;

namespace MsgToPdfConverter
{
    class LogDocxParts
    {
        static void Main(string[] args)
        {
            string docxPath = @"C:\Users\kipsa\Desktop\output\embededPDFs.docx";
            string logPath = @"C:\Users\kipsa\Desktop\output\embededPDFs_parts.txt";
            using (var sw = new StreamWriter(logPath, false))
            {
                var origOut = Console.Out;
                Console.SetOut(sw);
                MsgToPdfConverter.Utils.DocxEmbeddedExtractor.LogAllParts(docxPath);
                Console.SetOut(origOut);
            }
     
        }
    }
}
