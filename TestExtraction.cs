using System;
using System.IO;
using MsgToPdfConverter.Utils;

namespace MsgToPdfConverter
{
    class TestExtraction
    {
        public static void TestEmbeddedExtraction(string testFile)
        {
            if (!File.Exists(testFile))
            {
                Console.WriteLine($"Test file {testFile} not found");
                return;
            }

            Console.WriteLine($"Testing extraction from {testFile}");
            Console.WriteLine("========================================");

            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "MsgToPdf_Test_" + Guid.NewGuid());
                Directory.CreateDirectory(tempDir);
                Console.WriteLine($"Temp directory: {tempDir}");

                var embedded = InteropEmbeddedExtractor.ExtractEmbeddedObjects(testFile, tempDir);
                Console.WriteLine($"Total extracted objects: {embedded.Count}");

                foreach (var obj in embedded)
                {
                    string fileName = Path.GetFileName(obj.FilePath);
                    long fileSize = 0;
                    if (File.Exists(obj.FilePath))
                    {
                        fileSize = new FileInfo(obj.FilePath).Length;
                    }
                    Console.WriteLine($"- {fileName} (Page {obj.PageNumber}, OLE Class: {obj.OleClass}) - {fileSize} bytes");
                }

                Console.WriteLine("\nExtraction test completed!");
                Console.WriteLine($"Files saved in: {tempDir}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Test failed: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }
    }
}
