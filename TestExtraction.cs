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
#if DEBUG
                DebugLogger.Log($"Test file {testFile} not found");
#endif
                return;
            }

            #if DEBUG
            DebugLogger.Log($"Testing extraction from {testFile}");
            DebugLogger.Log("========================================");
            #endif

            try
            {
                string tempDir = Path.Combine(Path.GetTempPath(), "MsgToPdf_Test_" + Guid.NewGuid());
                Directory.CreateDirectory(tempDir);
                #if DEBUG
                DebugLogger.Log($"Temp directory: {tempDir}");
                #endif

                var embedded = InteropEmbeddedExtractor.ExtractEmbeddedObjects(testFile, tempDir);
                #if DEBUG
                DebugLogger.Log($"Total extracted objects: {embedded.Count}");
                #endif

                foreach (var obj in embedded)
                {
                    string fileName = Path.GetFileName(obj.FilePath);
                    long fileSize = 0;
                    if (File.Exists(obj.FilePath))
                    {
                        fileSize = new FileInfo(obj.FilePath).Length;
                    }
#if DEBUG
                    DebugLogger.Log($"- {fileName} (Page {obj.PageNumber}, OLE Class: {obj.OleClass}) - {fileSize} bytes");
#endif
                }

                #if DEBUG
                DebugLogger.Log("\nExtraction test completed!");
                DebugLogger.Log($"Files saved in: {tempDir}");
                #endif
            }
            catch (Exception ex)
            {
#if DEBUG
                DebugLogger.Log($"Test failed: {ex.Message}");
                DebugLogger.Log($"Stack trace: {ex.StackTrace}");
#endif
            }
        }
    }
}
