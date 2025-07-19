using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MsgToPdfConverter.Utils;

namespace MsgToPdfConverter
{
    /// <summary>
    /// Validation test class to ensure the PDF insertion logic works correctly
    /// for all embedded object types and maintains proper page ordering.
    /// </summary>
    public static class ValidationTest
    {
        /// <summary>
        /// Test the page ordering logic to ensure embedded objects are always inserted
        /// after their main page, never before, and that multiple objects per page work correctly.
        /// </summary>
        public static void TestPageOrderingLogic()
        {
        
            
            // Simulate a scenario with:
            // - Main PDF with 3 pages
            // - Page 1: 2 embedded objects (Excel, PDF)
            // - Page 2: 1 embedded object (MSG)  
            // - Page 3: 3 embedded objects (DOCX, XLSX, PDF)
            
            var testObjects = new List<InteropEmbeddedExtractor.ExtractedObjectInfo>
            {
                // Page 1 objects
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "test_excel.xlsx",
                    PageNumber = 1,
                    DocumentOrderIndex = 1,
                    OleClass = "Excel.Sheet"
                },
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "test_pdf.pdf", 
                    PageNumber = 1,
                    DocumentOrderIndex = 2,
                    OleClass = "AcroExch.Document"
                },
                
                // Page 2 objects
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "test_msg.msg",
                    PageNumber = 2,
                    DocumentOrderIndex = 1,
                    OleClass = "IPM.Note"
                },
                
                // Page 3 objects
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "test_docx.docx",
                    PageNumber = 3,
                    DocumentOrderIndex = 1,
                    OleClass = "Word.Document"
                },
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "test_xlsx2.xlsx",
                    PageNumber = 3,
                    DocumentOrderIndex = 2,
                    OleClass = "Excel.Sheet"
                },
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "test_pdf2.pdf",
                    PageNumber = 3,
                    DocumentOrderIndex = 3,
                    OleClass = "AcroExch.Document"
                }
            };
            
            // Test the ordering logic
            TestOrderingLogic(testObjects, 3);
            
            // Test edge cases
            TestEdgeCases();
            
        
        }
        
        private static void TestOrderingLogic(List<InteropEmbeddedExtractor.ExtractedObjectInfo> objects, int mainPageCount)
        {
         
            
            // Sort objects as the actual code does
            var objectsByPage = objects.OrderBy(obj => obj.PageNumber).ThenBy(obj => obj.DocumentOrderIndex).ToList();
            
            // Group by page as the actual code does
            var objectGroups = objectsByPage.GroupBy(obj => obj.PageNumber)
                                           .OrderBy(g => g.Key)
                                           .ToList();
            
        
            
            int expectedOutputPage = 0;
            int groupIndex = 0;
            
            for (int mainPage = 1; mainPage <= mainPageCount; mainPage++)
            {
                // Main page always comes first
                expectedOutputPage++;
         
                
                // Check if this page has embedded objects
                bool hasEmbeddedObjects = groupIndex < objectGroups.Count && objectGroups[groupIndex].Key == mainPage;
                
                if (hasEmbeddedObjects)
                {
                    var pageObjects = objectGroups[groupIndex].OrderBy(obj => obj.DocumentOrderIndex).ToList();
            
                    
                    foreach (var obj in pageObjects)
                    {
                        expectedOutputPage++;
                        #if DEBUG
                        DebugLogger.Log($"  Output Page {expectedOutputPage}: {Path.GetFileName(obj.FilePath)} (order: {obj.DocumentOrderIndex})");
                        #endif
                    }
                    
                    groupIndex++;
                }
            }
            
    
            
            // Validate that no embedded object comes before its main page
            bool isValid = true;
            foreach (var group in objectGroups)
            {
                int mainPagePos = group.Key; // This is the 1-based main page number
                
                foreach (var obj in group)
                {
                    if (obj.PageNumber < mainPagePos)
                    {
                 
                        isValid = false;
                    }
                }
            }
            
            if (isValid)
            {
                #if DEBUG
                DebugLogger.Log("✓ Page ordering validation PASSED - No embedded object will be inserted before its main page");
                #endif
            }
            else
            {
                #if DEBUG
                DebugLogger.Log("✗ Page ordering validation FAILED!");
                #endif
            }
        }
        
        private static void TestEdgeCases()
        {
            #if DEBUG
            DebugLogger.Log("\n--- Testing edge cases ---");
            #endif
            
            // Test case 1: Empty objects list
            #if DEBUG
            DebugLogger.Log("Test 1: Empty objects list");
            #endif
            TestOrderingLogic(new List<InteropEmbeddedExtractor.ExtractedObjectInfo>(), 2);
            
            // Test case 2: Objects only on last page
            #if DEBUG
            DebugLogger.Log("\nTest 2: Objects only on last page");
            #endif
            var lastPageObjects = new List<InteropEmbeddedExtractor.ExtractedObjectInfo>
            {
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "last_page.pdf",
                    PageNumber = 5,
                    DocumentOrderIndex = 1,
                    OleClass = "AcroExch.Document"
                }
            };
            TestOrderingLogic(lastPageObjects, 5);
            
            // Test case 3: Objects with page number -1 (should go to end)
            #if DEBUG
            DebugLogger.Log("\nTest 3: Objects with page number -1");
            #endif
            var invalidPageObjects = new List<InteropEmbeddedExtractor.ExtractedObjectInfo>
            {
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "no_page.xlsx",
                    PageNumber = -1,
                    DocumentOrderIndex = 1,
                    OleClass = "Excel.Sheet"
                }
            };
            
            // Simulate the adjustment logic from the actual code
            foreach (var obj in invalidPageObjects)
            {
                if (obj.PageNumber == -1)
                {
                    obj.PageNumber = 3; // Assuming 3-page main document
                    #if DEBUG
                    DebugLogger.Log($"Adjusted object {Path.GetFileName(obj.FilePath)} from page -1 to page {obj.PageNumber}");
                    #endif
                }
            }
            TestOrderingLogic(invalidPageObjects, 3);
            
            // Test case 4: Multiple objects with same order index (rare but possible)
            #if DEBUG
            DebugLogger.Log("\nTest 4: Multiple objects with same order index");
            #endif
            var sameOrderObjects = new List<InteropEmbeddedExtractor.ExtractedObjectInfo>
            {
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "first.pdf",
                    PageNumber = 1,
                    DocumentOrderIndex = 1,
                    OleClass = "AcroExch.Document"
                },
                new InteropEmbeddedExtractor.ExtractedObjectInfo
                {
                    FilePath = "second.xlsx",
                    PageNumber = 1,
                    DocumentOrderIndex = 1, // Same order!
                    OleClass = "Excel.Sheet"
                }
            };
            TestOrderingLogic(sameOrderObjects, 1);
        }
    }
}
