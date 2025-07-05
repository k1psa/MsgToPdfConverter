using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OpenMcdf;

namespace MsgToPdfConverter.Utils
{
    public class OlePackageInfo
    {
        public string FileName { get; set; }
        public string ContentType { get; set; }
        public byte[] Data { get; set; }
    }

    public static class OlePackageExtractor
    {
        /// <summary>
        /// Extracts the real embedded file (e.g. PDF) from an OLEObject .bin (as found in .docx embeddings)
        /// </summary>
        /// <param name="oleObjectBytes">The bytes of the OLEObject .bin</param>
        /// <returns>OlePackageInfo with file name, content type, and data, or null if not found</returns>
        public static OlePackageInfo ExtractPackage(byte[] oleObjectBytes)
        {
            using (var ms = new MemoryStream(oleObjectBytes))
            using (var cf = new CompoundFile(ms))
            {
                // Enumerate all streams using reflection for compatibility
                var streamNames = new System.Collections.Generic.List<string>();
                var rootType = cf.RootStorage.GetType();
                
                var getStreamNamesProp = rootType.GetProperty("StreamNames");
                if (getStreamNamesProp != null)
                {
                    var names = getStreamNamesProp.GetValue(cf.RootStorage) as System.Collections.IEnumerable;
                    if (names != null)
                    {
                        foreach (var n in names)
                            streamNames.Add(n.ToString());
                    }
                }
                else
                {
                    Console.WriteLine("[DEBUG] StreamNames property not found, trying GetStreamNames method");
                    // Fallback: try GetStreamNames method
                    var getStreamNamesMethod = rootType.GetMethod("GetStreamNames");
                    if (getStreamNamesMethod != null)
                    {
                        Console.WriteLine("[DEBUG] Found GetStreamNames method");
                        var names = getStreamNamesMethod.Invoke(cf.RootStorage, null) as System.Collections.IEnumerable;
                        if (names != null)
                        {
                            foreach (var n in names)
                                streamNames.Add(n.ToString());
                        }
                    }
                    else
                    {
                        // Try common stream names manually
                        string[] commonNames = { "Package", "CONTENTS", "Contents", "Data", "ObjectPool", "\u0001Ole", "\u0001CompObj", "\u0001Ole10Native" };
                        foreach (var name in commonNames)
                        {
                            try
                            {
                                var stream = cf.RootStorage.GetStream(name);
                                streamNames.Add(name);
                            }
                            catch { }
                        }
                    }
                }
                // Check each stream (reduced logging)
                foreach (var streamName in streamNames)
                {
                    try
                    {
                        var stream = cf.RootStorage.GetStream(streamName);
                        Console.WriteLine($"[Stream] {streamName} (size: {stream.Size})");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Stream] {streamName} (exception: {ex.Message})");
                    }
                }
                // Enumerate all storages (for completeness, using reflection)
                var storageNames = new System.Collections.Generic.List<string>();
                var getStorageNamesProp = rootType.GetProperty("StorageNames");
                if (getStorageNamesProp != null)
                {
                    var names = getStorageNamesProp.GetValue(cf.RootStorage) as System.Collections.IEnumerable;
                    if (names != null)
                    {
                        foreach (var n in names)
                            storageNames.Add(n.ToString());
                    }
                }
                else
                {
                    var getStorageNamesMethod = rootType.GetMethod("GetStorageNames");
                    if (getStorageNamesMethod != null)
                    {
                        var names = getStorageNamesMethod.Invoke(cf.RootStorage, null) as System.Collections.IEnumerable;
                        if (names != null)
                        {
                            foreach (var n in names)
                                storageNames.Add(n.ToString());
                        }
                    }
                }
                foreach (var storageName in storageNames)
                {
                    Console.WriteLine($"[Storage] {storageName}");
                }
                // Try 'Package' first, then any other stream with significant data
                CFStream foundStream = null;
                string foundStreamName = null;
                try
                {
                    foundStream = cf.RootStorage.GetStream("Package");
                    foundStreamName = "Package";
                }
                catch { }
                if (foundStream == null)
                {
                    // Try any other stream with data
                    foreach (var streamName in streamNames)
                    {
                        try
                        {
                            var stream = cf.RootStorage.GetStream(streamName);
                            if (stream.Size > 128) // Arbitrary threshold to skip tiny streams
                            {
                                foundStream = stream;
                                foundStreamName = streamName;
                                break;
                            }
                        }
                        catch { }
                    }
                }
                if (foundStream == null)
                {
                    Console.WriteLine("[DEBUG] No plausible stream found in OLE object.");
                    return null;
                }
                var data = foundStream.GetData();
                Console.WriteLine($"[DEBUG] '{foundStreamName}' stream size: {data.Length}");
                Console.WriteLine($"[DEBUG] foundStreamName == 'Ole10Native': {foundStreamName == "Ole10Native"}");
                Console.WriteLine($"[DEBUG] foundStreamName bytes: {string.Join(" ", System.Text.Encoding.UTF8.GetBytes(foundStreamName ?? "").Select(b => b.ToString("X2")))}");
                
                // Try to parse based on stream name
                if (foundStreamName == "Ole10Native" || foundStreamName == "\u0001Ole10Native")
                {
                    Console.WriteLine("[DEBUG] Detected Ole10Native stream, using specialized parser");
                    // Ole10Native format: different from Package format
                    try
                    {
                        return ParseOle10Native(data);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[DEBUG] Ole10Native parsing failed: {ex.Message}");
                        string fallbackName = "Ole10Native_raw.bin";
                        return new OlePackageInfo { FileName = fallbackName, ContentType = "application/octet-stream", Data = data };
                    }
                }
                else
                {
                    Console.WriteLine($"[DEBUG] Stream '{foundStreamName}' is not Ole10Native, using standard OLE Package parser");
                    // Try to parse as standard OLE Package
                    try
                    {
                        using (var br = new BinaryReader(new MemoryStream(data)))
                        {
                            br.ReadUInt32(); // Unknown, usually 2 or 3
                            int nameLen = br.ReadInt32();
                            string fileName = System.Text.Encoding.Unicode.GetString(br.ReadBytes(nameLen)).TrimEnd('\0');
                            int pathLen = br.ReadInt32();
                            string filePath = System.Text.Encoding.Unicode.GetString(br.ReadBytes(pathLen)).TrimEnd('\0');
                            int tempLen = br.ReadInt32();
                            br.ReadBytes(tempLen); // temp path
                            int dataLen = br.ReadInt32();
                            byte[] fileData = br.ReadBytes(dataLen);
                            string ext = Path.GetExtension(fileName).ToLowerInvariant();
                            string contentType = ext == ".pdf" ? "application/pdf" : "application/octet-stream";
                            Console.WriteLine($"[DEBUG] OLE extracted file: {fileName}, size: {fileData.Length}, contentType: {contentType}");
                            return new OlePackageInfo { FileName = fileName, ContentType = contentType, Data = fileData };
                        }
                    }
                    catch (Exception ex)
                    {
                        // Not a standard OLE Package, just dump the stream as a file
                        Console.WriteLine($"[DEBUG] Stream '{foundStreamName}' is not a standard OLE Package: {ex.Message}");
                        string fallbackName = foundStreamName + ".bin";
                        return new OlePackageInfo { FileName = fallbackName, ContentType = "application/octet-stream", Data = data };
                    }
                }
            }
        }

        /// <summary>
        /// Parses Ole10Native format, which is different from standard OLE Package format
        /// Ole10Native format: [4 bytes size][4 bytes type][filename\0][filepath\0][4 bytes size][data]
        /// </summary>
        private static OlePackageInfo ParseOle10Native(byte[] data)
        {
            Console.WriteLine($"[DEBUG] ParseOle10Native: data length = {data.Length}");
            
            // Dump first 100 bytes as hex for debugging
            var hexDump = string.Join(" ", data.Take(Math.Min(100, data.Length)).Select(b => b.ToString("X2")));
            Console.WriteLine($"[DEBUG] First 100 bytes: {hexDump}");
            
            using (var br = new BinaryReader(new MemoryStream(data)))
            {
                try
                {
                    // Ole10Native format has several variants, try different approaches
                    
                    // Approach 1: Standard Ole10Native format
                    uint totalSize = br.ReadUInt32();
                    Console.WriteLine($"[DEBUG] Ole10Native total size: {totalSize}");
                    
                    // Read type/format field
                    uint typeField = br.ReadUInt32();
                    Console.WriteLine($"[DEBUG] Ole10Native type field: {typeField}");
                    
                    // Read filename (null-terminated string)
                    string fileName = ReadNullTerminatedString(br);
                    Console.WriteLine($"[DEBUG] Ole10Native filename: '{fileName}'");
                    
                    // Read file path (null-terminated string)
                    string filePath = ReadNullTerminatedString(br);
                    Console.WriteLine($"[DEBUG] Ole10Native filepath: '{filePath}'");
                    
                    // Skip any additional fields (some Ole10Native variants have extra data)
                    // Look for the next size field that indicates the actual data
                    long currentPos = br.BaseStream.Position;
                    Console.WriteLine($"[DEBUG] Current position after strings: {currentPos}");
                    
                    // Try to find the data size field - it should be near the end before the actual data
                    // Sometimes there are additional null bytes or other fields
                    uint dataSize = 0;
                    bool foundDataSize = false;
                    
                    // Save current position to try multiple approaches
                    long searchStartPos = br.BaseStream.Position;
                    
                    // Approach 1: Look for the most common Ole10Native pattern
                    // After filename and filepath, there might be some padding, then a size field
                    
                    // The Ole10Native format typically has these fields after the strings:
                    // [4 bytes - temporary path length] [temporary path] [4 bytes - data size] [actual data]
                    // Let's try to find the actual data by looking for file signatures
                    
                    // First, try the traditional approach with size fields
                    for (int skip = 0; skip <= 32 && searchStartPos + skip + 4 < br.BaseStream.Length - 4; skip += 4)
                    {
                        br.BaseStream.Position = searchStartPos + skip;
                        uint candidate = br.ReadUInt32();
                        long remaining = br.BaseStream.Length - br.BaseStream.Position;
                        
                        // Check if this looks like a valid data size
                        if (candidate > 0 && candidate <= remaining)
                        {
                            // Look ahead to see if the data at this position looks like a valid file
                            long checkPos = br.BaseStream.Position;
                            if (checkPos + candidate <= br.BaseStream.Length)
                            {
                                byte[] testData = br.ReadBytes(Math.Min(16, (int)candidate));
                                br.BaseStream.Position = checkPos; // Reset position
                                
                                // Check if it looks like a valid file (PDF starts with %PDF, MSG with OLE signature)
                                bool looksValid = ValidateFileSignature(testData, fileName);
                                
                                if (looksValid && candidate >= 100 && candidate <= 104857600)
                                {
                                    dataSize = candidate;
                                    foundDataSize = true;
                                    Console.WriteLine($"[DEBUG] Found validated data size: {dataSize} at skip {skip}");
                                    break;
                                }
                            }
                        }
                    }
                    
                    // Approach 2: If traditional approach fails, scan for file signatures directly
                    if (!foundDataSize)
                    {
                        Console.WriteLine("[DEBUG] Traditional approach failed, scanning for file signatures...");
                        br.BaseStream.Position = searchStartPos;
                        
                        // For MSG files, scan more extensively since they can be deeply embedded
                        int maxScanDistance = fileName?.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) == true ? 1000 : 200;
                        
                        // Scan up to maxScanDistance bytes ahead looking for file signatures
                        for (int offset = 0; offset < Math.Min(maxScanDistance, br.BaseStream.Length - searchStartPos - 16); offset++)
                        {
                            br.BaseStream.Position = searchStartPos + offset;
                            byte[] testBytes = br.ReadBytes(16);
                            
                            if (ValidateFileSignature(testBytes, fileName))
                            {
                                Console.WriteLine($"[DEBUG] Found file signature at offset {offset} from current position");
                                dataSize = (uint)(br.BaseStream.Length - (searchStartPos + offset));
                                foundDataSize = true;
                                br.BaseStream.Position = searchStartPos + offset; // Position at start of data
                                break;
                            }
                        }
                    }
                    
                    // Approach 3: For MSG files specifically, try alternative parsing methods
                    if (!foundDataSize && fileName?.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        Console.WriteLine("[DEBUG] Trying MSG-specific parsing approaches...");
                        br.BaseStream.Position = searchStartPos;
                        
                        // Sometimes MSG files are wrapped in additional layers
                        // Try reading different size fields and looking for OLE signatures
                        for (int attemptOffset = 0; attemptOffset < Math.Min(500, br.BaseStream.Length - searchStartPos - 20); attemptOffset += 4)
                        {
                            br.BaseStream.Position = searchStartPos + attemptOffset;
                            
                            // Try reading as if there's a size field here
                            if (br.BaseStream.Position + 4 < br.BaseStream.Length)
                            {
                                uint candidateSize = br.ReadUInt32();
                                long remainingAfterSize = br.BaseStream.Length - br.BaseStream.Position;
                                
                                // Check if this size makes sense and if the data after it looks like OLE
                                if (candidateSize > 100 && candidateSize <= remainingAfterSize && candidateSize <= 104857600)
                                {
                                    byte[] testData = br.ReadBytes(Math.Min(8, (int)candidateSize));
                                    br.BaseStream.Position -= testData.Length; // Reset
                                    
                                    // Check for OLE signature specifically for MSG
                                    if (testData.Length >= 8 && 
                                        testData[0] == 0xD0 && testData[1] == 0xCF && testData[2] == 0x11 && testData[3] == 0xE0 &&
                                        testData[4] == 0xA1 && testData[5] == 0xB1 && testData[6] == 0x1A && testData[7] == 0xE1)
                                    {
                                        Console.WriteLine($"[DEBUG] Found MSG OLE signature at offset {attemptOffset} with size {candidateSize}");
                                        dataSize = candidateSize;
                                        foundDataSize = true;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    
                    if (!foundDataSize)
                    {
                        Console.WriteLine("[DEBUG] Could not find data size field, using remaining data");
                        // Fallback: use all remaining data after trying to skip potential headers
                        br.BaseStream.Position = searchStartPos;
                        
                        // Skip what looks like it might be padding or additional fields
                        long dataStartPos = searchStartPos;
                        if (br.BaseStream.Length - searchStartPos > 12)
                        {
                            // Try skipping 4, 8, or 12 bytes to account for potential fields
                            dataStartPos = searchStartPos + 4;
                        }
                        
                        br.BaseStream.Position = dataStartPos;
                        dataSize = (uint)(br.BaseStream.Length - br.BaseStream.Position);
                        Console.WriteLine($"[DEBUG] Using fallback data size: {dataSize} from position {dataStartPos}");
                    }
                    
                    // Read the actual file data
                    byte[] fileData = br.ReadBytes((int)dataSize);
                    Console.WriteLine($"[DEBUG] Read {fileData.Length} bytes of file data");
                    
                    // Validate the data by checking for known file signatures
                    bool isValidData = ValidateFileData(fileData, fileName);
                    Console.WriteLine($"[DEBUG] Initial validation result: {isValidData}");
                    
                    if (!isValidData && foundDataSize)
                    {
                        Console.WriteLine("[DEBUG] Data validation failed, trying fallback approach");
                        // Try using all remaining data instead
                        br.BaseStream.Position = searchStartPos;
                        fileData = br.ReadBytes((int)(br.BaseStream.Length - br.BaseStream.Position));
                        Console.WriteLine($"[DEBUG] Fallback: Read {fileData.Length} bytes");
                        isValidData = ValidateFileData(fileData, fileName);
                        Console.WriteLine($"[DEBUG] Fallback validation result: {isValidData}");
                    }
                    
                    // For MSG files specifically, if still not valid, try additional fallback methods
                    if (!isValidData && fileName?.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        Console.WriteLine("[DEBUG] MSG file still not valid, trying progressive fallbacks");
                        
                        // Try different starting positions by skipping potential wrapper headers
                        int[] skipOffsets = { 8, 12, 16, 20, 24, 32, 64, 128 };
                        foreach (int skipOffset in skipOffsets)
                        {
                            if (searchStartPos + skipOffset < br.BaseStream.Length - 100)
                            {
                                br.BaseStream.Position = searchStartPos + skipOffset;
                                byte[] testData = br.ReadBytes(8);
                                
                                // Check for OLE signature at this position
                                if (testData.Length >= 8 && 
                                    testData[0] == 0xD0 && testData[1] == 0xCF && testData[2] == 0x11 && testData[3] == 0xE0 &&
                                    testData[4] == 0xA1 && testData[5] == 0xB1 && testData[6] == 0x1A && testData[7] == 0xE1)
                                {
                                    Console.WriteLine($"[DEBUG] Found MSG OLE signature at skip offset {skipOffset}");
                                    br.BaseStream.Position = searchStartPos + skipOffset;
                                    fileData = br.ReadBytes((int)(br.BaseStream.Length - br.BaseStream.Position));
                                    isValidData = true;
                                    break;
                                }
                            }
                        }
                        
                        // If still not valid, save the raw data anyway with detailed logging for analysis
                        if (!isValidData)
                        {
                            Console.WriteLine($"[DEBUG] MSG file validation failed completely. Raw data length: {fileData.Length}");
                            if (fileData.Length > 0)
                            {
                                var first32 = string.Join(" ", fileData.Take(32).Select(b => b.ToString("X2")));
                                Console.WriteLine($"[DEBUG] First 32 bytes of invalid MSG data: {first32}");
                                
                                // Return the data anyway - it might be a variant that can still be processed
                                // We'll add a marker to the filename to indicate it needs validation
                                fileName = "UNVALIDATED_" + fileName;
                                Console.WriteLine($"[DEBUG] Returning unvalidated MSG data as: {fileName}");
                            }
                        }
                    }
                    
                    // Clean up filename - remove invalid characters
                    if (string.IsNullOrEmpty(fileName))
                    {
                        fileName = "Ole10Native_file.bin";
                    }
                    else
                    {
                        // Remove invalid path characters
                        fileName = string.Join("_", fileName.Split(Path.GetInvalidFileNameChars()));
                        if (string.IsNullOrEmpty(fileName))
                        {
                            fileName = "Ole10Native_file.bin";
                        }
                    }
                    
                    string ext = Path.GetExtension(fileName).ToLowerInvariant();
                    string contentType = ext == ".pdf" ? "application/pdf" : 
                                        ext == ".msg" ? "application/vnd.ms-outlook" :
                                        "application/octet-stream";
                    
                    Console.WriteLine($"[DEBUG] Ole10Native extracted file: {fileName}, size: {fileData.Length}, contentType: {contentType}, validated: {isValidData}");
                    return new OlePackageInfo { FileName = fileName, ContentType = contentType, Data = fileData };
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"[DEBUG] Ole10Native parsing failed: {ex.Message}");
                    // Last resort: just return the raw data
                    string fallbackName = "Ole10Native_raw.bin";
                    return new OlePackageInfo { FileName = fallbackName, ContentType = "application/octet-stream", Data = data };
                }
            }
        }
        
        private static string ReadNullTerminatedString(BinaryReader br)
        {
            var bytes = new List<byte>();
            byte b;
            while (br.BaseStream.Position < br.BaseStream.Length && (b = br.ReadByte()) != 0)
            {
                bytes.Add(b);
            }
            return System.Text.Encoding.ASCII.GetString(bytes.ToArray());
        }

        /// <summary>
        /// Validates if the extracted data looks like a valid file based on magic bytes
        /// </summary>
        private static bool ValidateFileData(byte[] data, string fileName)
        {
            if (data == null || data.Length < 4)
                return false;
                
            return ValidateFileSignature(data, fileName);
        }

        /// <summary>
        /// Validates file signatures (magic bytes) for common file types
        /// </summary>
        private static bool ValidateFileSignature(byte[] data, string fileName)
        {
            if (data == null || data.Length < 4)
                return false;
                
            // Check common file signatures
            string ext = Path.GetExtension(fileName).ToLowerInvariant();
            
            // PDF files start with %PDF
            if (ext == ".pdf" && data.Length >= 4)
            {
                bool isPdf = data[0] == 0x25 && data[1] == 0x50 && data[2] == 0x44 && data[3] == 0x46; // %PDF
                if (isPdf) Console.WriteLine($"[DEBUG] PDF signature check: {isPdf}");
                return isPdf;
            }
            
            // MSG files (Outlook messages) - OLE compound document signature
            if (ext == ".msg" && data.Length >= 8)
            {
                bool isMsg = data[0] == 0xD0 && data[1] == 0xCF && data[2] == 0x11 && data[3] == 0xE0 &&
                            data[4] == 0xA1 && data[5] == 0xB1 && data[6] == 0x1A && data[7] == 0xE1;
                Console.WriteLine($"[DEBUG] MSG signature check: {isMsg} (first 8 bytes: {string.Join(" ", data.Take(8).Select(b => b.ToString("X2")))})");
                return isMsg;
            }
            
            // For unknown file types, do basic sanity checks
            // Check if it's not all zeros or all same byte
            byte firstByte = data[0];
            bool allSame = data.Take(Math.Min(100, data.Length)).All(b => b == firstByte);
            if (allSame && firstByte == 0)
            {
                Console.WriteLine($"[DEBUG] File signature check failed: all zeros");
                return false;
            }
            
            return true;
        }

        private static void DumpStorage(CFStorage storage, string indent)
        {
            // Only try to get the 'Package' stream, no enumeration
            try
            {
                var stream = storage.GetStream("Package");
                if (stream != null)
                    Console.WriteLine($"{indent}[Stream] Package");
                else
                    Console.WriteLine($"{indent}[Stream] Package not found");
            }
            catch
            {
                Console.WriteLine($"{indent}[Stream] Package not found (exception)");
            }
        }
    }
}
