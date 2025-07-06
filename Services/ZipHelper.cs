using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace MsgToPdfConverter.Services
{
    public static class ZipHelper
    {
        public class ZipEntryInfo
        {
            public string FileName { get; set; }
            public byte[] Data { get; set; }
        }

        public static List<ZipEntryInfo> ExtractZipEntries(string zipFilePath)
        {
            var entries = new List<ZipEntryInfo>();
            using (var zip = ZipFile.OpenRead(zipFilePath))
            {
                foreach (var entry in zip.Entries)
                {
                    if (string.IsNullOrEmpty(entry.Name)) continue; // skip folders
                    using (var ms = new MemoryStream())
                    using (var entryStream = entry.Open())
                    {
                        entryStream.CopyTo(ms);
                        entries.Add(new ZipEntryInfo { FileName = entry.Name, Data = ms.ToArray() });
                    }
                }
            }
            return entries;
        }
    }
}
