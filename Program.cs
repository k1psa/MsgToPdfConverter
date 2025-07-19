using System;
using System;
using System.IO;
using System.Linq;

namespace MsgToPdfConverter
{
    public class Program
    {
        // Entry point now handled in App.xaml.cs OnStartup

        public static bool IsDotNetFrameworkInstalled()
        {
            try
            {
                // Check if .NET Framework 4.8 or higher is installed
                using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"))
                {
                    if (key != null)
                    {
                        var release = key.GetValue("Release");
                        if (release != null && (int)release >= 528040) // .NET Framework 4.8
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            catch
            {
                return false; // Assume not installed if we can't check
            }
        }

        public static bool IsOfficeInstalled()
        {
            // Check for Word and Excel registry keys
            try
            {
                using (var wordKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("Word.Application"))
                using (var excelKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey("Excel.Application"))
                {
                    return wordKey != null && excelKey != null;
                }
            }
            catch
            {
                return false;
            }
        }

        // (Removed duplicate private static IsDotNetFrameworkInstalled and IsOfficeInstalled)
    }
}
