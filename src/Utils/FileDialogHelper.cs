using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System;

namespace MsgToPdfConverter.Utils
{
    public static class FileDialogHelper
    {
        public static List<string> OpenMsgFileDialog()
        {
            var result = new List<string>();

            using (var dialog = new System.Windows.Forms.OpenFileDialog())
            {
                dialog.Title = "Select .msg Files or Folders (type folder path in filename)";
                dialog.Filter = "Outlook Message Files (*.msg)|*.msg|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.Multiselect = true;
                dialog.CheckFileExists = false;
                dialog.CheckPathExists = false;
                dialog.ValidateNames = false;
                dialog.DereferenceLinks = false;
                dialog.FileName = "Folder Selection";

                var dialogResult = dialog.ShowDialog();
                if (dialogResult == DialogResult.OK)
                {
                    foreach (string selectedPath in dialog.FileNames)
                    {
                        // Clean the path (remove the dummy filename if present)
                        string cleanPath = selectedPath;
                        if (selectedPath.EndsWith("\\Folder Selection") || selectedPath.EndsWith("/Folder Selection"))
                        {
                            cleanPath = Path.GetDirectoryName(selectedPath);
                        }

                        if (Directory.Exists(cleanPath))
                        {
                            // It's a directory - add all .msg files recursively
                            var msgFiles = Directory.GetFiles(cleanPath, "*.msg", SearchOption.AllDirectories);
                            result.AddRange(msgFiles);
                        }
                        else if (File.Exists(selectedPath))
                        {
                            // It's a file
                            if (Path.GetExtension(selectedPath).ToLowerInvariant() == ".msg")
                            {
                                result.Add(selectedPath);
                            }
                        }
                        else
                        {
                            // Try parent directory
                            string folderPath = Path.GetDirectoryName(selectedPath);
                            if (Directory.Exists(folderPath))
                            {
                                var msgFiles = Directory.GetFiles(folderPath, "*.msg", SearchOption.AllDirectories);
                                result.AddRange(msgFiles);
                            }
                        }
                    }
                }
            }

            return result;
        }
    }
}