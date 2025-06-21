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
                dialog.Title = "Select .msg Files or navigate to Folders and click Add";
                dialog.Filter = "Outlook Message Files (*.msg)|*.msg|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.Multiselect = true;
                dialog.CheckFileExists = false;
                dialog.CheckPathExists = true;
                dialog.ValidateNames = false;
                dialog.DereferenceLinks = false;

                var dialogResult = dialog.ShowDialog();
                if (dialogResult == DialogResult.OK)
                {
                    foreach (string selectedPath in dialog.FileNames)
                    {
                        if (File.Exists(selectedPath))
                        {
                            // It's a file
                            if (Path.GetExtension(selectedPath).ToLowerInvariant() == ".msg")
                            {
                                result.Add(selectedPath);
                            }
                        }
                        else
                        {
                            // Try to treat it as a folder path
                            string folderPath = Path.GetDirectoryName(selectedPath);
                            if (Directory.Exists(folderPath))
                            {
                                var msgFiles = Directory.GetFiles(folderPath, "*.msg", SearchOption.AllDirectories);
                                result.AddRange(msgFiles);
                            }
                            else if (Directory.Exists(selectedPath))
                            {
                                var msgFiles = Directory.GetFiles(selectedPath, "*.msg", SearchOption.AllDirectories);
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