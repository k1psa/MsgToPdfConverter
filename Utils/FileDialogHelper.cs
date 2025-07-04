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

        public static string OpenFolderDialog()
        {
            using (var dialog = new System.Windows.Forms.OpenFileDialog())
            {
                dialog.Title = "Select Output Folder";
                dialog.Filter = "Folders|*.*";
                dialog.CheckFileExists = false;
                dialog.CheckPathExists = false;
                dialog.ValidateNames = false;
                dialog.DereferenceLinks = false;
                dialog.FileName = "Select this folder";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = dialog.FileName;
                    if (selectedPath.EndsWith("Select this folder") || selectedPath.EndsWith("Select this folder".Replace(" ", "")))
                    {
                        return System.IO.Path.GetDirectoryName(selectedPath);
                    }
                    else if (System.IO.Directory.Exists(selectedPath))
                    {
                        return selectedPath;
                    }
                    else if (System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(selectedPath)))
                    {
                        return System.IO.Path.GetDirectoryName(selectedPath);
                    }
                }
            }
            return null;
        }

        public static string SavePdfFileDialog(string defaultFileName = "Binder1.pdf")
        {
            using (var dialog = new System.Windows.Forms.SaveFileDialog())
            {
                dialog.Title = "Save Combined PDF As";
                dialog.Filter = "PDF Files (*.pdf)|*.pdf|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.FileName = defaultFileName;
                dialog.OverwritePrompt = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.FileName;
                }
            }
            return null;
        }
    }
}