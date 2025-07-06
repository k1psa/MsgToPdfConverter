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
                dialog.Title = "Select Files or Folders for Conversion (type folder path in filename)";
                dialog.Filter = "All Supported Files (*.msg;*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.zip;*.7z;*.jpg;*.jpeg;*.png;*.bmp;*.gif)|*.msg;*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.zip;*.7z;*.jpg;*.jpeg;*.png;*.bmp;*.gif|All Files (*.*)|*.*";
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

                        string[] supportedExtensions = new[] { ".msg", ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip", ".7z", ".jpg", ".jpeg", ".png", ".bmp", ".gif" };

                        if (Directory.Exists(cleanPath))
                        {
                            // It's a directory - add all supported files recursively
                            foreach (var ext in supportedExtensions)
                            {
                                var files = Directory.GetFiles(cleanPath, "*" + ext, SearchOption.AllDirectories);
                                result.AddRange(files);
                            }
                        }
                        else if (File.Exists(selectedPath))
                        {
                            // It's a file
                            if (Array.Exists(supportedExtensions, e => e == Path.GetExtension(selectedPath).ToLowerInvariant()))
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
                                foreach (var ext in supportedExtensions)
                                {
                                    var files = Directory.GetFiles(folderPath, "*" + ext, SearchOption.AllDirectories);
                                    result.AddRange(files);
                                }
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