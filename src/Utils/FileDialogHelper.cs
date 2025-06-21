using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace MsgToPdfConverter.Utils
{
    public static class FileDialogHelper
    {
        public static List<string> OpenMsgFileDialog()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select .msg Files",
                Filter = "Outlook Message Files (*.msg)|*.msg",
                Multiselect = true
            };

            var result = new List<string>();

            if (openFileDialog.ShowDialog() == true)
            {
                result.AddRange(openFileDialog.FileNames);
            }

            return result;
        }        public static List<string> OpenMsgFolderDialog()
        {
            var result = new List<string>();

            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Select Folder Containing .msg Files";
                folderDialog.ShowNewFolderButton = false;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = folderDialog.SelectedPath;
                    if (!string.IsNullOrWhiteSpace(folderPath) && Directory.Exists(folderPath))
                    {
                        result.AddRange(Directory.GetFiles(folderPath, "*.msg", SearchOption.AllDirectories));
                    }
                }
            }

            return result;
        }
    }
}