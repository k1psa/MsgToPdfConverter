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
        }

        public static List<string> OpenMsgFolderDialog()
        {
            var result = new List<string>();
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "Select a folder containing .msg files";
                var dr = dialog.ShowDialog();
                if (dr == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    string folder = dialog.SelectedPath;
                    result.AddRange(Directory.GetFiles(folder, "*.msg", SearchOption.AllDirectories));
                }
            }
            return result;
        }
    }
}