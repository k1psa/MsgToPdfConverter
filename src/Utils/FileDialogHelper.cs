using Microsoft.Win32;
using System.Collections.Generic;

namespace MsgToPdfConverter.Utils
{
    public static class FileDialogHelper
    {
        public static List<string> OpenMsgFileDialog()
        {
            var openFileDialog = new OpenFileDialog
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
    }
}