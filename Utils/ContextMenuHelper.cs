using Microsoft.Win32;
using System;
using System.IO;

namespace MsgToPdfConverter.Utils
{
    public static class ContextMenuHelper
    {
        private const string MenuKeyPath = @"Software\Classes\*\shell\AddToMsgToPDF";
        private const string CommandKeyPath = MenuKeyPath + "\\command";
        private const string FolderMenuKeyPath = @"Software\Classes\Directory\shell\AddToMsgToPDF";
        private const string FolderCommandKeyPath = FolderMenuKeyPath + "\\command";
        private const string MenuText = "Add to MsgToPDF list";

        public static void SetContextMenu(bool enable)
        {
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            if (enable)
            {
                // For files
                using (var key = Registry.CurrentUser.CreateSubKey(MenuKeyPath))
                {
                    key.SetValue(null, MenuText);
                    key.SetValue("Icon", exePath);
                    if (key.GetValue("Extended") != null)
                        key.DeleteValue("Extended");
                }
                using (var key = Registry.CurrentUser.CreateSubKey(CommandKeyPath))
                {
                    key.SetValue(null, $"\"{exePath}\" \"%1\"");
                }
                // For folders
                using (var key = Registry.CurrentUser.CreateSubKey(FolderMenuKeyPath))
                {
                    key.SetValue(null, MenuText);
                    key.SetValue("Icon", exePath);
                    if (key.GetValue("Extended") != null)
                        key.DeleteValue("Extended");
                }
                using (var key = Registry.CurrentUser.CreateSubKey(FolderCommandKeyPath))
                {
                    key.SetValue(null, $"\"{exePath}\" \"%1\"");
                }
            }
            else
            {
                try { Registry.CurrentUser.DeleteSubKeyTree(MenuKeyPath, false); } catch { }
                try { Registry.CurrentUser.DeleteSubKeyTree(FolderMenuKeyPath, false); } catch { }
            }
        }

        public static bool IsContextMenuEnabled()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(MenuKeyPath))
            {
                return key != null;
            }
        }
    }
}
