using Microsoft.Win32;
using System;
using System.IO;

namespace MsgToPdfConverter.Utils
{
    public static class ContextMenuHelper
    {
        private const string MenuKeyPath = @"Software\Classes\*\shell\AddToMsgToPDF";
        private const string CommandKeyPath = MenuKeyPath + "\\command";
        private const string MenuText = "Add to MsgToPDF list";

        public static void SetContextMenu(bool enable)
        {
            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            if (enable)
            {
                using (var key = Registry.CurrentUser.CreateSubKey(MenuKeyPath))
                {
                    key.SetValue(null, MenuText);
                }
                using (var key = Registry.CurrentUser.CreateSubKey(CommandKeyPath))
                {
                    key.SetValue(null, $"\"{exePath}\" \"%1\"");
                }
            }
            else
            {
                try { Registry.CurrentUser.DeleteSubKeyTree(MenuKeyPath, false); } catch { }
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
