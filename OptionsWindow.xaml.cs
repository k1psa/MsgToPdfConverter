using System.Windows;
using System.Windows.Controls;

namespace MsgToPdfConverter
{
    public partial class OptionsWindow : Window
    {
        public bool DeleteFilesAfterConversion { get; private set; }
        public string CloseButtonBehavior { get; private set; }

        public OptionsWindow(bool deleteFilesAfterConversion, string closeButtonBehavior)
        {
            InitializeComponent();
            DeleteFilesAfterConversionCheckBox.IsChecked = deleteFilesAfterConversion;
            switch (closeButtonBehavior)
            {
                case "Minimize to tray":
                case "Minimize":
                    CloseBehaviorComboBox.SelectedIndex = 1;
                    break;
                case "Exit":
                    CloseBehaviorComboBox.SelectedIndex = 2;
                    break;
                default:
                    CloseBehaviorComboBox.SelectedIndex = 0;
                    break;
            }
            // Set context menu checkbox from settings
            EnableContextMenuCheckBox.IsChecked = Properties.Settings.Default.EnableContextMenu;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DeleteFilesAfterConversion = DeleteFilesAfterConversionCheckBox.IsChecked == true;
            Properties.Settings.Default.DeleteMsgAfterConversion = DeleteFilesAfterConversion;

            string selected = ((ComboBoxItem)CloseBehaviorComboBox.SelectedItem).Content.ToString();
            CloseButtonBehavior = selected == "Minimize to tray" ? "Minimize" : selected;
            Properties.Settings.Default.CloseButtonBehavior = CloseButtonBehavior;

            // Save context menu setting
            bool enableContextMenu = EnableContextMenuCheckBox.IsChecked == true;
            Properties.Settings.Default.EnableContextMenu = enableContextMenu;
            Properties.Settings.Default.Save();
            // Apply context menu
            Utils.ContextMenuHelper.SetContextMenu(enableContextMenu);
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
