using System.Windows;
using System.Windows.Controls;

namespace MsgToPdfConverter
{
    public partial class OptionsWindow : Window
    {
        public bool DeleteMsgAfterConversion { get; private set; }
        public string CloseButtonBehavior { get; private set; }

        public OptionsWindow(bool deleteMsgAfterConversion, string closeButtonBehavior)
        {
            InitializeComponent();
            DeleteMsgAfterConversionCheckBox.IsChecked = deleteMsgAfterConversion;
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
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DeleteMsgAfterConversion = DeleteMsgAfterConversionCheckBox.IsChecked == true;
            Properties.Settings.Default.DeleteMsgAfterConversion = DeleteMsgAfterConversion;

            string selected = ((ComboBoxItem)CloseBehaviorComboBox.SelectedItem).Content.ToString();
            CloseButtonBehavior = selected == "Minimize to tray" ? "Minimize" : selected;
            Properties.Settings.Default.CloseButtonBehavior = CloseButtonBehavior;

            Properties.Settings.Default.Save();
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
