using System.Windows;

namespace MsgToPdfConverter
{
    public partial class OptionsWindow : Window
    {
        public bool ExtractOriginalOnly { get; set; }
        public bool DeleteMsgAfterConversion { get; set; }

        public OptionsWindow(bool extractOriginalOnly, bool deleteMsgAfterConversion)
        {
            InitializeComponent();
            ExtractOriginalOnlyCheckBox.IsChecked = extractOriginalOnly;
            DeleteMsgAfterConversionCheckBox.IsChecked = deleteMsgAfterConversion;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ExtractOriginalOnly = ExtractOriginalOnlyCheckBox.IsChecked == true;
            DeleteMsgAfterConversion = DeleteMsgAfterConversionCheckBox.IsChecked == true;
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
