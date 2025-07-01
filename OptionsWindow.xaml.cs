using System.Windows;

namespace MsgToPdfConverter
{
    public partial class OptionsWindow : Window
    {
        public bool DeleteMsgAfterConversion { get; set; }

        public OptionsWindow(bool deleteMsgAfterConversion)
        {
            InitializeComponent();
            DeleteMsgAfterConversionCheckBox.IsChecked = deleteMsgAfterConversion;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
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
