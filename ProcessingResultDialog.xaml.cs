using System.Windows;

namespace MsgToPdfConverter
{
    public partial class ProcessingResultDialog : Window
    {
        public ProcessingResultDialog(string message, string title = "Processing Results", bool isError = false, bool isWarning = false)
        {
            InitializeComponent();
            this.Title = title;
            MessageTextBlock.Text = message;
            if (isError)
            {
                MessageTextBlock.Foreground = System.Windows.Media.Brushes.DarkRed;
            }
            else if (isWarning)
            {
                MessageTextBlock.Foreground = System.Windows.Media.Brushes.DarkOrange;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }
    }
}
