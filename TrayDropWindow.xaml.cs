using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Data;

namespace MsgToPdfConverter
{
    public partial class TrayDropWindow : Window
    {
        public event Action<IDataObject> DataDropped;

        public TrayDropWindow()
        {
            Console.WriteLine("[DEBUG] TrayDropWindow constructor");
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += TrayDropWindow_DragEnter;
            this.DragOver += TrayDropWindow_DragOver;
            this.Drop += TrayDropWindow_Drop;
        }

        private void TrayDropWindow_DragEnter(object sender, DragEventArgs e)
        {
            Console.WriteLine($"[DEBUG] TrayDropWindow_DragEnter: Data formats: {string.Join(", ", e.Data.GetFormats())}");
            if (e.Data.GetDataPresent(DataFormats.FileDrop) ||
                e.Data.GetDataPresent("FileGroupDescriptorW") ||
                e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }

        private void TrayDropWindow_DragOver(object sender, DragEventArgs e)
        {
            Console.WriteLine($"[DEBUG] TrayDropWindow_DragOver: Data formats: {string.Join(", ", e.Data.GetFormats())}");
            if (e.Data.GetDataPresent(DataFormats.FileDrop) ||
                e.Data.GetDataPresent("FileGroupDescriptorW") ||
                e.Data.GetDataPresent("FileGroupDescriptor"))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void TrayDropWindow_Drop(object sender, DragEventArgs e)
        {
            Console.WriteLine($"[DEBUG] TrayDropWindow_Drop: Data formats: {string.Join(", ", e.Data.GetFormats())}");
            if (e.Data != null)
            {
                DataDropped?.Invoke(e.Data);
            }
            // Do not hide the window after drop; let user close it manually
        }

        public new void Show()
        {
            Console.WriteLine("[DEBUG] TrayDropWindow.Show() called");
            base.Show();
        }

        public override void OnApplyTemplate()
        {
            base.OnApplyTemplate();
            var closeButton = this.FindName("CloseButton") as System.Windows.Controls.Button;
            if (closeButton != null)
                closeButton.Click += (s, e) => this.Hide();
        }
    }
}
