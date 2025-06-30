using System;
using System.Windows;

namespace MsgToPdfConverter
{
    public partial class MainWindow : Window
    {
        private MainWindowViewModel _viewModel;
        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainWindowViewModel();
            this.DataContext = _viewModel;
        }

        // Drag-and-drop event handlers delegate to ViewModel
        private void FilesListBox_Drop(object sender, DragEventArgs e)
        {
            Console.WriteLine("FilesListBox_Drop event triggered");
            _viewModel.HandleDrop(e.Data);
        }
        private void FilesListBox_DragEnter(object sender, DragEventArgs e)
        {
            Console.WriteLine("FilesListBox_DragEnter event triggered");
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
        private void FilesListBox_DragOver(object sender, DragEventArgs e)
        {
            Console.WriteLine("FilesListBox_DragOver event triggered");
            FilesListBox_DragEnter(sender, e);
        }
        private void FilesListBox_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            var listBox = sender as System.Windows.Controls.ListBox;
            if (e.Key == System.Windows.Input.Key.Delete && listBox != null && listBox.SelectedItems.Count > 0)
            {
                Console.WriteLine($"FilesListBox_KeyDown: Deleting {listBox.SelectedItems.Count} items");
                var items = new System.Collections.Generic.List<string>();
                foreach (var item in listBox.SelectedItems)
                {
                    if (item is string s)
                        items.Add(s);
                }
                if (_viewModel.RemoveSelectedFilesCommand.CanExecute(items))
                    _viewModel.RemoveSelectedFilesCommand.Execute(items);
                e.Handled = true;
            }
        }
    }
}