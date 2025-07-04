using System;
using System.Windows;
using System.Windows.Media;

namespace MsgToPdfConverter
{
    public partial class MainWindow : Window
    {
        private MainWindowViewModel _viewModel;
        private System.Windows.Forms.NotifyIcon _trayIcon;

        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainWindowViewModel();
            this.DataContext = _viewModel;

            // Initialize tray icon
            _trayIcon = new System.Windows.Forms.NotifyIcon();
            try
            {
                string iconPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "icon.ico");
                if (System.IO.File.Exists(iconPath))
                {
                    _trayIcon.Icon = new System.Drawing.Icon(iconPath);
                }
                else
                {
                    // Fallback to default if icon.ico is missing
                    _trayIcon.Icon = System.Drawing.SystemIcons.Application;
                    Console.WriteLine($"[DEBUG] icon.ico not found at {iconPath}, using default icon.");
                }
            }
            catch (Exception ex)
            {
                _trayIcon.Icon = System.Drawing.SystemIcons.Application;
                Console.WriteLine($"[DEBUG] Failed to load icon.ico: {ex.Message}");
            }
            _trayIcon.Visible = false;
            _trayIcon.DoubleClick += TrayIcon_DoubleClick;

            // Add context menu to tray icon
            var contextMenu = new System.Windows.Forms.ContextMenuStrip();
            var restoreItem = new System.Windows.Forms.ToolStripMenuItem("Restore Window");
            restoreItem.Click += (s, e) => RestoreFromTray();
            var exitItem = new System.Windows.Forms.ToolStripMenuItem("Exit");
            exitItem.Click += (s, e) => ExitFromTray();
            var resetItem = new System.Windows.Forms.ToolStripMenuItem("Reset Close Behavior");
            resetItem.Click += (s, e) => ResetCloseBehaviorFromTray();
            contextMenu.Items.Add(restoreItem);
            contextMenu.Items.Add(resetItem);
            contextMenu.Items.Add(exitItem);
            _trayIcon.ContextMenuStrip = contextMenu;
        }

        // Drag-and-drop event handlers delegate to ViewModel
        private void FilesListBox_Drop(object sender, DragEventArgs e)
        {
            var listBox = sender as System.Windows.Controls.ListBox;
            var droppedData = e.Data.GetData(typeof(string)) as string;
            var target = GetObjectDataFromPoint(listBox, e.GetPosition(listBox)) as string;
            // Only reorder if dropped on another item
            if (droppedData != null && target != null && droppedData != target)
            {
                int oldIndex = listBox.Items.IndexOf(droppedData);
                int newIndex = listBox.Items.IndexOf(target);
                _viewModel.MoveFile(oldIndex, newIndex);
                listBox.SelectedItem = droppedData;
            }
            else if (droppedData != null && target == null)
            {
                // Dropped in empty space: move to end
                int oldIndex = listBox.Items.IndexOf(droppedData);
                int newIndex = listBox.Items.Count - 1;
                if (oldIndex != newIndex)
                {
                    _viewModel.MoveFile(oldIndex, newIndex);
                    listBox.SelectedItem = droppedData;
                }
            }
            else
            {
                // Fallback to original drop logic for files/folders, but suppress irrelevant error for internal reordering
                if (!(droppedData is string))
                {
                    _viewModel.HandleDrop(e.Data);
                }
            }
        }

        private object GetObjectDataFromPoint(System.Windows.Controls.ListBox source, Point point)
        {
            var element = source.InputHitTest(point) as UIElement;
            while (element != null)
            {
                if (element is System.Windows.Controls.ListBoxItem)
                {
                    return ((System.Windows.Controls.ListBoxItem)element).DataContext;
                }
                element = VisualTreeHelper.GetParent(element) as UIElement;
            }
            return null;
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

        private void TrayIcon_DoubleClick(object sender, EventArgs e)
        {
            this.Show();
            this.WindowState = WindowState.Normal;
            _trayIcon.Visible = false;
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            string behavior = Properties.Settings.Default.CloseButtonBehavior ?? "Ask";
            if (behavior == "Minimize")
            {
                e.Cancel = true;
                this.Hide();
                _trayIcon.Visible = true;
                return;
            }
            else if (behavior == "Ask")
            {
                var result = MessageBox.Show("Do you want to minimize to tray instead of exiting?", "Close", MessageBoxButton.YesNoCancel);
                if (result == MessageBoxResult.Yes)
                {
                    Properties.Settings.Default.CloseButtonBehavior = "Minimize";
                    Properties.Settings.Default.Save();
                    e.Cancel = true;
                    this.Hide();
                    _trayIcon.Visible = true;
                    return;
                }
                else if (result == MessageBoxResult.No)
                {
                    Properties.Settings.Default.CloseButtonBehavior = "Exit";
                    Properties.Settings.Default.Save();
                    // Allow exit
                }
                else
                {
                    e.Cancel = true;
                    return;
                }
            }
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            base.OnClosing(e);
        }

        private void RestoreFromTray()
        {
            this.Dispatcher.Invoke(() =>
            {
                this.Show();
                this.WindowState = WindowState.Normal;
                _trayIcon.Visible = false;
            });
        }

        private void ExitFromTray()
        {
            this.Dispatcher.Invoke(() =>
            {
                _trayIcon.Visible = false;
                _trayIcon.Dispose();
                this.Close();
            });
        }

        private void ResetCloseBehaviorFromTray()
        {
            this.Dispatcher.Invoke(() =>
            {
                Properties.Settings.Default.CloseButtonBehavior = "Ask";
                Properties.Settings.Default.Save();
                System.Windows.MessageBox.Show("Close button behavior has been reset. You will be prompted next time you close the window.", "Reset", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            });
        }

        private void OpenOptionsWindow()
        {
            // Load current settings
            bool deleteMsg = Properties.Settings.Default["DeleteMsgAfterConversion"] is bool d ? d : false;
            string closeBehavior = Properties.Settings.Default.CloseButtonBehavior ?? "Ask";
            var options = new OptionsWindow(deleteMsg, closeBehavior);
            options.Owner = this;
            if (options.ShowDialog() == true)
            {
                Properties.Settings.Default["DeleteMsgAfterConversion"] = options.DeleteMsgAfterConversion;
                Properties.Settings.Default.CloseButtonBehavior = options.CloseButtonBehavior;
                Properties.Settings.Default.Save();
            }
        }

        // For drag-and-drop reordering
        private Point _dragStartPoint;
        private void FilesListBox_PreviewMouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(null);
        }

        private void FilesListBox_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
            {
                var pos = e.GetPosition(null);
                if (Math.Abs(pos.X - _dragStartPoint.X) > SystemParameters.MinimumHorizontalDragDistance ||
                    Math.Abs(pos.Y - _dragStartPoint.Y) > SystemParameters.MinimumVerticalDragDistance)
                {
                    var listBox = sender as System.Windows.Controls.ListBox;
                    if (listBox?.SelectedItem == null) return;
                    DragDrop.DoDragDrop(listBox, listBox.SelectedItem, DragDropEffects.Move);
                }
            }
        }
    }
}