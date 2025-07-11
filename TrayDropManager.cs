using System;
using System.Windows;
using System.Windows.Forms;
using DrawingPoint = System.Drawing.Point;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace MsgToPdfConverter
{
    public class TrayDropManager
    {
        private NotifyIcon _trayIcon;
        private TrayDropWindow _dropWindow;
        private MainWindowViewModel _viewModel;
        private bool _dropWindowVisible = false;
        private DispatcherTimer _topmostTimer;

        public TrayDropManager(NotifyIcon trayIcon, MainWindowViewModel viewModel)
        {
            Console.WriteLine("[DEBUG] TrayDropManager constructor");
            _trayIcon = trayIcon;
            _viewModel = viewModel;
            _dropWindow = new TrayDropWindow();
            _dropWindow.DataDropped += OnDataDropped;
            _topmostTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            _topmostTimer.Tick += (s, e) => {
                if (_dropWindowVisible)
                {
                    _dropWindow.Topmost = true;
                    _dropWindow.Activate();
                }
            };
        }

        public void Enable()
        {
            Console.WriteLine("[DEBUG] TrayDropManager.Enable() called");
            _trayIcon.MouseClick += TrayIcon_MouseClick;
        }

        private void TrayIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                if (!_dropWindowVisible)
                {
                    // Position the drop window just above the taskbar, right-aligned
                    var screen = System.Windows.Forms.Screen.PrimaryScreen;
                    int margin = 20;
                    _dropWindow.Left = screen.WorkingArea.Right - _dropWindow.Width - margin;
                    _dropWindow.Top = screen.WorkingArea.Bottom - _dropWindow.Height - margin;
                    _dropWindow.Topmost = true;
                    _dropWindow.Show();
                    _dropWindow.Activate();
                    _dropWindow.Focus();
                    _dropWindowVisible = true;
                    _topmostTimer.Start();
                }
                else
                {
                    _dropWindow.Hide();
                    _dropWindowVisible = false;
                    _topmostTimer.Stop();
                    Console.WriteLine("[DEBUG] Hiding drop window");
                }
            }
        }

        private void OnDataDropped(System.Windows.IDataObject data)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                _viewModel.HandleDrop(data);
            });
        }
    }
}
