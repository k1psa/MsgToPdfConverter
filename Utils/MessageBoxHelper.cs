using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Interop;

namespace MsgToPdfConverter.Utils
{
    public static class MessageBoxHelper
    {
        private delegate IntPtr HookProc(int nCode, IntPtr wParam, IntPtr lParam);
        private const int WH_CBT = 5;
        private const int HCBT_ACTIVATE = 5;
        private static HookProc _hookProc;
        private static IntPtr _hHook = IntPtr.Zero;
        private static Window _ownerWindow;
        private static string _caption;

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        private static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static MessageBoxResult ShowCentered(Window owner, string messageBoxText, string caption, MessageBoxButton button, MessageBoxImage icon)
        {
            _ownerWindow = owner;
            _caption = caption;
            _hookProc = new HookProc(CbtHookProc);
            _hHook = SetWindowsHookEx(WH_CBT, _hookProc, IntPtr.Zero, GetCurrentThreadId());

            try
            {
                return MessageBox.Show(owner, messageBoxText, caption, button, icon);
            }
            finally
            {
                if (_hHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hHook);
                    _hHook = IntPtr.Zero;
                }
            }
        }

        private static IntPtr CbtHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == HCBT_ACTIVATE)
            {
                // Center the MessageBox
                IntPtr hMsgBox = wParam;
                if (_ownerWindow != null)
                {
                    WindowInteropHelper helper = new WindowInteropHelper(_ownerWindow);
                    IntPtr hOwner = helper.Handle;
                    if (GetWindowRect(hOwner, out RECT ownerRect) && GetWindowRect(hMsgBox, out RECT msgRect))
                    {
                        int ownerCenterX = (ownerRect.Left + ownerRect.Right) / 2;
                        int ownerCenterY = (ownerRect.Top + ownerRect.Bottom) / 2;
                        int msgWidth = msgRect.Right - msgRect.Left;
                        int msgHeight = msgRect.Bottom - msgRect.Top;
                        int newX = ownerCenterX - msgWidth / 2;
                        int newY = ownerCenterY - msgHeight / 2;
                        MoveWindow(hMsgBox, newX, newY, msgWidth, msgHeight, true);
                        SetForegroundWindow(hMsgBox);
                    }
                }
                // Unhook after centering
                if (_hHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_hHook);
                    _hHook = IntPtr.Zero;
                }
            }
            return CallNextHookEx(_hHook, nCode, wParam, lParam);
        }
    }
}
