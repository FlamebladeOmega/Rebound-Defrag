using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using Windows.Foundation;
using Windows.Foundation.Collections;
using WinUIEx;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace ReboundDefrag
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class AdminWindow : WindowEx
    {
        public AdminWindow()
        {
            this.InitializeComponent();
            this.AppWindow.TitleBar.ExtendsContentIntoTitleBar = true;
            this.AppWindow.TitleBar.PreferredHeightOption = Microsoft.UI.Windowing.TitleBarHeightOption.Collapsed;
            this.Maximize();
            this.SystemBackdrop = new TransparentTintBackdrop();
            HideFromTaskbar();
            Launch();
        }

        public async void Launch()
        {
            await Task.Delay(50);
            this.SetIsAlwaysOnTop(true);
            await AdminDialog.ShowAsync();
            Close();
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const int WS_EX_APPWINDOW = 0x00040000;

        private void HideFromTaskbar()
        {
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            IntPtr exStyle = GetWindowLongPtr(hWnd, GWL_EXSTYLE);

            // Remove the WS_EX_APPWINDOW style and add the WS_EX_TOOLWINDOW style
            IntPtr newExStyle = (IntPtr)((long)exStyle & ~WS_EX_APPWINDOW | WS_EX_TOOLWINDOW);
            SetWindowLongPtr(hWnd, GWL_EXSTYLE, newExStyle);

            // Update the window's position to apply the style changes
            SetWindowPos(hWnd, IntPtr.Zero, 0, 0, 0, 0, 0x027);
        }
    }
}
