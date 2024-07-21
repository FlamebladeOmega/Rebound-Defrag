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
    public sealed partial class TaskWindow : WindowEx
    {
        public static void RemoveIcon(WindowEx window)
        {
            IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(window);
            SetWindowLongPtr(hWnd, -20, 0x00000001L);
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, long dwNewLong);

        public TaskWindow()
        {
            this.InitializeComponent();
            this.IsMaximizable = false;
            this.IsMinimizable = false;
            this.SetWindowSize(600, 350);
            this.CenterOnScreen();
            this.IsResizable = false;
            RemoveIcon(this);
        }
    }
}
