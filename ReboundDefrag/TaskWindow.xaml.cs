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
using System.Text;
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

        public class DiskItem : Item
        {
            public string DriveLetter { get; set; }
            public string MediaType { get; set; }
        }

        public TaskWindow()
        {
            this.InitializeComponent();
            this.IsMaximizable = false;
            this.IsMinimizable = false;
            this.SetWindowSize(600, 450);
            this.CenterOnScreen();
            this.IsResizable = false;
            this.Title = "Optimization schedule";
            RemoveIcon(this);
            LoadData();
        }

        private async void LoadData()
        {
            // Get the logical drives bitmask
            uint drivesBitMask = GetLogicalDrives();
            if (drivesBitMask == 0)
            {
                Console.WriteLine("Failed to get logical drives.");
                return;
            }

            List<DiskItem> items = new List<DiskItem>();
            Console.WriteLine("System Partitions:");
            for (char driveLetter = 'A'; driveLetter <= 'Z'; driveLetter++)
            {
                uint mask = 1u << (driveLetter - 'A');
                if ((drivesBitMask & mask) != 0)
                {
                    string drive = $"{driveLetter}:\\";
                    Console.WriteLine(drive);

                    StringBuilder volumeName = new StringBuilder(261);
                    StringBuilder fileSystemName = new StringBuilder(261);
                    if (GetVolumeInformation(drive, volumeName, volumeName.Capacity, out _, out _, out _, fileSystemName, fileSystemName.Capacity))
                    {
                        var newDriveLetter = drive.ToString().Remove(2, 1);
                        Console.WriteLine($"  Volume Name: {volumeName}");
                        Console.WriteLine($"  File System: {fileSystemName}");

                        if (volumeName.ToString() != string.Empty)
                        {
                            items.Add(new DiskItem
                            {
                                Name = $"{volumeName} ({newDriveLetter})",
                                ImagePath = "ms-appx:///Assets/Icon1.png",
                                MediaType = null,
                                DriveLetter = drive,
                            });
                        }
                        else
                        {
                            items.Add(new DiskItem
                            {
                                Name = $"({newDriveLetter})",
                                ImagePath = "ms-appx:///Assets/Icon1.png",
                                MediaType = null,
                                DriveLetter = drive,
                            });
                        }
                    }
                    else
                    {
                        Console.WriteLine($"  Failed to get volume information for {drive}");
                    }
                }
            }
            MyListView.ItemsSource = items;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint GetLogicalDrives();

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool GetVolumeInformation(
            string lpRootPathName,
            StringBuilder lpVolumeNameBuffer,
            int nVolumeNameSize,
            out uint lpVolumeSerialNumber,
            out uint lpMaximumComponentLength,
            out uint lpFileSystemFlags,
            StringBuilder lpFileSystemNameBuffer,
            int nFileSystemNameSize);

    }
}
