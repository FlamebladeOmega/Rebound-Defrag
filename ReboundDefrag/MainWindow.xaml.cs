using Microsoft.UI;
using Microsoft.UI.Composition.SystemBackdrops;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using WinUIEx;
using Microsoft.UI.Xaml.Hosting;
using Windows.Media.Core;
using Windows.Media.Playback;
using Windows.UI;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.Eventing.Reader;
using WinUIEx.Messaging;
using Windows.UI.Core;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Security.Principal;
using System.Text.RegularExpressions;
using Windows.Devices.Portable;
using Windows.Management;
using Microsoft.Management.Infrastructure;
using System.Management;
using Microsoft.Win32.SafeHandles;
using Microsoft.UI.Windowing;

namespace ReboundDefrag
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : WindowEx
    {
        public MainWindow()
        {
            this.InitializeComponent();
            LoadWindowProperties();
        }

        public void LoadWindowProperties()
        {
            // Set standard window properties
            SystemBackdrop = new MicaBackdrop();
            Title = "Defragment and Optimize Drives";
            LoadData();
            this.IsMaximizable = false;
            this.SetWindowSize(800, 670);
            this.IsResizable = false;
            this.AppWindow.DefaultTitleBarShouldMatchAppModeTheme = true;
            this.CenterOnScreen();

            // Begin window message reading
            var hWnd = WinRT.Interop.WindowNative.GetWindowHandle(this);
            WindowMessageMonitor mon = new WindowMessageMonitor(hWnd);
            mon.WindowMessageReceived += MessageReceived;
            void MessageReceived(object sender, WindowMessageEventArgs e)
            {
                // Variables
                const int WM_DEVICECHANGE = 0x0219;
                const int DBT_DEVICEARRIVAL = 0x8000;
                const int DBT_DEVICEREMOVECOMPLETE = 0x8004;

                // Switch messages
                switch (e.Message.MessageId)
                {
                    default:
                        {
                            break;
                        }
                    case WM_DEVICECHANGE:
                        {
                            switch ((int)e.Message.WParam)
                            {
                                case DBT_DEVICEARRIVAL:
                                    {
                                        // Drive or partition inserted
                                        MyListView.ItemsSource = null;
                                        LoadData();
                                        break;
                                    }
                                case DBT_DEVICEREMOVECOMPLETE:
                                    {
                                        // Drive or partition removed
                                        MyListView.ItemsSource = null;
                                        LoadData();
                                        break;
                                    }
                                default:
                                    {
                                        // Drive or partition action
                                        break;
                                    }
                            }
                            break;
                        }
                }
            };

            // Create a timer to reload the message listener every 5 seconds
            var timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(5);
            timer.Tick += (sender, e) =>
            {
                mon.WindowMessageReceived += MessageReceived;
            };
            timer.Start();
        }

        public class DiskItem : Item
        {
            public string DriveLetter { get; set; }
            public string MediaType { get; set; }
        }

        #region DLL

        private static IntPtr SetWindowLong(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 4)
            {
                return SetWindowLongPtr32(hWnd, nIndex, dwNewLong);
            }
            return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
        }

        [DllImport("User32.dll", CharSet = CharSet.Auto, EntryPoint = "SetWindowLong")]
        private static extern IntPtr SetWindowLongPtr32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("User32.dll", CharSet = CharSet.Auto, EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        private const int GWL_HWNDPARENT = (-8);

        [DllImport("User32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool EnableWindow(IntPtr hWnd, bool bEnable);

        public static void CreateModalWindow(WindowEx parentWindow, WindowEx childWindow, bool summonWindowAutomatically = true, bool blockInput = false)
        {
            IntPtr hWndChildWindow = WinRT.Interop.WindowNative.GetWindowHandle(childWindow);
            IntPtr hWndParentWindow = WinRT.Interop.WindowNative.GetWindowHandle(parentWindow);
            SetWindowLong(hWndChildWindow, GWL_HWNDPARENT, hWndParentWindow);
            (childWindow.AppWindow.Presenter as OverlappedPresenter).IsModal = true;
            if (blockInput == true)
            {
                EnableWindow(hWndParentWindow, false);
                childWindow.Closed += ChildWindow_Closed;
                void ChildWindow_Closed(object sender, Microsoft.UI.Xaml.WindowEventArgs args)
                {
                    EnableWindow(hWndParentWindow, true);
                }
            }
            if (summonWindowAutomatically == true) childWindow.Show();
            WindowMessageMonitor _msgMonitor;

            _msgMonitor = new WindowMessageMonitor(childWindow);
            _msgMonitor.WindowMessageReceived += (_, e) =>
            {
                const int WM_NCLBUTTONDBLCLK = 0x00A3;
                if (e.Message.MessageId == WM_NCLBUTTONDBLCLK)
                {
                    // Disable double click on title bar to maximize window
                    e.Result = 0;
                    e.Handled = true;
                }
            };
        }

        [DllImport("dwmapi.dll", SetLastError = true)]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

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

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(
            IntPtr hDevice,
            uint dwIoControlCode,
            IntPtr lpInBuffer,
            uint nInBufferSize,
            IntPtr lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        private const uint IOCTL_DISK_GET_DRIVE_GEOMETRY_EX = 0x000700A0;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 0x00000003;
        private const uint GENERIC_READ = 0x80000000;

        [StructLayout(LayoutKind.Sequential)]
        private struct DISK_GEOMETRY_EX
        {
            public long DiskSize;
            public MEDIA_TYPE MediaType;
            // Other members are omitted for brevity
        }

        private enum MEDIA_TYPE
        {
            Unknown,
            RemovableMedia,
            FixedMedia,
            // Other values are omitted for brevity
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern DriveType GetDriveType(string lpRootPathName);

        private enum DriveType : uint
        {
            DRIVE_UNKNOWN = 0,
            DRIVE_NO_ROOT_DIR = 1,
            DRIVE_REMOVABLE = 2,
            DRIVE_FIXED = 3,
            DRIVE_REMOTE = 4,
            DRIVE_CDROM = 5,
            DRIVE_RAMDISK = 6
        }

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern uint GetDriveTypeWin32(string lpRootPathName);

        private const int InvalidHandleValue = -1;
        private const uint FILE_DEVICE_SSD = 0x00000060;

        [StructLayout(LayoutKind.Sequential)]
        private struct STORAGE_PROPERTY_QUERY
        {
            public uint PropertyId;
            public uint QueryType;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
            public byte[] AdditionalParameters;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct STORAGE_DESCRIPTOR_HEADER
        {
            public uint Version;
            public uint Size;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct STORAGE_DEVICE_DESCRIPTOR
        {
            public uint Version;
            public uint Size;
            public byte DeviceType;
            public byte DeviceTypeModifier;
            [MarshalAs(UnmanagedType.U1)]
            public bool RemovableMedia;
            [MarshalAs(UnmanagedType.U1)]
            public bool CommandQueueing;
            public uint VendorIdOffset;
            public uint ProductIdOffset;
            public uint ProductRevisionOffset;
            public uint SerialNumberOffset;
            public STORAGE_BUS_TYPE BusType;
            public uint RawPropertiesLength;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 10240)]
            public byte[] RawDeviceProperties;
        }

        private enum STORAGE_BUS_TYPE
        {
            BusTypeUnknown = 0x00,
            BusTypeScsi = 0x1,
            BusTypeAtapi = 0x2,
            BusTypeAta = 0x3,
            BusType1394 = 0x4,
            BusTypeSsa = 0x5,
            BusTypeFibre = 0x6,
            BusTypeUsb = 0x7,
            BusTypeRAID = 0x8,
            BusTypeiScsi = 0x9,
            BusTypeSas = 0xA,
            BusTypeSata = 0xB,
            BusTypeSd = 0xC,
            BusTypeMmc = 0xD,
            BusTypeVirtual = 0xE,
            BusTypeFileBackedVirtual = 0xF,
            BusTypeSpaces = 0x10,
            BusTypeNvme = 0x11,
            BusTypeSCM = 0x12,
            BusTypeUfs = 0x13,
            BusTypeMax = 0x14,
            BusTypeMaxReserved = 0x7F
        }

        #endregion DLL

        private void ListviewSelectionChange(object sender, SelectionChangedEventArgs args)
        {
            LoadSelectedItemInfo(GetStatus());
        }

        public string GetLastOptimizeDate()
        {
            if (MyListView.SelectedItem != null)
            {
                DiskItem selectedItem = MyListView.SelectedItem as DiskItem;
                // Handle the selection change event

                string selI;
                try
                {
                    List<string> i = DefragInfo.GetEventLogEntriesForID(258);
                    return i.Last(s => s.Contains($"({selectedItem.DriveLetter.ToString().Remove(2, 1)})"));
                }
                catch
                {
                    return "Never....";
                }
            }
            else return "Unknown....";
        }

        public void LoadSelectedItemInfo(string status, string info = "....")
        {
            if (info == "....")
            {
                info = GetLastOptimizeDate();
            }
            else
            {
                info = "Unknown....";
            }
            if (MyListView.SelectedItem != null)
            {
                DiskItem selectedItem = MyListView.SelectedItem as DiskItem;
                DetailsBar.Title = selectedItem.Name;
                DetailsBar.Message = $"Media type: {selectedItem.MediaType}\nLast analyzed or optimized: {info.Substring(0, info.Length - 4)}\nCurrent status: {status}";
            }
        }

        public string GetStatus()
        {
            string status = string.Empty;

            try
            {
                if (MyListView.SelectedItem != null)
                {
                    List<string> i = DefragInfo.GetEventLogEntriesForID(258);
                    var selectedItem = MyListView.SelectedItem as DiskItem;

                    var selI = i.Last(s => s.Contains($"({selectedItem.DriveLetter.ToString().Remove(2, 1)})"));

                    var localDate = DateTime.Parse(selI.Substring(0, selI.Length - 4));

                    // Get the current local date and time
                    DateTime currentDate = DateTime.Now;

                    // Calculate the days passed
                    if (localDate != null)
                    {
                        TimeSpan timeSpan = (TimeSpan)(currentDate - localDate);
                        int daysPassed = timeSpan.Days;

                        if (daysPassed == 0)
                        {
                            return $"OK (Last optimized: today)";
                        }

                        if (daysPassed == 1)
                        {
                            return $"OK (Last optimized: yesterday)";
                        }

                        if (daysPassed < 50)
                        {
                            return $"OK (Last optimized: {daysPassed} days ago)";
                        }

                        if (daysPassed >= 50)
                        {
                            return $"Needs optimization (Last optimized: {daysPassed} days ago)";
                        }

                        else return "Unknown";
                    }
                    else
                    {
                        return "Unknown";
                    }
                }
                else
                {
                    return "Please select an item to proceed.";
                }
            }
            catch (Exception ex)
            {
                return "Needs optimization";
            }
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

                        DriveType driveType;

                        try
                        {
                            // Get drive type
                            driveType = GetDriveType(drive);
                        }
                        catch
                        {
                            driveType = DriveType.DRIVE_UNKNOWN;
                        }

                        string mediaType = "Unknown";
                        mediaType = await GetDriveTypeDescriptionAsync(drive);

                        if (volumeName.ToString() != string.Empty)
                        {
                            items.Add(new DiskItem
                            {
                                Name = $"{volumeName} ({newDriveLetter})",
                                ImagePath = "ms-appx:///Assets/Icon1.png",
                                MediaType = mediaType,
                                DriveLetter = drive,
                            });
                        }
                        else
                        {
                            items.Add(new DiskItem
                            {
                                Name = $"({newDriveLetter})",
                                ImagePath = "ms-appx:///Assets/Icon1.png",
                                MediaType = mediaType,
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

        private async void LoadDriveTypes()
        {
            var drives = System.IO.DriveInfo.GetDrives();
            List<string> driveTypes = new List<string>();

            foreach (var drive in drives)
            {
                string driveTypeDescription = $"{drive.Name} - {await GetDriveTypeDescriptionAsync(drive.RootDirectory.FullName)}";
                driveTypes.Add(driveTypeDescription);
            }
        }

        private async Task<string> GetDriveTypeDescriptionAsync(string driveRoot)
        {
            DriveType driveType = GetDriveType(driveRoot);

            switch (driveType)
            {
                case DriveType.DRIVE_REMOVABLE:
                    return "Removable";
                case DriveType.DRIVE_FIXED:
                    return StorageDeviceService.HasNoSeekPenalty(StorageDeviceService.GetDiskExtents(driveRoot.Substring(0, 1).ToCharArray()[0]).ToString()) ? "HDD" : "SSD"; // Await the asynchronous method
                case DriveType.DRIVE_REMOTE:
                    return "Network";
                case DriveType.DRIVE_CDROM:
                    return "CD-ROM";
                case DriveType.DRIVE_RAMDISK:
                    return "RAM Disk";
                case DriveType.DRIVE_UNKNOWN:
                default:
                    return "Unknown";
            }
        }

        private void Button_Click(object sender, SplitButtonClickEventArgs e)
        {
            OptimizeSelected();
        }

        public async void OptimizeSelected()
        {
            foreach (DiskItem item in MyListView.SelectedItems)
            {
                string volume = item.DriveLetter.ToString().Remove(1, 2);
                string arguments = $"Optimize-Volume -DriveLetter {volume}"; // /O to optimize the drive
                                                                             //string arguments = $"Defrag {volume}: /O /U"; // /O to optimize the drive

                /*if (!IsAdmin())
                {
                    ShowMessage("The application must be running as Administrator to perform this task.");
                    return;
                }*/

                try
                {
                    var processInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = arguments,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        Verb = "runas"  // Run as administrator
                    };

                    var process = new Process
                    {
                        StartInfo = processInfo,
                        //EnableRaisingEvents = true
                    };

                    //process.OutputDataReceived += (s, ea) => UpdateProgress(ea.Data);
                    //process.ErrorDataReceived += (s, ea) => UpdateProgress(ea.Data);

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    LoadSelectedItemInfo("Optimizing...");
                    CurrentProgress.IsIndeterminate = true;
                    OptimizeButton.IsEnabled = false;

                    await process.WaitForExitAsync();

                    LoadSelectedItemInfo(GetStatus());
                    CurrentProgress.IsIndeterminate = false;
                    OptimizeButton.IsEnabled = true;
                }
                catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
                {
                    // Error 1223 indicates that the operation was canceled by the user (UAC prompt declined)
                    ShowMessage("Defragmentation was canceled by the user.");
                }
                catch (Exception ex)
                {
                    ShowMessage($"Error: {ex.Message}");
                }
            }
        }

        public async void OptimizeAll()
        {
            int i = 0;
            int j = (MyListView.ItemsSource as List<DiskItem>).Count;

            MyListView.SelectedIndex = 0;

            foreach (var item in MyListView.ItemsSource as List<DiskItem>)
            {
                string volume = (item as DiskItem).DriveLetter.ToString().Remove(1, 2);
                string arguments = $"Optimize-Volume -DriveLetter {volume}"; // /O to optimize the drive
                                                                             //string arguments = $"Defrag {volume}: /O /U"; // /O to optimize the drive

                /*if (!IsAdmin())
                {
                    ShowMessage("The application must be running as Administrator to perform this task.");
                    return;
                }*/

                try
                {
                    i++;

                    var processInfo = new ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = arguments,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true,
                        Verb = "runas"  // Run as administrator
                    };

                    var process = new Process
                    {
                        StartInfo = processInfo,
                        //EnableRaisingEvents = true
                    };

                    //process.OutputDataReceived += (s, ea) => UpdateProgress(ea.Data);
                    //process.ErrorDataReceived += (s, ea) => UpdateProgress(ea.Data);

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    LoadSelectedItemInfo("Optimizing...");
                    CurrentProgress.IsIndeterminate = true;
                    OptimizeButton.IsEnabled = false;
                    CurrentDisk.Visibility = Visibility.Visible;
                    CurrentDisk.Text = $"Drive {i}/{j} ({volume}:) - Optimizing...";

                    await process.WaitForExitAsync();

                    LoadSelectedItemInfo(GetStatus());
                    CurrentProgress.IsIndeterminate = false;
                    OptimizeButton.IsEnabled = true;
                    CurrentDisk.Visibility = Visibility.Collapsed;

                    if (MyListView.SelectedIndex + 1 != j) MyListView.SelectedIndex++;
                }
                catch (System.ComponentModel.Win32Exception ex) when (ex.NativeErrorCode == 1223)
                {
                    // Error 1223 indicates that the operation was canceled by the user (UAC prompt declined)
                    ShowMessage("Defragmentation was canceled by the user.");
                }
                catch (Exception ex)
                {
                    ShowMessage($"Error: {ex.Message}");
                }
            }

            i = 0;
        }

        private void UpdateProgress(string data)
        {
            if (string.IsNullOrWhiteSpace(data)) return;

            // Parse the progress data from PowerShell output
            var match = Regex.Match(data, @"\[(o+)(\s*)\]");

            if (match.Success)
            {
                int progress = (int)((double)match.Groups[1].Value.Length / (match.Groups[1].Value.Length + match.Groups[2].Value.Length) * 100);

                // Update the UI with the progress
                DispatcherQueue.TryEnqueue(() =>
                {
                    UpdateProgress2($"{progress}% completed");
                });
            }
            else
            {
                // Handle other types of messages
                DispatcherQueue.TryEnqueue(() =>
                {
                    UpdateProgress2(data + Environment.NewLine);
                });
            }
        }

        private void ParseOutput(string data)
        {
            // Example line to match: "50% complete."
            //UpdateProgress(data);
            if (data != null)
            {
                Debug.WriteLine(data);
                    int i = 0;
                    foreach (char c in data)
                    {
                        if (c == 'o') i++;
                    }
                    UpdateProgress($"{i}%");
                //ShowMessage(data);
            }
        }

        public void UpdateProgress2(string progress)
        {
            var selectedItem = new DiskItem();
            selectedItem.DriveLetter = "C:/";
            selectedItem = MyListView.SelectedItem as DiskItem;
            // Update UI from the main thread
            _ = DispatcherQueue.TryEnqueue(() =>
            {
                DetailsBar.Title = selectedItem.Name;
                string selI;
                try
                {
                    List<string> i = DefragInfo.GetEventLogEntriesForID(258);
                    selI = i.Last(s => s.Contains($"({selectedItem.DriveLetter.ToString().Remove(2, 1)})"));
                }
                catch
                {
                    selI = "Never....";
                }
                DetailsBar.Message = $"Media type: {selectedItem.MediaType}\nLast analyzed or optimized: {selI.Substring(0, selI.Length - 4)}\nCurrent status: Optimizing ({progress}%)";
            });
        }

        private async void ShowMessage(string message)
        {
            try
            {
                var dialog = new ContentDialog
                {
                    Title = "Defragmentation",
                    Content = message,
                    CloseButtonText = "OK",
                };

                dialog.XamlRoot = this.Content.XamlRoot;
                _ = await dialog.ShowAsync();
            }
            catch (Exception ex)
            {

            }
        }

        private void MenuFlyoutItem_Click(object sender, RoutedEventArgs e)
        {
            OptimizeSelected();
        }

        private void MenuFlyoutItem_Click_1(object sender, RoutedEventArgs e)
        {
            OptimizeAll();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            CreateModalWindow(this, new TaskWindow(), true, true);
        }
    }
    public class DefragInfo
    {
        public static DateTime? GetLastDefragTime(string driveLetter)
        {
            string driveRoot = driveLetter + ":\\";
            string query = "*[System[Provider[@Name='Microsoft-Windows-Defrag'] and (EventID=258 or EventID=259)]]";

            EventLogQuery eventsQuery = new EventLogQuery("Application", PathType.LogName, query);
            eventsQuery.ReverseDirection = true;
            eventsQuery.TolerateQueryErrors = true;

            List<EventRecord> events = new List<EventRecord>();

            using (EventLogReader logReader = new EventLogReader(eventsQuery))
            {
                for (EventRecord eventInstance = logReader.ReadEvent(); eventInstance != null; eventInstance = logReader.ReadEvent())
                {
                    if (eventInstance.Properties.Count > 0)
                    {
                        string message = eventInstance.Properties.Last().Value.ToString();
                        if (message.Contains(driveRoot))
                        {
                            return eventInstance.TimeCreated?.ToLocalTime();
                        }
                    }
                }
            }

            return null;
        }

        public static List<string> GetEventLogEntriesForID(int eventID)
        {
            List<string> eventMessages = new List<string>();

            // Define the query
            string logName = "Application"; // Windows Logs > Application
            string queryStr = "*[System/EventID=" + eventID + "]";

            EventLogQuery query = new EventLogQuery(logName, PathType.LogName, queryStr);

            // Create the reader
            using (EventLogReader reader = new EventLogReader(query))
            {
                for (EventRecord eventInstance = reader.ReadEvent(); eventInstance != null; eventInstance = reader.ReadEvent())
                {
                    // Extract the message from the event
                    string sb = eventInstance.TimeCreated.ToString() + eventInstance.FormatDescription().ToString().Substring(eventInstance.FormatDescription().ToString().Length - 4);

                    eventMessages.Add(sb.ToString());
                }
            }

            return eventMessages;
        }
    }

    public class StorageDevice
    {
        public string DeviceID { get; set; }
        public string MediaType { get; set; }
    }

    public class StorageDeviceService
    {
        // For CreateFile to get handle to drive
        private const uint GENERIC_READ = 0x80000000;
        private const uint GENERIC_WRITE = 0x40000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 3;
        private const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

        // CreateFile to get handle to drive
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(
            [MarshalAs(UnmanagedType.LPWStr)]
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        // For control codes
        private const uint FILE_DEVICE_MASS_STORAGE = 0x0000002d;
        private const uint IOCTL_STORAGE_BASE = FILE_DEVICE_MASS_STORAGE;
        private const uint FILE_DEVICE_CONTROLLER = 0x00000004;
        private const uint IOCTL_SCSI_BASE = FILE_DEVICE_CONTROLLER;
        private const uint METHOD_BUFFERED = 0;
        private const uint FILE_ANY_ACCESS = 0;
        private const uint FILE_READ_ACCESS = 0x00000001;
        private const uint FILE_WRITE_ACCESS = 0x00000002;

        private static uint CTL_CODE(uint DeviceType, uint Function,
                                     uint Method, uint Access)
        {
            return ((DeviceType << 16) | (Access << 14) |
                    (Function << 2) | Method);
        }

        // For DeviceIoControl to check no seek penalty
        private const uint StorageDeviceSeekPenaltyProperty = 7;
        private const uint PropertyStandardQuery = 0;

        [StructLayout(LayoutKind.Sequential)]
        private struct STORAGE_PROPERTY_QUERY
        {
            public uint PropertyId;
            public uint QueryType;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 1)]
            public byte[] AdditionalParameters;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct DEVICE_SEEK_PENALTY_DESCRIPTOR
        {
            public uint Version;
            public uint Size;
            [MarshalAs(UnmanagedType.U1)]
            public bool IncursSeekPenalty;
        }

        // DeviceIoControl to check no seek penalty
        [DllImport("kernel32.dll", EntryPoint = "DeviceIoControl",
                   SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(
            SafeFileHandle hDevice,
            uint dwIoControlCode,
            ref STORAGE_PROPERTY_QUERY lpInBuffer,
            uint nInBufferSize,
            ref DEVICE_SEEK_PENALTY_DESCRIPTOR lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        // For DeviceIoControl to check nominal media rotation rate
        private const uint ATA_FLAGS_DATA_IN = 0x02;

        [StructLayout(LayoutKind.Sequential)]
        private struct ATA_PASS_THROUGH_EX
        {
            public ushort Length;
            public ushort AtaFlags;
            public byte PathId;
            public byte TargetId;
            public byte Lun;
            public byte ReservedAsUchar;
            public uint DataTransferLength;
            public uint TimeOutValue;
            public uint ReservedAsUlong;
            public IntPtr DataBufferOffset;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
            public byte[] PreviousTaskFile;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
            public byte[] CurrentTaskFile;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ATAIdentifyDeviceQuery
        {
            public ATA_PASS_THROUGH_EX header;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
            public ushort[] data;
        }

        // DeviceIoControl to check nominal media rotation rate
        [DllImport("kernel32.dll", EntryPoint = "DeviceIoControl",
                   SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(
            SafeFileHandle hDevice,
            uint dwIoControlCode,
            ref ATAIdentifyDeviceQuery lpInBuffer,
            uint nInBufferSize,
            ref ATAIdentifyDeviceQuery lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        // For error message
        private const uint FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern uint FormatMessage(
            uint dwFlags,
            IntPtr lpSource,
            uint dwMessageId,
            uint dwLanguageId,
            StringBuilder lpBuffer,
            uint nSize,
            IntPtr Arguments);

        // Method for no seek penalty
        public static bool HasNoSeekPenalty(string sDrive)
        {
            SafeFileHandle hDrive = CreateFileW(
                sDrive,
                0, // No access to drive
                FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                IntPtr.Zero);

            if (hDrive == null || hDrive.IsInvalid)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Console.WriteLine("CreateFile failed. " + message);
            }

            uint IOCTL_STORAGE_QUERY_PROPERTY = CTL_CODE(
                IOCTL_STORAGE_BASE, 0x500,
                METHOD_BUFFERED, FILE_ANY_ACCESS); // From winioctl.h

            STORAGE_PROPERTY_QUERY query_seek_penalty =
                new STORAGE_PROPERTY_QUERY();
            query_seek_penalty.PropertyId = StorageDeviceSeekPenaltyProperty;
            query_seek_penalty.QueryType = PropertyStandardQuery;

            DEVICE_SEEK_PENALTY_DESCRIPTOR query_seek_penalty_desc =
                new DEVICE_SEEK_PENALTY_DESCRIPTOR();

            uint returned_query_seek_penalty_size;

            bool query_seek_penalty_result = DeviceIoControl(
                hDrive,
                IOCTL_STORAGE_QUERY_PROPERTY,
                ref query_seek_penalty,
                (uint)Marshal.SizeOf(query_seek_penalty),
                ref query_seek_penalty_desc,
                (uint)Marshal.SizeOf(query_seek_penalty_desc),
                out returned_query_seek_penalty_size,
                IntPtr.Zero);

            hDrive.Close();

            if (query_seek_penalty_result == false)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Console.WriteLine("DeviceIoControl failed. " + message);
                return false;
            }
            else
            {
                if (query_seek_penalty_desc.IncursSeekPenalty == false)
                {
                    Console.WriteLine("This drive has NO SEEK penalty.");
                    return false;
                }
                else
                {
                    Console.WriteLine("This drive has SEEK penalty.");
                    return true;
                }
            }
        }

        // Method for nominal media rotation rate
        // (Administrative privilege is required)
        public static bool HasNominalMediaRotationRate(string sDrive)
        {
            SafeFileHandle hDrive = CreateFileW(
                sDrive,
                GENERIC_READ | GENERIC_WRITE, // Administrative privilege is required
                FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                IntPtr.Zero);

            if (hDrive == null || hDrive.IsInvalid)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Console.WriteLine("CreateFile failed. " + message);
            }

            uint IOCTL_ATA_PASS_THROUGH = CTL_CODE(
                IOCTL_SCSI_BASE, 0x040b, METHOD_BUFFERED,
                FILE_READ_ACCESS | FILE_WRITE_ACCESS); // From ntddscsi.h

            ATAIdentifyDeviceQuery id_query = new ATAIdentifyDeviceQuery();
            id_query.data = new ushort[256];

            id_query.header.Length = (ushort)Marshal.SizeOf(id_query.header);
            id_query.header.AtaFlags = (ushort)ATA_FLAGS_DATA_IN;
            id_query.header.DataTransferLength =
                (uint)(id_query.data.Length * 2); // Size of "data" in bytes
            id_query.header.TimeOutValue = 3; // Sec
            id_query.header.DataBufferOffset = (IntPtr)Marshal.OffsetOf(
                typeof(ATAIdentifyDeviceQuery), "data");
            id_query.header.PreviousTaskFile = new byte[8];
            id_query.header.CurrentTaskFile = new byte[8];
            id_query.header.CurrentTaskFile[6] = 0xec; // ATA IDENTIFY DEVICE

            uint retval_size;

            bool result = DeviceIoControl(
                hDrive,
                IOCTL_ATA_PASS_THROUGH,
                ref id_query,
                (uint)Marshal.SizeOf(id_query),
                ref id_query,
                (uint)Marshal.SizeOf(id_query),
                out retval_size,
                IntPtr.Zero);

            hDrive.Close();

            if (result == false)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Console.WriteLine("DeviceIoControl failed. " + message);
                return true;
            }
            else
            {
                // Word index of nominal media rotation rate
                // (1 means non-rotate device)
                const int kNominalMediaRotRateWordIndex = 217;

                if (id_query.data[kNominalMediaRotRateWordIndex] == 1)
                {
                    Console.WriteLine("This drive is NON-ROTATE device.");
                    return false;
                }
                else
                {
                    Console.WriteLine("This drive is ROTATE device.");
                    return true;
                }
            }
        }

        // Method for error message
        private static string GetErrorMessage(int code)
        {
            StringBuilder message = new StringBuilder(255);

            FormatMessage(
              FORMAT_MESSAGE_FROM_SYSTEM,
              IntPtr.Zero,
              (uint)code,
              0,
              message,
              (uint)message.Capacity,
              IntPtr.Zero);

            return message.ToString();
        }

        // For control codes
        private const uint IOCTL_VOLUME_BASE = 0x00000056;

        // For DeviceIoControl to get disk extents
        [StructLayout(LayoutKind.Sequential)]
        private struct DISK_EXTENT
        {
            public uint DiskNumber;
            public long StartingOffset;
            public long ExtentLength;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct VOLUME_DISK_EXTENTS
        {
            public uint NumberOfDiskExtents;
            [MarshalAs(UnmanagedType.ByValArray)]
            public DISK_EXTENT[] Extents;
        }

        // DeviceIoControl to get disk extents
        [DllImport("kernel32.dll", EntryPoint = "DeviceIoControl",
                   SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(
            SafeFileHandle hDevice,
            uint dwIoControlCode,
            IntPtr lpInBuffer,
            uint nInBufferSize,
            ref VOLUME_DISK_EXTENTS lpOutBuffer,
            uint nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped);

        // Method for disk extents
        public static int GetDiskExtents(char cDrive)
        {
            DriveInfo di = new DriveInfo(cDrive.ToString());
            if (di.DriveType != DriveType.Fixed)
            {
                Console.WriteLine("This drive is not fixed drive.");
            }

            string sDrive = "\\\\.\\" + cDrive.ToString() + ":";

            SafeFileHandle hDrive = CreateFileW(
                sDrive,
                0, // No access to drive
                FILE_SHARE_READ | FILE_SHARE_WRITE,
                IntPtr.Zero,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                IntPtr.Zero);

            if (hDrive == null || hDrive.IsInvalid)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Console.WriteLine("CreateFile failed. " + message);
            }

            uint IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS = CTL_CODE(
                IOCTL_VOLUME_BASE, 0,
                METHOD_BUFFERED, FILE_ANY_ACCESS); // From winioctl.h

            VOLUME_DISK_EXTENTS query_disk_extents =
                new VOLUME_DISK_EXTENTS();

            uint returned_query_disk_extents_size;

            bool query_disk_extents_result = DeviceIoControl(
                hDrive,
                IOCTL_VOLUME_GET_VOLUME_DISK_EXTENTS,
                IntPtr.Zero,
                0,
                ref query_disk_extents,
                (uint)Marshal.SizeOf(query_disk_extents),
                out returned_query_disk_extents_size,
                IntPtr.Zero);

            hDrive.Close();

            if (query_disk_extents_result == false ||
                query_disk_extents.Extents.Length != 1)
            {
                string message = GetErrorMessage(Marshal.GetLastWin32Error());
                Console.WriteLine("DeviceIoControl failed. " + message);
            }
            else
            {
                Console.WriteLine("The physical drive number is: " +
                                  query_disk_extents.Extents[0].DiskNumber);
            }

            return (int)query_disk_extents.Extents[0].DiskNumber;
        }
    }
}
