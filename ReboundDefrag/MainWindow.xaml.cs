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

namespace ReboundDefrag
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : WindowEx
    {
        int i = 0;

        public MainWindow()
        {
            this.InitializeComponent();
            SystemBackdrop = new MicaBackdrop();

            Title = "Defragment and Optimize Drives";
            LoadData();
            SetDarkMode(this);
            this.IsMaximizable = false;
            this.SetWindowSize(650, 400);
            this.IsResizable = false;
            AdministratorStatusTextBlock.Text = IsAdmin() is true
                ? "Running as admin."
                : "NOT running as admin.";
            //this.AppWindow.TitleBar.BackgroundColor = Colors.Transparent;
            //this.AppWindow.TitleBar.ButtonBackgroundColor = Colors.Transparent;
            //this.AppWindow.TitleBar.ButtonInactiveBackgroundColor = Colors.Transparent;
            WindowMessageMonitor mon = new WindowMessageMonitor(this);
            mon.WindowMessageReceived += MessageReceived;
            void MessageReceived(object sender, WindowMessageEventArgs e)
            {
                const int WM_DEVICECHANGE = 0x0219;
                const int DBT_DEVICEARRIVAL = 0x8000;
                const int DBT_DEVICEREMOVECOMPLETE = 0x8004;
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
                                        // Drive or partition inserted or removed
                                        MyListView.ItemsSource = null;
                                        LoadData();
                                        //Debug.Write($"{i} Done.\n");
                                        //i++;
                                        break;
                                    }
                                case DBT_DEVICEREMOVECOMPLETE:
                                    {
                                        // Drive or partition inserted or removed
                                        MyListView.ItemsSource = null;
                                        LoadData();
                                        //Debug.Write($"{i} Done.\n");
                                        //i++;
                                        break;
                                    }
                                default:
                                    {
                                        // Drive or partition inserted or removed
                                        /*MyListView.ItemsSource = null;
                                        LoadData();
                                        Debug.Write($"{i} Done.\n");
                                        i++;*/
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

        public static bool IsAdmin()
        {
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        public static void SetDarkMode(WindowEx window)
        {
            var tl = new CommunityToolkit.WinUI.Helpers.ThemeListener();
            if (tl.CurrentTheme == Microsoft.UI.Xaml.ApplicationTheme.Dark)
            {
                IntPtr hWnd = WinRT.Interop.WindowNative.GetWindowHandle(window);
                int i = 1;
                DwmSetWindowAttribute(hWnd, 20, ref i, sizeof(int));
            }
        }

        public class DiskItem : Item
        {
            public string DriveLetter { get; set; }
            public string MediaType { get; set; }
        }

        [DllImport("dwmapi.dll", SetLastError = true)]
        public static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

        private void ListviewSelectionChange(object sender, SelectionChangedEventArgs args)
        {
            ReloadSelectedItem();
        }

        public void ReloadSelectedItem()
        {
            if (MyListView.SelectedItem != null)
            {
                DiskItem selectedItem = MyListView.SelectedItem as DiskItem;
                // Handle the selection change event

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
                DetailsBar.Title = selectedItem.Name;
                DetailsBar.Message = $"Media type: {selectedItem.MediaType}\nLast analyzed or optimized: {selI.Substring(0, selI.Length - 4)}\nCurrent status: NOT IMPLEMENTED";
            }
        }

        // Importing the necessary functions from kernel32.dll
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
                        mediaType = GetDriveTypeDescription(drive);

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

        private void LoadDriveTypes()
        {
            var drives = System.IO.DriveInfo.GetDrives();
            List<string> driveTypes = new List<string>();

            foreach (var drive in drives)
            {
                string driveTypeDescription = $"{drive.Name} - {GetDriveTypeDescription(drive.RootDirectory.FullName)}";
                driveTypes.Add(driveTypeDescription);
            }
        }

        private string GetDriveTypeDescription(string driveRoot)
        {
            DriveType driveType = GetDriveType(driveRoot);

            switch (driveType)
            {
                case DriveType.DRIVE_REMOVABLE:
                    return "Removable";
                case DriveType.DRIVE_FIXED:
                    return /*DiskInfo.IsSolidStateDrive(driveRoot) ? "Fixed (SSD)" : "Fixed (HDD)";*/ "Fixed (SSD/HDD)";
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

        private bool IsDriveSSD(string driveRoot)
        {
            const int FILE_SHARE_READ = 1;
            const int FILE_SHARE_WRITE = 2;
            const uint OPEN_EXISTING = 3;
            const uint IOCTL_STORAGE_QUERY_PROPERTY = 0x002D1400;
            const uint StorageDeviceProperty = 0;
            const uint PropertyStandardQuery = 0;

            IntPtr hDevice = IntPtr.Zero;
            try
            {
                hDevice = CreateFile(driveRoot, 0, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
                if (hDevice == InvalidHandleValue)
                    throw new Exception("CreateFile failed.");

                STORAGE_PROPERTY_QUERY query = new STORAGE_PROPERTY_QUERY
                {
                    PropertyId = StorageDeviceProperty,
                    QueryType = PropertyStandardQuery
                };

                int size = Marshal.SizeOf(typeof(STORAGE_DESCRIPTOR_HEADER)) + Marshal.SizeOf(typeof(STORAGE_DEVICE_DESCRIPTOR));
                IntPtr ptr = Marshal.AllocHGlobal(size);
                try
                {
                    uint bytesReturned;
                    if (!DeviceIoControl(hDevice, IOCTL_STORAGE_QUERY_PROPERTY, 0, (uint)Marshal.SizeOf(query), ptr, (uint)size, out bytesReturned, IntPtr.Zero))
                        throw new Exception("DeviceIoControl failed.");

                    STORAGE_DEVICE_DESCRIPTOR descriptor = (STORAGE_DEVICE_DESCRIPTOR)Marshal.PtrToStructure(ptr, typeof(STORAGE_DEVICE_DESCRIPTOR));

                    // Check if the device is an SSD
                    return ((descriptor.DeviceType & FILE_DEVICE_SSD) != 0);
                }
                finally
                {
                    Marshal.FreeHGlobal(ptr);
                }
            }
            finally
            {
                if (hDevice != IntPtr.Zero && hDevice != InvalidHandleValue)
                    CloseHandle(hDevice);
            }
        }

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

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            string volume = (MyListView.SelectedItem as DiskItem).DriveLetter.ToString().Remove(2, 1);
            string arguments = $"dfrgui.exe \n{volume} /O /U"; // /O to optimize the drive

            try
            {
                var processInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = false,
                    Verb = "runas"  // Run as administrator
                };

                var process = new Process
                {
                    StartInfo = processInfo,
                    EnableRaisingEvents = true
                };

                process.OutputDataReceived += (sender, e) => ParseOutput(e.Data);
                process.ErrorDataReceived += (sender, e) => ParseOutput(e.Data);

                try
                {
                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();

                    await Task.Run(() => process.WaitForExit());

                    ShowMessage("Defragmentation completed.");
                }
                catch (Exception ex)
                {
                    ParseOutput($"Error: {ex.Message}");
                }
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

        private void ParseOutput(string data)
        {
            // Example line to match: "50% complete."
            //UpdateProgress(data);
            if (data != null)
            {
                Debug.WriteLine(data);
                if (data.Contains("% complete"))
                {
                    var i = data.IndexOf("% complete");
                    UpdateProgress(data.Substring(i - 3, 3));
                }
                //ShowMessage(data);
            }
        }

        public void UpdateProgress(string progress)
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
}
