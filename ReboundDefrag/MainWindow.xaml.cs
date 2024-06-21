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
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;

namespace ReboundDefrag
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public MainWindow()
        {
            this.InitializeComponent();
            SystemBackdrop = new MicaBackdrop
            {
                Kind = MicaKind.BaseAlt,
            };

            ExtendsContentIntoTitleBar = true;
            Title = "Rebound Defragmentation Tool";

            LoadData();
        }

        private void ListviewSelectionChange(object sender, SelectionChangedEventArgs args)
        {
            if (MyListView.SelectedItem != null)
            {
                string selectedItem = MyListView.SelectedItem.ToString();
                // Handle the selection change event
            }
        }

        private void LoadData()
        {
            List<Item> items = new List<Item>
            {
                new Item { Name = "Item 1", ImagePath = "ms-appx:///Assets/Icon1.png" },
                new Item { Name = "Item 2", ImagePath = "ms-appx:///Assets/Icon1.png" },
                new Item { Name = "Item 3", ImagePath = "ms-appx:///Assets/Icon1.png" },
                new Item { Name = "Item 4", ImagePath = "ms-appx:///Assets/Icon1.png" },
                new Item { Name = "Item 5", ImagePath = "ms-appx:///Assets/Icon1.png" }
            };
            MyListView.ItemsSource = items;
        }
    }
}
