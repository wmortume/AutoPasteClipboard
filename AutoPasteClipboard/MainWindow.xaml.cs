using AutoPasteClipboard.Models;
using LiteDB;
using NHotkey;
using NHotkey.Wpf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Windows.ApplicationModel.DataTransfer;
using WindowsInput;
using WindowsInput.Native;
using Clipboard = Windows.ApplicationModel.DataTransfer.Clipboard; //prevents ambiguity

namespace AutoPasteClipboard
{
    public partial class MainWindow : Window
    {
        readonly List<ClipboardHistoryItem> clipboardHistoryItems = new List<ClipboardHistoryItem>();
        readonly string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);         

        public MainWindow()
        {
            Mutex mutex = new Mutex(true, "AutoPasteClipboard", out bool newInstance);

            if (!newInstance)
            {
                MessageBox.Show("Application is already running.");
                Application.Current.Shutdown();
            }

            Clipboard.HistoryChanged += (object sender, ClipboardHistoryChangedEventArgs e) => { UpdateClipboardListView(); };
            Directory.CreateDirectory(Path.Combine(documents, "Auto Paste Clipboard"));
            InitializeComponent();
            UpdateClipboardListView();
            UpdateProfileComboBox(true);
            LoadHotkey();
        }

        private async void UpdateClipboardListView()
        {
            clipboardHistoryItems.Clear();
            List<ListViewItem> clipboardTexts = new List<ListViewItem>();
            ClipboardHistoryItemsResult items = await Clipboard.GetHistoryItemsAsync();

            foreach (ClipboardHistoryItem item in items.Items.Reverse())
            {
                if (item.Content.Contains(DataFormats.Text))
                {
                    string data = await item.Content.GetTextAsync();
                    clipboardTexts.Add(new ListViewItem() { Content = data });
                    clipboardHistoryItems.Add(item);
                }
            }

            ClipboardListView.ItemsSource = clipboardTexts;

            foreach (object item in ClipboardListView.Items)
            {
                ((ListViewItem)item).Style = (Style)FindResource(resourceKey: "ListViewItemStyle");
            }
        }

        private async void UpdateClipboardOnProfileSelectionChanged(object sender, EventArgs e)
        {
            Indicator.IsBusy = true;
            if (ProfileComboBox.IsLoaded && ProfileComboBox.SelectedValue != null && ProfileComboBox.Text != ProfileComboBox.SelectedValue as string && ProfileComboBox.SelectedValue as string != "Default")
            {
                if (ProfileComboBox.Items.Contains("Default"))
                {
                    UpdateProfileComboBox(false, ProfileComboBox.SelectedValue as string);
                }

                Clipboard.ClearHistory();

                using (LiteDatabase db = new LiteDatabase($"Filename={Path.Combine(documents, "Auto Paste Clipboard", "data.db")}; Connection=shared"))
                {
                    ILiteCollection<ClipboardProfile> collection = db.GetCollection<ClipboardProfile>("clipboard");

                    ClipboardProfile clipboardProfile = collection.FindOne(x => x.Profile == ProfileComboBox.SelectedValue as string);

                    foreach (string item in clipboardProfile.Clipboard)
                    {
                        DataPackage data = new DataPackage();
                        data.SetText(item);
                        Clipboard.SetContent(data);
                        await Task.Delay(450);
                    }
                }
            }
            Indicator.IsBusy = false;
        }

        private void UpdateProfileComboBox(bool addDefaultProfile, string currentDropDownProfile = null)
        {
            using (LiteDatabase db = new LiteDatabase($"Filename={Path.Combine(documents, "Auto Paste Clipboard", "data.db")}; Connection=shared"))
            {
                ILiteCollection<ClipboardProfile> collection = db.GetCollection<ClipboardProfile>("clipboard");
                IEnumerable<string> profiles = collection.FindAll().Select(x => x.Profile);
                IEnumerable<string> totalProfiles = addDefaultProfile ? profiles.Prepend("Default") : profiles;
                ProfileComboBox.ItemsSource = totalProfiles;

                if (currentDropDownProfile != null)
                {
                    int index = ProfileComboBox.Items.IndexOf(currentDropDownProfile);
                    ProfileComboBox.SelectedIndex = index;
                }
                else
                {
                    ProfileComboBox.SelectedIndex = 0;
                }
            }
        }

        private void DeleteProfile(object sender, RoutedEventArgs e)
        {
            if (ProfileComboBox.Text != "Default")
            {
                using (LiteDatabase db = new LiteDatabase(Path.Combine(documents, "Auto Paste Clipboard", "data.db")))
                {
                    ILiteCollection<ClipboardProfile> collection = db.GetCollection<ClipboardProfile>("clipboard");
                    ClipboardProfile clipboardProfile = collection.FindOne(x => x.Profile == ProfileComboBox.Text);
                    collection.Delete(clipboardProfile.Profile);
                }

                if (ProfileComboBox.Items.Count == 1)
                {
                    UpdateProfileComboBox(true);
                }
                else
                {
                    UpdateProfileComboBox(false);
                }
            }
        }

        private void SaveClipboard(object sender, RoutedEventArgs e)
        {
            List<string> clipboard = new List<string>();

            foreach (ListViewItem item in ClipboardListView.Items)
            {
                clipboard.Add(item.Content as string);
            }

            if (ClipboardListView.Items.Count != 0)
            {
                ClipboardProfile clipboardProfile = new ClipboardProfile
                {
                    Profile = ProfileNameTextBox.Text,
                    Clipboard = clipboard
                };

                using (LiteDatabase db = new LiteDatabase(Path.Combine(documents, "Auto Paste Clipboard", "data.db")))
                {
                    ILiteCollection<ClipboardProfile> collection = db.GetCollection<ClipboardProfile>("clipboard");

                    collection.Upsert(clipboardProfile); //update or insert
                }

                UpdateProfileComboBox(false, clipboardProfile.Profile);
            }
        }

        private void ClearClipboard(object sender, RoutedEventArgs e) { Clipboard.ClearHistory(); }

        private void LoadHotkey()
        {
            using (LiteDatabase db = new LiteDatabase(Path.Combine(documents, "Auto Paste Clipboard", "data.db")))
            {
                ILiteCollection<Hotkey> collection = db.GetCollection<Hotkey>("hotkey");

                if (collection.Count() > 0)
                {
                    Hotkey hotkey = collection.FindById(collection.Min());
                    HkTextBox.Hotkey = hotkey;
                    HotkeyManager.Current.AddOrReplace("Paste", HkTextBox.Hotkey.Key, HkTextBox.Hotkey.Modifiers, AutoPaste);
                }
            }
        }

        private void SetHotkey(object sender, RoutedEventArgs e)
        {
            if (Clipboard.IsHistoryEnabled() && HkTextBox.Hotkey != null)
            {
                string[] modifiers = HkTextBox.Hotkey.Modifiers.ToString().Split(',').OrderByDescending(d => d == "Control").ThenBy(a => a).ToArray(); //Following Ctrl + Shift + Alt Ordering Convention
                string hotkey = $"{ (modifiers[0] == "None" ? "" : string.Join(" + ", modifiers) + " + ") } {HkTextBox.Hotkey.Key}";

                try
                {
                    using (LiteDatabase db = new LiteDatabase(Path.Combine(documents, "Auto Paste Clipboard", "data.db")))
                    {
                        ILiteCollection<Hotkey> collection = db.GetCollection<Hotkey>("hotkey");

                        if (collection.Count() > 0)
                        {
                            collection.DeleteAll();
                        }

                        collection.Insert(HkTextBox.Hotkey);
                    }

                    HotkeyManager.Current.AddOrReplace("Paste", HkTextBox.Hotkey.Key, HkTextBox.Hotkey.Modifiers, AutoPaste);
                    MessageBox.Show($"{hotkey} has been successfully set as the hotkey.");
                    WindowState = WindowState.Minimized;
                }
                catch (HotkeyAlreadyRegisteredException)
                {
                    MessageBox.Show($"{hotkey} is being used by another application.", "Alert", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            else
            {
                MessageBox.Show("You must enable clipboard history on windows settings and set a hotkey.", "Alert", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        private void AutoPaste(object sender, HotkeyEventArgs e)
        {
            if (WindowState != WindowState.Minimized)
            {
                WindowState = WindowState.Minimized;
                Thread.Sleep(TimeSpan.FromSeconds(1));
            }
            InputSimulator input = new InputSimulator();

            for (int i = 0; i < clipboardHistoryItems.Count; i++)
            {
                Clipboard.SetHistoryItemAsContent(clipboardHistoryItems[i]);
                input.Keyboard.Sleep((int)Delay.Value).ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);

                switch (DelimiterComboBox.SelectedIndex)
                {
                    case 1:
                        input.Keyboard.Sleep((int)Delay.Value).KeyPress(VirtualKeyCode.TAB);
                        break;
                    case 2:
                        input.Keyboard.Sleep((int)Delay.Value).KeyPress(VirtualKeyCode.RETURN);
                        break;
                    case 3:
                        input.Keyboard.Sleep((int)Delay.Value).KeyPress(VirtualKeyCode.SPACE);
                        break;
                    case 4:
                        input.Keyboard.Sleep((int)Delay.Value).KeyPress(VirtualKeyCode.OEM_COMMA);
                        break;
                    case 5:
                        input.Keyboard.Sleep((int)Delay.Value).KeyPress(VirtualKeyCode.OEM_PERIOD);
                        break;
                    case 6:
                        input.Keyboard.Sleep((int)Delay.Value).KeyPress(VirtualKeyCode.OEM_COMMA).Sleep((int)Delay.Value).KeyPress(VirtualKeyCode.SPACE);
                        break;
                    case 7:
                        input.Keyboard.Sleep((int)Delay.Value).KeyPress(VirtualKeyCode.OEM_PERIOD).Sleep((int)Delay.Value).KeyPress(VirtualKeyCode.SPACE);
                        break;
                    default:
                        break;
                }
            }

            e.Handled = true;
        }
    }
}
