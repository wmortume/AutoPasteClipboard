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
using System.Windows.Input;
using Windows.ApplicationModel.DataTransfer;
using WindowsInput;
using WindowsInput.Native;
using Clipboard = Windows.ApplicationModel.DataTransfer.Clipboard; //prevents ambiguity

namespace AutoPasteClipboard
{
    public partial class MainWindow : Window
    {
        CancellationTokenSource CTS;
        readonly List<ClipboardHistoryItem> clipboardHistoryItems = new List<ClipboardHistoryItem>();
        readonly string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public MainWindow()
        {
            Clipboard.HistoryChanged += (object sender, ClipboardHistoryChangedEventArgs e) => { UpdateClipboardListView(); UndoChangesBtn.IsEnabled = ProfileComboBox.SelectedValue as string != "Default"; };
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

        private void UpdateClipboardOnProfileSelectionChanged(object sender, EventArgs e)
        {
            if (ProfileComboBox.Text != ProfileComboBox.SelectedValue as string) //makes sure it doesn't update if selection hasn't changed
            {
                UpdateClipboard();
            }
        }

        private void UndoClipboardChanges(object sender, RoutedEventArgs e)
        {
            UpdateClipboard();
        }

        private async void UpdateClipboard()
        {
            Indicator.IsBusy = true;
            if (ProfileComboBox.IsLoaded && ProfileComboBox.SelectedValue != null && ProfileComboBox.SelectedValue as string != "Default")
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

                    DelimiterComboBox.SelectedIndex = clipboardProfile.Delimeter;
                    DelayUpDownBox.Value = clipboardProfile.Delay;
                    ProfileNameTextBox.Text = ProfileComboBox.SelectedValue as string;
                }
            }
            UndoChangesBtn.IsEnabled = false;
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
                    Clipboard = clipboard,
                    Delimeter = DelimiterComboBox.SelectedIndex,
                    Delay = (int)DelayUpDownBox.Value
                };

                using (LiteDatabase db = new LiteDatabase(Path.Combine(documents, "Auto Paste Clipboard", "data.db")))
                {
                    ILiteCollection<ClipboardProfile> collection = db.GetCollection<ClipboardProfile>("clipboard");

                    if (collection.Exists(c => c.Profile == clipboardProfile.Profile))
                    {
                        MessageBox.Show("Clipboard has successfully been updated.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    collection.Upsert(clipboardProfile); //update or insert
                }

                UpdateProfileComboBox(false, clipboardProfile.Profile);
                UndoChangesBtn.IsEnabled = false;
            }
            else
            {
                MessageBox.Show("There's nothing to save. Try copying a text for it to appear on here.", "Alert", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }

        private void ClearClipboard(object sender, RoutedEventArgs e) { Clipboard.ClearHistory(); UndoChangesBtn.IsEnabled = false; }

        private void LoadHotkey()
        {
            using (LiteDatabase db = new LiteDatabase(Path.Combine(documents, "Auto Paste Clipboard", "data.db")))
            {
                ILiteCollection<Hotkey> collection = db.GetCollection<Hotkey>("hotkey");

                if (collection.Count() > 0)
                {
                    Hotkey hotkey = collection.FindById(collection.Min());
                    HkTextBox.Hotkey = hotkey;
                    try
                    {
                        HotkeyManager.Current.AddOrReplace("Paste", HkTextBox.Hotkey.Key, HkTextBox.Hotkey.Modifiers, AutoPaste);
                        HotkeyManager.Current.AddOrReplace("CancelPaste", Key.Delete, ModifierKeys.Control, CancelPaste);
                    }
                    catch (HotkeyAlreadyRegisteredException)
                    {
                        MessageBox.Show("Application is already running or hotkeys are set to be used by another one.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        Application.Current.Shutdown();
                    }
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
                    MessageBox.Show($"{hotkey} has been successfully set as the hotkey.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
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

        private async void AutoPaste(object sender, HotkeyEventArgs e)
        {
            try
            {
                CTS = new CancellationTokenSource();
                CancellationToken cancelToken = CTS.Token;

                if (WindowState != WindowState.Minimized)
                {
                    WindowState = WindowState.Minimized;
                    await Task.Delay(TimeSpan.FromSeconds(1), cancelToken);
                }

                InputSimulator input = new InputSimulator();

                for (int i = 0; i < clipboardHistoryItems.Count; i++)
                {
                    Clipboard.SetHistoryItemAsContent(clipboardHistoryItems[i]);
                    await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                    input.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);

                    switch (DelimiterComboBox.SelectedIndex)
                    {
                        case 1:
                            await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                            input.Keyboard.KeyPress(VirtualKeyCode.TAB);
                            break;
                        case 2:
                            await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                            input.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                            break;
                        case 3:
                            await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                            input.Keyboard.KeyPress(VirtualKeyCode.SPACE);
                            break;
                        case 4:
                            await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                            input.Keyboard.KeyPress(VirtualKeyCode.OEM_COMMA);
                            break;
                        case 5:
                            await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                            input.Keyboard.KeyPress(VirtualKeyCode.OEM_PERIOD);
                            break;
                        case 6:
                            await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                            input.Keyboard.KeyPress(VirtualKeyCode.OEM_COMMA);
                            await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                            input.Keyboard.KeyPress(VirtualKeyCode.SPACE);
                            break;
                        case 7:
                            await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                            input.Keyboard.KeyPress(VirtualKeyCode.OEM_PERIOD);
                            await Task.Delay((int)DelayUpDownBox.Value, cancelToken);
                            input.Keyboard.KeyPress(VirtualKeyCode.SPACE);
                            break;
                        default:
                            break;
                    }
                }
                e.Handled = true;
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) //handles async exceptions
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (CTS != null)
                {
                    CTS.Dispose();
                    CTS = null;
                }
            }
        }

        private void CancelPaste(object sender, HotkeyEventArgs e)
        {
            if (CTS != null)
            {
                CTS.Cancel();
            }
            e.Handled = true;
        }
    }
}
