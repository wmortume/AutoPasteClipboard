using NHotkey;
using NHotkey.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows;
using Windows.ApplicationModel.DataTransfer;
using WindowsInput;
using WindowsInput.Native;
using Clipboard = Windows.ApplicationModel.DataTransfer.Clipboard;

namespace AutoPasteClipboard
{
    public partial class MainWindow : Window
    {
        readonly List<ClipboardHistoryItem> clipboardHistoryItems = new List<ClipboardHistoryItem>();
        public MainWindow()
        {
            Clipboard.HistoryChanged += Clipboard_HistoryChanged;
            InitializeComponent();
            UpdateClipboardListView();
        }


        private void Clipboard_HistoryChanged(object sender, ClipboardHistoryChangedEventArgs e)
        {
            UpdateClipboardListView();
        }

        private async void UpdateClipboardListView()
        {
            List<string> clipboardTexts = new List<string>();
            clipboardHistoryItems.Clear();

            ClipboardHistoryItemsResult items = await Clipboard.GetHistoryItemsAsync();
            foreach (ClipboardHistoryItem item in items.Items)
            {
                if (item.Content.Contains(DataFormats.Text))
                {
                    string data = await item.Content.GetTextAsync();
                    clipboardTexts.Add(data);
                    clipboardHistoryItems.Add(item);
                }
            }

            clipboardTexts.Reverse();
            clipboardHistoryItems.Reverse();

            ClipboardListView.ItemsSource = clipboardTexts;
        }

        private void SetHotkey(object sender, RoutedEventArgs e)
        {
            if (Clipboard.IsHistoryEnabled())
            {
                string[] modifiers = HkTextBox.Hotkey.Modifiers.ToString().Split(',').OrderByDescending(d => d == "Control").ThenBy(a => a).ToArray(); //Following Ctrl + Shift + Alt Ordering Convention
                string hotkey = $"{ (modifiers[0] == "None" ? "" : string.Join(" + ", modifiers) + " + ") } {HkTextBox.Hotkey.Key}";
                try
                {
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
                MessageBox.Show("You must enable clipboard history on windows settings.", "Alert", MessageBoxButton.OK, MessageBoxImage.Exclamation);
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

                switch (Delimiter.SelectedIndex)
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
