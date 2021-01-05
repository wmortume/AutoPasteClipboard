using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace AutoPasteClipboard.Controls
{

    public partial class HotkeyEditorControl : UserControl
    {
        public static readonly DependencyProperty HotkeyProperty =
            DependencyProperty.Register(nameof(Hotkey), typeof(Hotkey),
                typeof(HotkeyEditorControl),
                new FrameworkPropertyMetadata(default(Hotkey),
                    FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));

        public Hotkey Hotkey
        {
            get => (Hotkey)GetValue(HotkeyProperty);
            set => SetValue(HotkeyProperty, value);
        }

        public HotkeyEditorControl()
        {
            InitializeComponent();
            DataContext = this;
        }

        private static bool HasKeyChar(Key key) =>
            new[]
            {
                Key.LWin, Key.RWin, Key.OemQuestion, Key.OemQuotes, Key.OemPlus, Key.OemOpenBrackets, Key.OemCloseBrackets, Key.OemMinus, Key.DeadCharProcessed,
                Key.Oem1, Key.Oem5, Key.Oem7, Key.OemPeriod, Key.OemComma, Key.Add, Key.Divide, Key.Multiply, Key.Subtract, Key.Oem102, Key.Decimal
            }
        .Contains(key);

        private void HotkeyTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // Don't let the event pass further
            // because we don't want standard textbox shortcuts working
            e.Handled = true;

            // Get modifiers and key data
            ModifierKeys modifiers = Keyboard.Modifiers;
            Key key = e.Key;

            // If nothing was pressed - return
            if (key == Key.None)
                return;

            // If Alt is used as modifier - the key needs to be extracted from SystemKey
            if (key == Key.System)
                key = e.SystemKey;

            // Pressing delete, backspace or escape without modifiers clears the current value
            if (modifiers == ModifierKeys.None && (key == Key.Delete || key == Key.Back || key == Key.Escape))
            {
                Hotkey = null;
                return;
            }

            // If no actual key was pressed - return
            if (new[] { Key.LeftCtrl, Key.RightCtrl, Key.LeftAlt, Key.RightAlt, Key.LeftShift, Key.RightShift, Key.Clear, Key.OemClear, Key.Apps }.Contains(key))
            {
                return;
            }

            // If Enter/Space/Tab is pressed without modifiers - return            
            if (modifiers == ModifierKeys.None && (key == Key.Enter || key == Key.Space || key == Key.Tab))
            {
                return;
            }

            // If key has a character and pressed without modifiers or only with Shift - return            
            if (HasKeyChar(key) && (modifiers == ModifierKeys.None || modifiers == ModifierKeys.Shift))
            {
                return;
            }

            // Set value
            Hotkey = new Hotkey() { Key = key, Modifiers = modifiers};
        }
    }
}
