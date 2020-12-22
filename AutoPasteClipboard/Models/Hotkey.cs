using System.Text;
using System.Windows.Input;

namespace AutoPasteClipboard
{
    public class Hotkey
    {
        public Key Key { get; }

        public ModifierKeys Modifiers { get; }

        public Hotkey(Key key, ModifierKeys modifiers = ModifierKeys.None)
        {
            Key = key;
            Modifiers = modifiers;
        }

        public override string ToString()
        {
            if (Key == Key.None && Modifiers == ModifierKeys.None)
                return "< Shortcut >";

            StringBuilder buffer = new StringBuilder();

            if (Modifiers.HasFlag(ModifierKeys.Control))
                buffer.Append("Ctrl + ");
            if (Modifiers.HasFlag(ModifierKeys.Shift))
                buffer.Append("Shift + ");
            if (Modifiers.HasFlag(ModifierKeys.Alt))
                buffer.Append("Alt + ");

            switch (Key)
            {
                case Key.D0:
                case Key.D1:
                case Key.D2:
                case Key.D3:
                case Key.D4:
                case Key.D5:
                case Key.D6:
                case Key.D7:
                case Key.D8:
                case Key.D9:
                    buffer.Append((int)KeyInterop.KeyFromVirtualKey((int)Key) - 20).ToString();
                    break;
                case Key.Return:
                    buffer.Append("Enter");
                    break;
                case Key.OemComma:
                    buffer.Append(',');
                    break;
                case Key.OemPeriod:
                    buffer.Append('.');
                    break;
                case Key.OemQuestion:
                    buffer.Append('/');
                    break;
                case Key.Oem1:
                    buffer.Append(';');
                    break;
                case Key.OemQuotes:
                    buffer.Append('\'');
                    break;
                case Key.OemOpenBrackets:
                    buffer.Append('[');
                    break;
                case Key.Oem6:
                    buffer.Append(']');
                    break;
                case Key.Oem5:
                    buffer.Append('\\');
                    break;
                case Key.OemMinus:
                    buffer.Append('-');
                    break;
                case Key.OemPlus:
                    buffer.Append('+');
                    break;
                default:
                    buffer.Append(Key);
                    break;
            }

            return buffer.ToString();
        }
    }
}
