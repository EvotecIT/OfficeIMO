using Avalonia.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>
/// The key gestures Studio answers to. Commands in the catalog and window-level actions take their gestures from
/// here, so menus, command search, tooltips and the key handler always agree. "Primary" is Ctrl on Windows and
/// Linux and Command on macOS and iPadOS.
/// </summary>
internal static class StudioShortcuts {
    internal static bool UsesCommandKey => OperatingSystem.IsMacOS() || OperatingSystem.IsIOS();

    internal static KeyModifiers PrimaryModifier => UsesCommandKey ? KeyModifiers.Meta : KeyModifiers.Control;

    internal static KeyGesture Primary(Key key, KeyModifiers extra = KeyModifiers.None) => new(key, PrimaryModifier | extra);

    internal static KeyGesture Plain(Key key, KeyModifiers modifiers = KeyModifiers.None) => new(key, modifiers);

    // Window actions: not document commands, but shown in menus and answered by the window.
    internal static KeyGesture CommandSearch => Primary(Key.K);
    internal static KeyGesture CommandSearchAlternate => Primary(Key.P, KeyModifiers.Shift);
    internal static KeyGesture Find => Primary(Key.F);
    internal static KeyGesture CloseTab => Primary(Key.W);
    internal static KeyGesture ReopenTab => Primary(Key.T, KeyModifiers.Shift);
    internal static KeyGesture NextTab => new(Key.Tab, KeyModifiers.Control);
    internal static KeyGesture PreviousTab => new(Key.Tab, KeyModifiers.Control | KeyModifiers.Shift);
    internal static KeyGesture FullScreen => UsesCommandKey ? new(Key.F, KeyModifiers.Meta | KeyModifiers.Control) : Plain(Key.F11);
    internal static KeyGesture Quit => UsesCommandKey ? Primary(Key.Q) : OperatingSystem.IsWindows() ? new(Key.F4, KeyModifiers.Alt) : Primary(Key.Q);
    internal static KeyGesture ApplicationMenu => Plain(Key.F10);
    // macOS conventions with no Windows or Linux counterpart.
    internal static KeyGesture? ToggleSidebar => UsesCommandKey ? new(Key.S, KeyModifiers.Meta | KeyModifiers.Alt) : null;
    internal static KeyGesture? Minimize => UsesCommandKey ? Primary(Key.M) : null;

    internal static bool Matches(KeyGesture gesture, Key key, KeyModifiers modifiers) =>
        gesture.Key == key && gesture.KeyModifiers == modifiers;

    /// <summary>Formats a gesture the way the platform writes shortcuts: "Ctrl+Shift+S" or "⇧⌘S".</summary>
    internal static string Describe(KeyGesture gesture) {
        string key = DescribeKey(gesture.Key);
        if (UsesCommandKey) {
            string symbols = (gesture.KeyModifiers.HasFlag(KeyModifiers.Control) ? "⌃" : string.Empty) +
                             (gesture.KeyModifiers.HasFlag(KeyModifiers.Alt) ? "⌥" : string.Empty) +
                             (gesture.KeyModifiers.HasFlag(KeyModifiers.Shift) ? "⇧" : string.Empty) +
                             (gesture.KeyModifiers.HasFlag(KeyModifiers.Meta) ? "⌘" : string.Empty);
            return symbols + key;
        }
        var parts = new List<string>(4);
        if (gesture.KeyModifiers.HasFlag(KeyModifiers.Control)) parts.Add("Ctrl");
        if (gesture.KeyModifiers.HasFlag(KeyModifiers.Alt)) parts.Add("Alt");
        if (gesture.KeyModifiers.HasFlag(KeyModifiers.Shift)) parts.Add("Shift");
        if (gesture.KeyModifiers.HasFlag(KeyModifiers.Meta)) parts.Add("Win");
        parts.Add(key);
        return string.Join("+", parts);
    }

    private static string DescribeKey(Key key) => key switch {
        >= Key.D0 and <= Key.D9 => ((char)('0' + (key - Key.D0))).ToString(),
        >= Key.NumPad0 and <= Key.NumPad9 => ((char)('0' + (key - Key.NumPad0))).ToString(),
        Key.OemPlus or Key.Add => "+",
        Key.OemMinus or Key.Subtract => "−",
        Key.OemComma => ",",
        Key.OemPeriod => ".",
        Key.Up => "↑",
        Key.Down => "↓",
        Key.Left => "←",
        Key.Right => "→",
        _ => key.ToString()
    };
}
