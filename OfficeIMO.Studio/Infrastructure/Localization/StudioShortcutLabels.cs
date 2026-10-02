namespace OfficeIMO.Studio.Infrastructure.Localization;

/// <summary>Uses Apple modifier names in shared localized shortcut hints.</summary>
internal static class StudioShortcutLabels {
    internal static string Format(string text) => OperatingSystem.IsMacOS() || OperatingSystem.IsIOS()
        ? text.Replace("Control+", "⌘", StringComparison.Ordinal).Replace("Ctrl+Shift+", "⇧⌘", StringComparison.Ordinal)
            .Replace("Ctrl+", "⌘", StringComparison.Ordinal)
            .Replace("Ctrl-click", "⌘-click", StringComparison.Ordinal)
            .Replace("Alt+", "⌥", StringComparison.Ordinal)
        : text;
}
