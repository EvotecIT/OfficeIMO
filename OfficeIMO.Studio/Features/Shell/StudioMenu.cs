using System.Windows.Input;
using Avalonia.Controls;
using Avalonia.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>
/// One menu definition for every platform. Windows and Linux render it as the in-window application menu; macOS renders
/// it as the native menu bar. Entries carry the shared commands, so availability, titles and shortcuts stay identical.
/// </summary>
internal abstract record StudioMenuEntry;

internal sealed record StudioMenuGroup(string Id, string Title, IReadOnlyList<StudioMenuEntry> Items) : StudioMenuEntry;

internal sealed record StudioMenuCommand(string Id, string Title, ICommand Command, KeyGesture? Gesture = null) : StudioMenuEntry {
    internal object? Parameter { get; init; }
    /// <summary>Shown on a disabled command so the menu explains why it is unavailable.</summary>
    internal string? ToolTip { get; init; }
    /// <summary>Set for radio entries such as the open documents or the theme; null for ordinary commands.</summary>
    internal bool? IsChecked { get; init; }
}

internal sealed record StudioMenuSeparator : StudioMenuEntry {
    internal static StudioMenuSeparator Instance { get; } = new();
}

internal static class StudioMenuRenderer {
    /// <summary>Renders an entry for the in-window menu; a top-level group can be given a header with its access key.</summary>
    internal static object ToMenuItem(StudioMenuEntry entry, string? header = null) {
        switch (entry) {
            case StudioMenuGroup group:
                var menu = new MenuItem { Header = header ?? group.Title, Tag = group.Id, IsEnabled = group.Items.Count > 0 };
                foreach (StudioMenuEntry item in group.Items) menu.Items.Add(ToMenuItem(item));
                return menu;
            case StudioMenuCommand command:
                var menuItem = new MenuItem {
                    Header = command.Title, Tag = command.Id, Command = command.Command, CommandParameter = command.Parameter,
                    InputGesture = command.Gesture,
                    ToggleType = command.IsChecked is null ? MenuItemToggleType.None : MenuItemToggleType.Radio,
                    IsChecked = command.IsChecked == true
                };
                if (command.ToolTip is { Length: > 0 } tip) {
                    Avalonia.Controls.ToolTip.SetTip(menuItem, tip);
                    Avalonia.Controls.ToolTip.SetShowOnDisabled(menuItem, true);
                }
                return menuItem;
            default:
                return new Separator();
        }
    }

    /// <summary>Renders an entry for the macOS menu bar.</summary>
    internal static NativeMenuItemBase ToNative(StudioMenuEntry entry) {
        switch (entry) {
            case StudioMenuGroup group:
                var menu = new NativeMenu();
                foreach (StudioMenuEntry item in group.Items) menu.Items.Add(ToNative(item));
                return new NativeMenuItem(group.Title) { Menu = menu, IsEnabled = group.Items.Count > 0 };
            case StudioMenuCommand command:
                return new NativeMenuItem(command.Title) {
                    Command = command.Command, CommandParameter = command.Parameter, Gesture = command.Gesture,
                    ToggleType = command.IsChecked is null ? MenuItemToggleType.None : MenuItemToggleType.Radio,
                    IsChecked = command.IsChecked == true
                };
            default:
                return new NativeMenuItemSeparator();
        }
    }
}
