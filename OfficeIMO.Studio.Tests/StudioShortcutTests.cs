using Avalonia.Input;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioShortcutTests {
    private static KeyModifiers Primary => StudioShortcuts.PrimaryModifier;

    [Fact]
    public void Catalog_gestures_drive_labels_and_respect_text_fields() {
        using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        StudioCommandCatalog commands = document.Commands;

        Assert.Same(commands["Save"], commands.FindByGesture(Key.S, Primary, textEntryFocused: false));
        Assert.Same(commands["SaveAs"], commands.FindByGesture(Key.S, Primary | KeyModifiers.Shift, textEntryFocused: false));
        // File commands still work from a text field; editing and zoom keys belong to the field.
        Assert.Same(commands["Open"], commands.FindByGesture(Key.O, Primary, textEntryFocused: true));
        Assert.Null(commands.FindByGesture(Key.Z, Primary, textEntryFocused: true));
        Assert.Same(commands["Undo"], commands.FindByGesture(Key.Z, Primary, textEntryFocused: false));
        Assert.Same(commands["ZoomIn"], commands.FindByGesture(Key.Add, Primary, textEntryFocused: false));
        Assert.Same(commands["FitPage"], commands.FindByGesture(Key.NumPad0, Primary, textEntryFocused: false));
        Assert.Null(commands.FindByGesture(Key.S, KeyModifiers.Alt, textEntryFocused: false));

        // Every command with gestures shows the first one, so menus, search and tooltips agree.
        Assert.All(commands.Items.Where(item => item.Gestures.Count > 0),
            item => Assert.Equal(StudioShortcuts.Describe(item.Gestures[0]), item.Shortcut));
        Assert.Equal(commands["Save"].Title + " (" + commands["Save"].Shortcut + ")", commands["Save"].ToolTipText);

        // No two commands claim the same gesture.
        var claimed = commands.Items.SelectMany(item => item.Gestures.Select(gesture => (gesture.Key, gesture.KeyModifiers))).ToArray();
        Assert.Equal(claimed.Length, claimed.Distinct().Count());
    }

    [Fact]
    public void Redo_follows_the_platform_convention() {
        using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        StudioCommandItem redo = document.Commands["Redo"];
        Assert.Same(redo, document.Commands.FindByGesture(Key.Z, Primary | KeyModifiers.Shift, textEntryFocused: false));
        if (!StudioShortcuts.UsesCommandKey) {
            Assert.Same(redo, document.Commands.FindByGesture(Key.Y, KeyModifiers.Control, textEntryFocused: false));
            Assert.Equal("Ctrl+Y", redo.Shortcut);
        }
    }

    [Fact]
    public void Toolbar_tooltips_explain_why_a_command_is_unavailable() {
        using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        StudioCommandItem watermark = document.Commands["Watermark"];
        Assert.False(watermark.CanExecute(null));
        Assert.False(string.IsNullOrWhiteSpace(watermark.UnavailableReason));
        Assert.Contains(watermark.Title, watermark.StatusToolTip, StringComparison.Ordinal);
        Assert.EndsWith(watermark.UnavailableReason!, watermark.StatusToolTip, StringComparison.Ordinal);
    }

    [Fact]
    public void Gestures_are_written_the_way_the_platform_shows_them() {
        string saveCopy = StudioShortcuts.Describe(StudioShortcuts.Primary(Key.S, KeyModifiers.Shift));
        string zoom = StudioShortcuts.Describe(StudioShortcuts.Primary(Key.OemPlus));
        if (StudioShortcuts.UsesCommandKey) {
            Assert.Equal("⇧⌘S", saveCopy);
            Assert.Equal("⌘+", zoom);
        } else {
            Assert.Equal("Ctrl+Shift+S", saveCopy);
            Assert.Equal("Ctrl++", zoom);
            Assert.Equal("Ctrl+,", StudioShortcuts.Describe(StudioShortcuts.Primary(Key.OemComma)));
            Assert.Equal("F9", StudioShortcuts.Describe(StudioShortcuts.Plain(Key.F9)));
        }
    }
}
