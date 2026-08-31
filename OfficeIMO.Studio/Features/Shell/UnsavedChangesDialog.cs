using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;

namespace OfficeIMO.Studio.Features.Shell;

internal enum UnsavedChangesDecision {
    Cancel,
    Discard,
    Save
}

internal sealed class UnsavedChangesDialog : Window {
    internal UnsavedChangesDialog(string documentName) {
        Title = "Unsaved changes";
        Width = 430;
        Height = 190;
        CanResize = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;

        var save = new Button { Content = "Save", Classes = { "primary" }, MinWidth = 84 };
        var discard = new Button { Content = "Discard", Classes = { "tool" }, MinWidth = 84 };
        var cancel = new Button { Content = "Cancel", Classes = { "tool" }, MinWidth = 84 };
        save.Click += (_, _) => Close(UnsavedChangesDecision.Save);
        discard.Click += (_, _) => Close(UnsavedChangesDecision.Discard);
        cancel.Click += (_, _) => Close(UnsavedChangesDecision.Cancel);

        Content = new StackPanel {
            Margin = new Thickness(24),
            Spacing = 18,
            Children = {
                new TextBlock {
                    Text = $"Save changes to {documentName}?",
                    FontSize = 18,
                    FontWeight = Avalonia.Media.FontWeight.SemiBold
                },
                new TextBlock {
                    Text = "Your page edits have not been saved yet.",
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap
                },
                new StackPanel {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Spacing = 8,
                    Children = { cancel, discard, save }
                }
            }
        };
    }
}
