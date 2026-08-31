using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed class PageDeletionDialog : Window {
    internal PageDeletionDialog(int pageCount) {
        if (pageCount < 1) throw new ArgumentOutOfRangeException(nameof(pageCount));

        Title = pageCount == 1 ? "Delete page?" : $"Delete {pageCount} pages?";
        Width = 440;
        Height = 200;
        CanResize = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;

        var delete = new Button {
            Content = pageCount == 1 ? "Delete page" : "Delete pages",
            Background = Avalonia.Media.Brushes.Firebrick,
            Foreground = Avalonia.Media.Brushes.White,
            MinWidth = 104
        };
        var cancel = new Button { Content = "Cancel", Classes = { "tool" }, MinWidth = 84 };
        delete.Click += (_, _) => Close(true);
        cancel.Click += (_, _) => Close(false);

        Content = new StackPanel {
            Margin = new Thickness(24),
            Spacing = 18,
            Children = {
                new TextBlock {
                    Text = Title,
                    FontSize = 18,
                    FontWeight = Avalonia.Media.FontWeight.SemiBold
                },
                new TextBlock {
                    Text = "The pages will be removed from this working copy. You can still use Undo before saving.",
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap
                },
                new StackPanel {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Spacing = 8,
                    Children = { cancel, delete }
                }
            }
        };
    }
}
