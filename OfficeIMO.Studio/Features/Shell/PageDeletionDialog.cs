using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed class PageDeletionDialog : Window {
    internal PageDeletionDialog(int pageCount, IStudioLocalizer? localizer = null) {
        if (pageCount < 1) throw new ArgumentOutOfRangeException(nameof(pageCount));
        localizer ??= StudioLocalization.Current;

        Title = pageCount == 1 ? localizer.Get("Dialog.DeletePageTitle") : localizer.Format("Dialog.DeletePagesTitle", pageCount);
        Width = 440;
        Height = 200;
        CanResize = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;

        var delete = new Button {
            Content = pageCount == 1 ? localizer.Get("Dialog.DeletePage") : localizer.Get("Dialog.DeletePages"),
            Background = Avalonia.Media.Brushes.Firebrick,
            Foreground = Avalonia.Media.Brushes.White,
            MinWidth = 104
        };
        var cancel = new Button { Content = localizer.Get("Common.Cancel"), Classes = { "tool" }, MinWidth = 84, IsCancel = true };
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
                    Text = localizer.Get("Dialog.DeletePagesDescription"),
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
        Opened += (_, _) => cancel.Focus();
    }
}
