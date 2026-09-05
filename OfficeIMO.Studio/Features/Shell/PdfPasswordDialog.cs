using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed class PdfPasswordDialog : Window {
    internal PdfPasswordDialog(string documentName, bool invalidPassword, IStudioLocalizer? localizer = null) {
        localizer ??= StudioLocalization.Current;
        Title = localizer.Get("Dialog.PasswordRequired");
        Width = 430;
        Height = invalidPassword ? 238 : 215;
        CanResize = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;

        var password = new TextBox {
            PasswordChar = '●',
            PlaceholderText = localizer.Get("Dialog.DocumentPassword")
        };
        var open = new Button { Content = localizer.Get("Common.Open"), Classes = { "primary" }, MinWidth = 84, IsDefault = true };
        var cancel = new Button { Content = localizer.Get("Common.Cancel"), Classes = { "tool" }, MinWidth = 84, IsCancel = true };
        open.Click += (_, _) => Close(password.Text ?? string.Empty);
        cancel.Click += (_, _) => Close(null);

        var content = new StackPanel {
            Margin = new Thickness(24),
            Spacing = 12,
            Children = {
                new TextBlock {
                    Text = documentName,
                    FontSize = 18,
                    FontWeight = FontWeight.SemiBold,
                    TextTrimming = TextTrimming.CharacterEllipsis
                },
                new TextBlock {
                    Text = localizer.Get("Dialog.EnterPassword"),
                    TextWrapping = TextWrapping.Wrap
                }
            }
        };
        if (invalidPassword) {
            content.Children.Add(new TextBlock {
                Text = localizer.Get("Dialog.InvalidPassword"),
                Foreground = Brushes.IndianRed,
                TextWrapping = TextWrapping.Wrap
            });
        }
        content.Children.Add(password);
        content.Children.Add(new StackPanel {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing = 8,
            Children = { cancel, open }
        });
        Content = content;
        Opened += (_, _) => password.Focus();
    }
}
