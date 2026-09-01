using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed class PdfPasswordDialog : Window {
    internal PdfPasswordDialog(string documentName, bool invalidPassword) {
        Title = "Password required";
        Width = 430;
        Height = invalidPassword ? 238 : 215;
        CanResize = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;

        var password = new TextBox {
            PasswordChar = '●',
            PlaceholderText = "Document password"
        };
        var open = new Button { Content = "Open", Classes = { "primary" }, MinWidth = 84, IsDefault = true };
        var cancel = new Button { Content = "Cancel", Classes = { "tool" }, MinWidth = 84, IsCancel = true };
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
                    Text = "Enter the password to open this PDF.",
                    TextWrapping = TextWrapping.Wrap
                }
            }
        };
        if (invalidPassword) {
            content.Children.Add(new TextBlock {
                Text = "That password did not open the document. Try again.",
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
