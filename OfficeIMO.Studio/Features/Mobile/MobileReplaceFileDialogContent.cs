using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Mobile;

/// <summary>Confirms replacement after read-only destination selection, before the caller's final review.</summary>
internal sealed class MobileReplaceFileDialogContent : StudioDialogContent {
    internal MobileReplaceFileDialogContent(string name, IStudioLocalizer localizer) {
        Title = name;
        Width = 460;
        Height = 200;
        SizeToContent = SizeToContent.Height;
        var cancel = new Button { Content = localizer.Get("Common.Cancel"), IsCancel = true, Classes = { "tool" } };
        var save = new Button { Content = localizer.Get("Common.Save"), Classes = { "primary" } };
        cancel.Click += (_, _) => Close(false);
        save.Click += (_, _) => Close(true);
        Content = new StackPanel {
            Margin = new Thickness(24), Spacing = 16, Children = {
                new TextBlock { Text = localizer.Format("Mobile.ReplaceFileConfirm", name), TextWrapping = TextWrapping.Wrap },
                new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right,
                    Spacing = 8, Children = { cancel, save } }
            }
        };
        Opened += (_, _) => cancel.Focus();
    }
}
