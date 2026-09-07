using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Explains the write guarantees of a stream-only destination before the user publishes edits.</summary>
internal sealed class ProviderSaveDialog : Window {
    internal ProviderSaveDialog(string name, IStudioLocalizer localizer, bool workflowOutput = false, bool folderOutput = false) {
        Title = localizer.Get("Dialog.ProviderSaveTitle");
        Width = 520;
        SizeToContent = SizeToContent.Height;
        CanResize = false;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        var cancel = new Button { Content = localizer.Get("Common.Cancel"), MinWidth = 92, IsCancel = true, Classes = { "tool" } };
        var save = new Button { Content = localizer.Get("Common.Save"), MinWidth = 92, Classes = { "primary" } };
        cancel.Click += (_, _) => Close(false);
        save.Click += (_, _) => Close(true);
        Content = new StackPanel {
            Margin = new Thickness(24), Spacing = 18,
            Children = {
                new TextBlock { Text = localizer.Format("Dialog.ProviderSaveName", name), FontSize = 18,
                    FontWeight = FontWeight.SemiBold, TextWrapping = TextWrapping.Wrap,
                    MaxLines = 3, TextTrimming = TextTrimming.CharacterEllipsis },
                new TextBlock { Text = localizer.Get(folderOutput ? "Dialog.ProviderFolderDescription" : workflowOutput ? "Dialog.ProviderWorkflowDescription" : "Dialog.ProviderSaveDescription"), TextWrapping = TextWrapping.Wrap },
                new TextBlock { Text = localizer.Get("Dialog.ProviderWorkflowRecovery"), TextWrapping = TextWrapping.Wrap,
                    IsVisible = workflowOutput },
                new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right,
                    Spacing = 8, Children = { cancel, save } }
            }
        };
        Opened += (_, _) => cancel.Focus();
    }
}
