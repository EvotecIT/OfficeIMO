using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Mobile;

/// <summary>Collects the name before choosing a Files folder; neither step publishes a placeholder.</summary>
internal sealed class MobileFileNameDialogContent : StudioDialogContent {
    internal MobileFileNameDialogContent(string title, string suggestedName, StudioFileType type, IStudioLocalizer localizer) {
        Title = title;
        Width = 460;
        SizeToContent = SizeToContent.Height;
        var name = new TextBox { Text = suggestedName, MaxLength = 255 };
        Avalonia.Automation.AutomationProperties.SetName(name, "Filename");
        var error = new TextBlock { Foreground = Brushes.IndianRed, TextWrapping = TextWrapping.Wrap };
        var cancel = new Button { Content = localizer.Get("Common.Cancel"), Classes = { "tool" }, IsCancel = true };
        var next = new Button { Content = "Choose folder", Classes = { "primary" }, IsDefault = true };
        cancel.Click += (_, _) => Close(null);
        next.Click += (_, _) => {
            string value = name.Text?.Trim() ?? string.Empty;
            try {
                StudioStorageAccess.ValidateOutputName(value);
                string? extension = type.Extensions.FirstOrDefault(item => item != "*");
                if (!Path.HasExtension(value) && extension is not null) value += "." + extension;
                if (extension is not null && !type.Extensions.Contains(Path.GetExtension(value).TrimStart('.'), StringComparer.OrdinalIgnoreCase))
                    throw new IOException("Use the selected output format: " + string.Join(", ", type.Extensions.Select(item => "." + item)) + ".");
                StudioStorageAccess.ValidateOutputName(value);
                Close(value);
            } catch (IOException failure) { error.Text = failure.Message; }
        };
        Content = new StackPanel {
            Margin = new Thickness(24), Spacing = 14, Children = {
                new TextBlock { Text = "Filename", FontWeight = FontWeight.SemiBold }, name,
                new TextBlock { Text = "Choose a folder in Files next. Your document is written after you finish the review.", TextWrapping = TextWrapping.Wrap },
                error,
                new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8, Children = { cancel, next } }
            }
        };
        Opened += (_, _) => { name.Focus(); name.SelectAll(); };
    }
}
