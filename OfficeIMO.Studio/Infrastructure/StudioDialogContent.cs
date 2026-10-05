using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Platform.Storage;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Reusable dialog content presented in a desktop window or an in-app mobile sheet.</summary>
public class StudioDialogContent : UserControl {
    public static readonly StyledProperty<string?> TitleProperty = AvaloniaProperty.Register<StudioDialogContent, string?>(nameof(Title));
    private bool _closed;

    public string? Title { get => GetValue(TitleProperty); set => SetValue(TitleProperty, value); }
    /// <summary>Whether the content already presents its title and dismissal actions inside the sheet.</summary>
    public bool HasContentHeading { get; set; }
    public bool CanResize { get; set; } = true;
    public SizeToContent SizeToContent { get; set; }
    public WindowStartupLocation WindowStartupLocation { get; set; } = WindowStartupLocation.CenterOwner;
    protected Control? Owner { get; private set; }
    protected IStorageProvider StorageProvider => TopLevel.GetTopLevel(this)!.StorageProvider;
    public event EventHandler? Opened;
    public event EventHandler? Closed;
    internal event Action<object?>? CloseRequested;

    public StudioDialogContent() {
        KeyDown += (_, e) => {
            if (e.Key != Key.Escape) return;
            Close(null);
            e.Handled = true;
        };
    }

    protected void Close(object? result = null) => CloseRequested?.Invoke(result);

    internal void Presented(Control owner) {
        Owner = owner;
        Opened?.Invoke(this, EventArgs.Empty);
    }

    internal void Dismissed() {
        if (_closed) return;
        _closed = true;
        Closed?.Invoke(this, EventArgs.Empty);
        Owner = null;
    }
}
