using System.Globalization;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Sign;

public sealed record SignatureStyleChoice(string Preview, FontFamily Family, FontStyle Style);

/// <summary>Creates a signature or initials image by typing, drawing, or choosing a picture.</summary>
public sealed partial class SignatureDialog : Window {
    // Handwriting-style faces that ship with Windows, macOS and common Linux desktops, with safe fallbacks.
    private static readonly (string Family, FontStyle Style)[] Faces = [
        ("Segoe Script, Snell Roundhand, URW Chancery L, Z003, serif", FontStyle.Normal),
        ("Ink Free, Bradley Hand, Comic Neue, Comic Sans MS, sans-serif", FontStyle.Normal),
        ("Lucida Handwriting, Apple Chancery, Georgia, serif", FontStyle.Italic)
    ];
    private byte[]? _image;

    public SignatureDialog() {
        InitializeComponent();
        Pad.InkChanged += (_, _) => UpdateState();
    }

    internal SignatureDialog(StudioSignatureKind kind, IStudioLocalizer localizer) : this() {
        string heading = localizer.Get(kind == StudioSignatureKind.Initials ? "FillSign.CreateInitials" : "FillSign.CreateSignature");
        Title = heading;
        HeadingText.Text = heading;
        NameInput.PlaceholderText = localizer.Get(kind == StudioSignatureKind.Initials ? "FillSign.InitialsPlaceholder" : "FillSign.NamePlaceholder");
        RefreshStyles();
        Opened += (_, _) => NameInput.Focus();
    }

    internal TextBox NameBox => NameInput;

    internal SignaturePad DrawingPad => Pad;

    internal void ShowMethod(int index) => MethodTabs.SelectedIndex = index;

    private void OnMethodChanged(object? sender, SelectionChangedEventArgs e) {
        if (TypePanel is null) return;
        TypePanel.IsVisible = MethodTabs.SelectedItem == TypeTab;
        DrawPanel.IsVisible = MethodTabs.SelectedItem == DrawTab;
        ImagePanel.IsVisible = MethodTabs.SelectedItem == ImageTab;
        UpdateState();
    }

    private void OnNameChanged(object? sender, TextChangedEventArgs e) => RefreshStyles();

    private void OnStyleChanged(object? sender, SelectionChangedEventArgs e) => UpdateState();

    private void RefreshStyles() {
        if (StyleList is null) return;
        int selected = Math.Max(0, StyleList.SelectedIndex);
        string preview = string.IsNullOrWhiteSpace(NameInput.Text) ? NameInput.PlaceholderText ?? string.Empty : NameInput.Text.Trim();
        StyleList.ItemsSource = Faces.Select(face => new SignatureStyleChoice(preview, new FontFamily(face.Family), face.Style)).ToArray();
        StyleList.SelectedIndex = selected;
        UpdateState();
    }

    private void OnClearClick(object? sender, RoutedEventArgs e) => Pad.Clear();

    private async void OnChooseImageClick(object? sender, RoutedEventArgs e) {
        if (!StorageProvider.CanOpen) return;
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            AllowMultiple = false,
            FileTypeFilter = [new FilePickerFileType("PNG, JPEG") { Patterns = ["*.png", "*.jpg", "*.jpeg"] }]
        });
        if (files.Count == 0) return;
        try {
            byte[]? image = await StudioStorageInput.ReadImageAsync(files, CancellationToken.None);
            if (image is not null) SetImage(image);
        } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or InvalidOperationException) {
            ImageEmptyText.Text = ex.Message;
            ImageEmptyText.IsVisible = true;
        }
    }

    internal void SetImage(byte[] image) {
        using var stream = new MemoryStream(image);
        ImagePreview.Source = new Bitmap(stream);
        _image = image;
        ImageEmptyText.IsVisible = false;
        UpdateState();
    }

    private void UpdateState() {
        if (UseButton is null || MethodTabs is null) return;
        object? tab = MethodTabs.SelectedItem;
        UseButton.IsEnabled = tab == DrawTab ? Pad.HasInk
            : tab == ImageTab ? _image is not null
            : !string.IsNullOrWhiteSpace(NameInput.Text);
    }

    private void OnCancelClick(object? sender, RoutedEventArgs e) => Close(null);

    private void OnUseClick(object? sender, RoutedEventArgs e) {
        byte[]? png = CreateImage();
        if (png is null) return;
        bool typed = MethodTabs.SelectedItem == TypeTab;
        Close(new StudioSignatureDraft(png, RememberBox.IsChecked == true,
            typed ? NameInput.Text?.Trim() : null,
            MethodTabs.SelectedItem == DrawTab ? Pad.NormalizedStrokes() : null));
    }

    internal byte[]? CreateImage() {
        if (MethodTabs.SelectedItem == DrawTab) return Pad.ToPng();
        if (MethodTabs.SelectedItem == ImageTab) return _image;
        if (string.IsNullOrWhiteSpace(NameInput.Text) || StyleList.SelectedItem is not SignatureStyleChoice style) return null;
        return RenderTyped(NameInput.Text.Trim(), style);
    }

    // Typed signatures are rendered at high resolution on a transparent background so they sit on any page.
    private static byte[] RenderTyped(string text, SignatureStyleChoice style) {
        const double fontSize = 96D;
        var formatted = new FormattedText(text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight,
            new Typeface(style.Family, style.Style), fontSize, new SolidColorBrush(SignaturePad.InkColor));
        double padding = fontSize * 0.2D;
        var size = new PixelSize(
            Math.Max(1, (int)Math.Ceiling(formatted.WidthIncludingTrailingWhitespace + padding * 2D)),
            Math.Max(1, (int)Math.Ceiling(formatted.Height + padding)));
        using var bitmap = new RenderTargetBitmap(size, new Vector(96D, 96D));
        using (DrawingContext context = bitmap.CreateDrawingContext()) context.DrawText(formatted, new Point(padding, padding / 2D));
        using var stream = new MemoryStream();
        bitmap.Save(stream, PngBitmapEncoderOptions.Default);
        return stream.ToArray();
    }
}
