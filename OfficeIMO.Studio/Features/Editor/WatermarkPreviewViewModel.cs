using Avalonia.Media;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Editor;

public sealed partial class WatermarkPreviewViewModel : ObservableObject, IDisposable {
    private readonly Func<PdfWatermarkOptions, int, CancellationToken, Task<PdfWatermarkPreview>> _prepare;
    private readonly Func<CancellationToken, Task<byte[]?>> _pickImage;
    private readonly IStudioLocalizer _localizer;
    private CancellationTokenSource? _cancellation;
    private byte[]? _image;
    private PdfWatermarkPreview? _prepared;
    private long _settingsVersion;
    private long _preparedVersion = -1;
    private bool _disposed;
    private string _watermarkId = Guid.NewGuid().ToString("N");
    private PdfStandardFont _font = PdfStandardFont.HelveticaBold;
    [ObservableProperty] private WatermarkChoice _selectedWatermark = null!;

    [ObservableProperty] private string _text = "CONFIDENTIAL";
    [ObservableProperty] private bool _useImage;
    [ObservableProperty] private string _color = "#64748B";
    [ObservableProperty] private decimal? _x;
    [ObservableProperty] private decimal? _y;
    [ObservableProperty] private decimal _width = 360;
    [ObservableProperty] private decimal _height = 100;
    [ObservableProperty] private decimal _fontSize = 42;
    [ObservableProperty] private decimal _rotation = -35;
    [ObservableProperty] private decimal _opacity = 35;
    [ObservableProperty] private bool _behindContent;
    [ObservableProperty] private string _pageRange = "";
    [ObservableProperty] private int _previewPage = 1;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string? _errorMessage;
    [ObservableProperty] private Bitmap? _previewImage;

    internal WatermarkPreviewViewModel(int pageCount, int currentPage, IStudioLocalizer localizer,
        Func<PdfWatermarkOptions, int, CancellationToken, Task<PdfWatermarkPreview>> prepare,
        Func<CancellationToken, Task<byte[]?>> pickImage, IReadOnlyList<PdfWatermarkOptions>? existing = null) {
        PageCount = pageCount; _previewPage = Math.Clamp(currentPage, 1, pageCount);
        _localizer = localizer; _prepare = prepare; _pickImage = pickImage;
        Watermarks = new[] { new WatermarkChoice(localizer.Get("Watermark.New"), null) }
            .Concat((existing ?? []).Select((options, index) => new WatermarkChoice(
                (index + 1).ToString(localizer.Culture) + ". " + (options.ImageBytes is null ? options.Text : localizer.Get("Watermark.Image")), options.Clone())))
            .ToArray();
        _selectedWatermark = Watermarks[0];
        if (Watermarks.Count == 2) SelectedWatermark = Watermarks[1];
    }

    public IReadOnlyList<WatermarkChoice> Watermarks { get; }
    partial void OnSelectedWatermarkChanged(WatermarkChoice value) {
        if (value is null) return;
        var options = value.Options?.Clone() ?? new PdfWatermarkOptions();
        _watermarkId = options.Id; _font = options.Font; _image = options.ImageBytes;
        Text = options.Text; UseImage = _image is not null;
        var color = options.Color.ToOfficeColor();
        Color = $"#{color.R:X2}{color.G:X2}{color.B:X2}";
        X = (decimal?)options.X; Y = (decimal?)options.Y;
        Width = (decimal)options.Width; Height = (decimal)options.Height; FontSize = (decimal)options.FontSize;
        Rotation = (decimal)options.RotationDegrees; Opacity = (decimal)(options.Opacity * 100);
        BehindContent = options.BehindContent;
        PageRange = options.TargetPages is null ? "" : string.Join(",", options.TargetPages.Resolve(PageCount));
        if (options.TargetPages is not null) {
            var pages = options.TargetPages.Resolve(PageCount);
            if (!pages.Contains(PreviewPage)) PreviewPage = pages[0];
        }
        _prepared = null; _settingsVersion++; ClearPreviewImage();
        OnPropertyChanged(nameof(CanApply)); OnPropertyChanged(nameof(HasImage));
    }

    public int PageCount { get; }
    public bool CanEdit => !IsBusy && !_disposed;
    public bool CanApply => CanEdit && _prepared is not null && _preparedVersion == _settingsVersion;
    public bool HasImage => _image is not null;
    internal PdfWatermarkPreview? Prepared => CanApply ? _prepared : null;

    protected override void OnPropertyChanged(System.ComponentModel.PropertyChangedEventArgs e) {
        base.OnPropertyChanged(e);
        if (e.PropertyName is nameof(Text) or nameof(UseImage) or nameof(Color) or nameof(X) or nameof(Y)
            or nameof(Width) or nameof(Height) or nameof(FontSize) or nameof(Rotation) or nameof(Opacity)
            or nameof(BehindContent) or nameof(PageRange) or nameof(PreviewPage)) {
            _settingsVersion++;
            _prepared = null;
            ClearPreviewImage();
            OnPropertyChanged(nameof(CanApply));
        }
    }

    partial void OnIsBusyChanged(bool value) {
        OnPropertyChanged(nameof(CanEdit)); OnPropertyChanged(nameof(CanApply));
        PreviewCommand.NotifyCanExecuteChanged(); ChooseImageCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task ChooseImageAsync(CancellationToken token) {
        try {
            byte[]? image = await _pickImage(token).ConfigureAwait(true);
            if (image is null || _disposed) return;
            _image = image; _prepared = null; _settingsVersion++;
            UseImage = true;
            OnPropertyChanged(nameof(HasImage)); OnPropertyChanged(nameof(CanApply));
            await PreviewAsync().ConfigureAwait(true);
        } catch (OperationCanceledException) { }
        catch (Exception error) { if (!_disposed) ErrorMessage = error.Message; }
    }

    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task PreviewAsync() {
        if (_disposed || IsBusy) return;
        IsBusy = true; ErrorMessage = null; _prepared = null;
        long version = _settingsVersion;
        using var cancellation = new CancellationTokenSource();
        _cancellation = cancellation;
        try {
            if (UseImage && _image is null) throw new InvalidOperationException(_localizer.Get("Watermark.ChooseImageFirst"));
            var color = Avalonia.Media.Colors.Gray;
            if (!UseImage && !Avalonia.Media.Color.TryParse(Color, out color))
                throw new InvalidOperationException(_localizer.Get("Watermark.InvalidColor"));
            var options = new PdfWatermarkOptions {
                Id = _watermarkId, Font = _font,
                Text = Text, ImageBytes = UseImage ? _image : null,
                X = X.HasValue ? (double)X.Value : null, Y = Y.HasValue ? (double)Y.Value : null,
                Width = (double)Width, Height = (double)Height, FontSize = (double)FontSize,
                RotationDegrees = (double)Rotation, Opacity = (double)Opacity / 100D,
                Color = PdfColor.FromRgb(color.R, color.G, color.B), BehindContent = BehindContent,
                TargetPages = string.IsNullOrWhiteSpace(PageRange) ? null : PdfPageSelector.Parse(PageRange)
            };
            PdfWatermarkPreview result = await _prepare(options, PreviewPage, cancellation.Token).ConfigureAwait(true);
            if (_disposed || version != _settingsVersion) return;
            using var stream = new MemoryStream(result.PageImage, writable: false);
            var image = new Bitmap(stream);
            var previous = PreviewImage; PreviewImage = image; previous?.Dispose();
            _prepared = result; _preparedVersion = version;
        } catch (OperationCanceledException) { }
        catch (Exception error) { if (!_disposed) ErrorMessage = error.Message; }
        finally {
            if (ReferenceEquals(_cancellation, cancellation)) _cancellation = null;
            if (!_disposed) IsBusy = false;
        }
    }

    public void Dispose() {
        _disposed = true; _cancellation?.Cancel(); _prepared = null;
        ClearPreviewImage(); _image = null;
    }

    private void ClearPreviewImage() {
        var previous = PreviewImage;
        PreviewImage = null;
        previous?.Dispose();
    }
}

public sealed record WatermarkChoice(string Label, PdfWatermarkOptions? Options) {
    public override string ToString() => Label;
}
