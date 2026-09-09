using System.ComponentModel;
using System.Security.Cryptography;
using Avalonia;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>A localized scan color choice.</summary>
public sealed record ScanColorChoice(OfficeScanColorMode Value, string Label) {
    /// <inheritdoc />
    public override string ToString() => Label;
}

/// <summary>Reviewed scan settings and source-coordinate selection over the shared PDF preparation owner.</summary>
public sealed partial class ScanPreparationViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<byte[]>> _readSource;
    private readonly Func<PdfOcrMergeOptions, string, CancellationToken, Task> _save;
    private readonly IStudioLocalizer _localizer;
    private CancellationTokenSource? _previewCancellation;
    private PdfOcrMergeOptions? _reviewedOptions;
    private string? _reviewedHash;
    private bool _disposed;
    internal ScanPreparationViewModel(Func<CancellationToken, Task<byte[]>> readSource,
        Func<PdfOcrMergeOptions, string, CancellationToken, Task> save, IStudioLocalizer localizer) {
        _readSource = readSource; _save = save; _localizer = localizer;
        ColorModes = [
            new(OfficeScanColorMode.PreserveColor, T("Color.Keep", "Keep color")),
            new(OfficeScanColorMode.Grayscale, T("Color.Gray", "Grayscale")),
            new(OfficeScanColorMode.Bilevel, T("Color.BlackWhite", "Black and white"))
        ];
    }
    [ObservableProperty] private int _pageNumber = 1;
    [ObservableProperty] private double _dpi = 150;
    [ObservableProperty] private bool _enableCleanup;
    [ObservableProperty] private bool _deskew = true;
    [ObservableProperty] private bool _normalizeBackground = true;
    [ObservableProperty] private int _quarterTurns;
    [ObservableProperty] private double _straightenDegrees;
    [ObservableProperty] private int _blackPoint;
    [ObservableProperty] private int _whitePoint = 255;
    [ObservableProperty] private double _gamma = 1;
    [ObservableProperty] private OfficeScanColorMode _colorMode = OfficeScanColorMode.Grayscale;
    [ObservableProperty] private bool _useRegion;
    [ObservableProperty] private Rect _region = new(0, 0, 1, 1);
    [ObservableProperty] private bool _usePerspective;
    [ObservableProperty] private Point _topLeft = new(0, 0);
    [ObservableProperty] private Point _topRight = new(1, 0);
    [ObservableProperty] private Point _bottomRight = new(1, 1);
    [ObservableProperty] private Point _bottomLeft = new(0, 1);
    [ObservableProperty] private bool _editCorners;
    [ObservableProperty] private Bitmap? _sourcePreview;
    [ObservableProperty] private Bitmap? _preparedPreview;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private bool _hostBusy;
    [ObservableProperty] private bool _isCurrent;
    [ObservableProperty] private string _status = string.Empty;
    public IReadOnlyList<ScanColorChoice> ColorModes { get; }
    public ScanColorChoice SelectedColorMode {
        get => ColorModes.First(choice => choice.Value == ColorMode);
        set { if (value != null) ColorMode = value.Value; }
    }
    partial void OnColorModeChanged(OfficeScanColorMode value) => OnPropertyChanged(nameof(SelectedColorMode));
    public bool CanPreview => !_disposed && !IsBusy && !HostBusy;
    public bool CanSave => CanPreview && IsCurrent;
    public bool CanEdit => CanPreview;
    public bool HasPreview => SourcePreview != null;
    protected override void OnPropertyChanged(PropertyChangedEventArgs e) {
        base.OnPropertyChanged(e);
        if (e.PropertyName is nameof(PageNumber) or nameof(Dpi) or nameof(EnableCleanup) or nameof(Deskew) or nameof(NormalizeBackground)
            or nameof(QuarterTurns) or nameof(StraightenDegrees) or nameof(BlackPoint) or nameof(WhitePoint) or nameof(Gamma)
            or nameof(ColorMode) or nameof(UseRegion) or nameof(Region) or nameof(UsePerspective) or nameof(TopLeft)
            or nameof(TopRight) or nameof(BottomRight) or nameof(BottomLeft)) Invalidate();
        if (e.PropertyName is nameof(IsBusy) or nameof(HostBusy) or nameof(IsCurrent)) {
            PreviewCommand.NotifyCanExecuteChanged(); SaveCommand.NotifyCanExecuteChanged();
            OnPropertyChanged(nameof(CanEdit)); OnPropertyChanged(nameof(CanSave));
        }
        if (e.PropertyName == nameof(SourcePreview)) OnPropertyChanged(nameof(HasPreview));
    }
    internal void Invalidate(bool clearSource = false) {
        _previewCancellation?.Cancel(); _reviewedOptions = null; _reviewedHash = null; IsCurrent = false;
        if (clearSource) {
            SourcePreview?.Dispose(); SourcePreview = null; PreparedPreview?.Dispose(); PreparedPreview = null;
            Region = new(0, 0, 1, 1); UseRegion = false;
            PageNumber = 1; UsePerspective = false;
        }
        Status = T("Refresh", "Preview the current settings before saving a prepared copy.");
    }
    internal PdfOcrMergeOptions ApplyTo(PdfOcrMergeOptions options) {
        options.Dpi = Dpi;
        options.ScanProcessing = EnableCleanup ? new OfficeScanProcessingOptions {
            ClockwiseQuarterTurns = QuarterTurns,
            StraightenDegrees = StraightenDegrees,
            Deskew = Deskew,
            NormalizeBackground = NormalizeBackground,
            BlackPoint = BlackPoint,
            WhitePoint = WhitePoint,
            Gamma = Gamma,
            ColorMode = ColorMode
        } : null;
        options.Perspective = UsePerspective ? new OfficeScanPerspectiveOptions {
            TopLeft = new(TopLeft.X, TopLeft.Y),
            TopRight = new(TopRight.X, TopRight.Y),
            BottomRight = new(BottomRight.X, BottomRight.Y),
            BottomLeft = new(BottomLeft.X, BottomLeft.Y)
        } : null;
        options.Regions = UseRegion ? new[] { new PdfOcrPageRegion(PageNumber, Region.X, Region.Y, Region.Width, Region.Height) } : [];
        return options;
    }
    [RelayCommand(CanExecute = nameof(CanPreview))]
    private async Task PreviewAsync(CancellationToken token) {
        Invalidate();
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _previewCancellation = operation; IsBusy = true;
        try {
            int page = PageNumber;
            var options = ApplyTo(new PdfOcrMergeOptions { ReadOptions = new() { PageSelection = PdfPageSelection.From(page) } });
            byte[] source = await _readSource(operation.Token).ConfigureAwait(true);
            var preview = await Task.Run(() => PdfDocument.Load(source).PreviewScanAsync(page, options, operation.Token), operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            if (_disposed || !ReferenceEquals(_previewCancellation, operation)) return;
            using var original = new MemoryStream(preview.GetSourcePng());
            using var processed = new MemoryStream(preview.GetPreparedPng());
            Bitmap sourceBitmap = new(original);
            Bitmap preparedBitmap;
            try { preparedBitmap = new(processed); } catch { sourceBitmap.Dispose(); throw; }
            SourcePreview?.Dispose(); PreparedPreview?.Dispose(); SourcePreview = sourceBitmap; PreparedPreview = preparedBitmap;
            _reviewedOptions = options.Clone(); _reviewedHash = Convert.ToHexString(SHA256.HashData(source)); IsCurrent = true;
            Status = T("Ready", "Prepared preview ready. Region OCR keeps the visible source page; saving a scan copy keeps only the prepared pixels.") +
                " " + string.Join(" ", preview.Diagnostics.Concat(preview.ScanProcessing?.Steps.Where(step => step.Applied).Select(step => step.Message) ?? []));
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) { } catch (Exception error) { if (ReferenceEquals(_previewCancellation, operation)) Status = error.Message; } finally { if (ReferenceEquals(_previewCancellation, operation)) _previewCancellation = null; IsBusy = false; }
    }
    [RelayCommand(CanExecute = nameof(CanSave))]
    private async Task SaveAsync(CancellationToken token) {
        if (_reviewedOptions == null || _reviewedHash == null) return;
        var options = _reviewedOptions.Clone(); string hash = _reviewedHash;
        IsBusy = true;
        try { await _save(options, hash, token).ConfigureAwait(true); } catch (Exception error) { Status = error.Message; } finally { IsBusy = false; }
    }
    [RelayCommand]
    private void ResetSelection() {
        if (!CanEdit) return;
        Region = new(0, 0, 1, 1); TopLeft = new(0, 0); TopRight = new(1, 0); BottomRight = new(1, 1); BottomLeft = new(0, 1);
    }
    public void Dispose() { _disposed = true; _previewCancellation?.Cancel(); SourcePreview?.Dispose(); PreparedPreview?.Dispose(); }
    private string T(string key, string fallback) => _localizer.GetOrDefault("Scan." + key, fallback);
}