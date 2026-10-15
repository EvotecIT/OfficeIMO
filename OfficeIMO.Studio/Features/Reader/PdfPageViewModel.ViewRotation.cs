using Avalonia.Media;

namespace OfficeIMO.Studio.Features.Reader;

public sealed partial class PdfPageViewModel {
    private int _viewRotationDegrees;

    /// <summary>The reader's clockwise turn, independent of the PDF page's saved rotation.</summary>
    public int ViewRotationDegrees => _viewRotationDegrees;
    public ITransform ViewRotationTransform => new RotateTransform(_viewRotationDegrees);
    public double ContentDisplayWidth => Math.Max(1, Math.Round(_pageWidth * _zoom, 2) - 2);
    public double ContentDisplayHeight => Math.Max(1, Math.Round(_pageHeight * _zoom, 2) - 2);

    internal void SetViewRotation(int degrees) {
        int normalized = ((degrees % 360) + 360) % 360;
        if (normalized % 90 != 0) throw new ArgumentOutOfRangeException(nameof(degrees));
        if (!SetProperty(ref _viewRotationDegrees, normalized, nameof(ViewRotationDegrees))) return;
        OnPropertyChanged(nameof(ViewRotationTransform));
        UpdateDisplaySize();
    }
}
