using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Format-neutral options for dependency-free PDF page image export.</summary>
public sealed class PdfImageExportOptions : OfficeImageExportOptions {
    /// <summary>Optional target DPI. PDF user units use 72 points per inch.</summary>
    public double? Dpi { get; set; }

    /// <summary>Optional maximum output width or height in pixels.</summary>
    public int? ThumbnailMaxDimension { get; set; }

    /// <summary>Maximum number of pages returned by one batch export.</summary>
    public int MaxPages { get; set; } = 100;

    internal PdfImageExportOptions Clone() {
        PdfImageExportOptions clone = CopyImageExportOptionsTo(new PdfImageExportOptions());
        clone.Dpi = Dpi;
        clone.ThumbnailMaxDimension = ThumbnailMaxDimension;
        clone.MaxPages = MaxPages;
        return clone;
    }

    internal double ResolveScale(OfficeDrawing drawing) {
        Guard.NotNull(drawing, nameof(drawing));
        double scale = Dpi.HasValue ? Dpi.Value / 72D : Scale;
        if (ThumbnailMaxDimension.HasValue) {
            scale = Math.Min(scale, ThumbnailMaxDimension.Value / Math.Max(drawing.Width, drawing.Height));
        }
        return scale;
    }

    internal void Validate() {
        ValidateImageExportOptions();
        if (Dpi.HasValue && (Dpi.Value <= 0D || double.IsNaN(Dpi.Value) || double.IsInfinity(Dpi.Value))) {
            throw new ArgumentOutOfRangeException(nameof(Dpi), "DPI must be finite and positive.");
        }
        if (ThumbnailMaxDimension.HasValue && ThumbnailMaxDimension.Value < 1) {
            throw new ArgumentOutOfRangeException(nameof(ThumbnailMaxDimension));
        }
        if (MaxPages < 1) throw new ArgumentOutOfRangeException(nameof(MaxPages));
    }
}
