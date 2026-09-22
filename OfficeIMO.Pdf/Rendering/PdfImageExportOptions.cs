using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Format-neutral options for dependency-free PDF page image export.</summary>
public sealed class PdfImageExportOptions : OfficeImageExportOptions {
    /// <summary>Creates PDF image-export options with a conservative 100-page batch limit.</summary>
    public PdfImageExportOptions() {
        MaximumOutputCount = 100;
    }

    /// <inheritdoc />
    public override double LogicalUnitsPerInch => 72D;

    /// <summary>Optional maximum output width or height in pixels.</summary>
    public int? ThumbnailMaxDimension { get; set; }

    /// <summary>Maximum distinct capability diagnostics retained while exporting one page.</summary>
    public int MaxDiagnosticsPerPage { get; set; } = 1_000;

    /// <summary>Maximum aggregate capability diagnostic characters retained while exporting one page.</summary>
    public int MaxDiagnosticCharactersPerPage { get; set; } = 1 * 1024 * 1024;

    internal PdfImageExportOptions Clone() {
        PdfImageExportOptions clone = CopyImageExportOptionsTo(new PdfImageExportOptions());
        clone.ThumbnailMaxDimension = ThumbnailMaxDimension;
        clone.MaxDiagnosticsPerPage = MaxDiagnosticsPerPage;
        clone.MaxDiagnosticCharactersPerPage = MaxDiagnosticCharactersPerPage;
        return clone;
    }

    internal double ResolveScale(OfficeDrawing drawing) {
        Guard.NotNull(drawing, nameof(drawing));
        double scale = Scale;
        if (ThumbnailMaxDimension.HasValue) {
            scale = Math.Min(scale, ThumbnailMaxDimension.Value / Math.Max(drawing.Width, drawing.Height));
        }
        return scale;
    }

    internal void Validate() {
        ValidateImageExportOptions();
        if (ThumbnailMaxDimension.HasValue && ThumbnailMaxDimension.Value < 1) {
            throw new ArgumentOutOfRangeException(nameof(ThumbnailMaxDimension));
        }
        if (MaxDiagnosticsPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDiagnosticsPerPage));
        if (MaxDiagnosticCharactersPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDiagnosticCharactersPerPage));
    }
}
