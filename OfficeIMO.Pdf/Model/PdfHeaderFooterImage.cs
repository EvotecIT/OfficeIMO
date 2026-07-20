using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a simple image rendered in a page header or footer.
/// </summary>
public sealed class PdfHeaderFooterImage {
    private readonly byte[] _data;

    /// <summary>Creates a header/footer image.</summary>
    public PdfHeaderFooterImage(byte[] data, double width, double height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch)
        : this(data, width, height, align, fit, alternativeText: null) {
    }

    /// <summary>Creates a meaningful header/footer image with alternate text.</summary>
    public PdfHeaderFooterImage(byte[] data, double width, double height, string? alternativeText)
        : this(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText) {
    }

    /// <summary>Creates a header/footer image.</summary>
    public PdfHeaderFooterImage(byte[] data, double width, double height, PdfAlign align, OfficeImageFit fit, string? alternativeText) {
        Guard.NotNullOrEmpty(data, nameof(data));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.LeftCenterRightAlign(align, nameof(align), "PDF header/footer image");
        PdfDocument.ValidateImageFit(fit, nameof(fit));
        if (alternativeText != null) {
            Guard.NotNullOrWhiteSpace(alternativeText, nameof(alternativeText));
        }

        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(data);
        PdfDocument.ValidateImageFitDimensions(prepared.Info, fit, nameof(fit));

        _data = prepared.Data;
        Width = width;
        Height = height;
        Align = align;
        Fit = fit;
        Info = prepared.Info;
        AlternativeText = alternativeText;
    }

    /// <summary>Image payload.</summary>
    public byte[] Data => (byte[])_data.Clone();

    /// <summary>Requested image box width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Requested image box height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Horizontal placement inside the page content width.</summary>
    public PdfAlign Align { get; }

    /// <summary>Image fit behavior inside the requested image box.</summary>
    public OfficeImageFit Fit { get; }

    /// <summary>Validated image metadata.</summary>
    public OfficeImageInfo Info { get; }

    /// <summary>Optional alternate text for meaningful header/footer images.</summary>
    public string? AlternativeText { get; }

    internal PdfHeaderFooterImage Clone() => new PdfHeaderFooterImage(_data, Width, Height, Align, Fit, AlternativeText);

    internal ImageBlock ToImageBlock() => new ImageBlock(_data, Width, Height, Info, new PdfImageStyle {
        Align = Align,
        Fit = Fit,
        AlternativeText = AlternativeText
    }, useDataSnapshot: true);
}
