using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a simple image rendered in a page header or footer.
/// </summary>
public sealed class PdfHeaderFooterImage {
    private readonly byte[] _data;
    private readonly PdfWriter.PdfImageStream? _preparedStream;
    private readonly PdfImageStyle _style;

    /// <summary>Creates a header/footer image.</summary>
    public PdfHeaderFooterImage(byte[] data, double width, double height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch)
        : this(data, width, height, align, fit, alternativeText: null) {
    }

    /// <summary>Creates a meaningful header/footer image with alternate text.</summary>
    public PdfHeaderFooterImage(byte[] data, double width, double height, string? alternativeText)
        : this(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText) {
    }

    /// <summary>Creates a header/footer image.</summary>
    public PdfHeaderFooterImage(byte[] data, double width, double height, PdfAlign align, OfficeImageFit fit, string? alternativeText)
        : this(data, width, height, align, new PdfImageStyle { Fit = fit, AlternativeText = alternativeText }) {
    }

    /// <summary>Creates a header/footer image with source crop, rotation, mirrors and alternate text from a reusable style.</summary>
    /// <remarks>The style is copied; the explicit alignment controls the image's position in its running-story group.</remarks>
    public PdfHeaderFooterImage(byte[] data, double width, double height, PdfAlign align, PdfImageStyle style) {
        Guard.NotNullOrEmpty(data, nameof(data));
        Guard.NotNull(style, nameof(style));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.LeftCenterRightAlign(align, nameof(align), "PDF header/footer image");
        _style = style.Clone();
        _style.Align = align;
        PdfDocument.ValidateImageStyleForBox(_style, width, height, nameof(style));

        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(data);
        PdfDocument.ValidateImageFitDimensions(prepared.Info, _style.Fit, nameof(style));

        _data = prepared.Data;
        _preparedStream = prepared.PreparedStream;
        Width = width;
        Height = height;
        Align = align;
        Fit = _style.Fit;
        Info = prepared.Info;
        AlternativeText = _style.AlternativeText;
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

    private PdfHeaderFooterImage(PdfHeaderFooterImage source) {
        _data = (byte[])source._data.Clone();
        _preparedStream = source._preparedStream;
        _style = source._style.Clone();
        Width = source.Width;
        Height = source.Height;
        Align = source.Align;
        Fit = source.Fit;
        Info = source.Info;
        AlternativeText = source.AlternativeText;
        RestsOnHeaderBaseline = source.RestsOnHeaderBaseline;
    }

    /// <summary>
    /// When true, a header image stands on the header text baseline, inline with that text, as a
    /// spreadsheet header picture does. By default a header image hangs below the baseline.
    /// </summary>
    internal bool RestsOnHeaderBaseline { get; set; }

    internal PdfHeaderFooterImage Clone() => new PdfHeaderFooterImage(this);

    internal ImageBlock ToImageBlock() => new ImageBlock(_data, Width, Height, Info, _style,
        useDataSnapshot: true, preparedStream: _preparedStream);
}
