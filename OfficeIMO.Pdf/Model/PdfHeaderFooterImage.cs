using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a simple image rendered in a page header or footer.
/// </summary>
public sealed class PdfHeaderFooterImage {
    private readonly byte[] _data;
    private readonly PdfWriter.PdfImageStream? _preparedStream;
    private readonly PdfImageStyle _style;

    /// <summary>Creates a meaningful header/footer image with alternate text.</summary>
    public PdfHeaderFooterImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? alternativeText)
        : this(data, width, height, PdfAlign.Left, OfficeImageFit.Stretch, alternativeText) {
    }

    /// <summary>Creates a header/footer image.</summary>
    public PdfHeaderFooterImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch, string? alternativeText = null)
        : this(data, width, height, align, new PdfImageStyle { Fit = fit, AlternativeText = alternativeText }) {
    }

    /// <summary>Creates a header/footer image with source crop, rotation, mirrors and alternate text from a reusable style.</summary>
    /// <remarks>The style is copied; the explicit alignment controls the image's position in its running-story group.</remarks>
    public PdfHeaderFooterImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, PdfAlign align, PdfImageStyle style) {
        Guard.NotNullOrEmpty(data, nameof(data));
        Guard.NotNull(style, nameof(style));
        Guard.Positive(width.ToPoints(), nameof(width));
        Guard.Positive(height.ToPoints(), nameof(height));
        Guard.LeftCenterRightAlign(align, nameof(align), "PDF header/footer image");
        _style = style.Clone();
        _style.Align = align;
        PdfDocument.ValidateImageStyleForBox(_style, width.ToPoints(), height.ToPoints(), nameof(style));

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
    public global::OfficeIMO.OfficeLength Width { get; }

    /// <summary>Requested image box height in PDF points.</summary>
    public global::OfficeIMO.OfficeLength Height { get; }

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
    }

    internal PdfHeaderFooterImage Clone() => new PdfHeaderFooterImage(this);

    internal ImageBlock ToImageBlock() => new ImageBlock(_data, Width.ToPoints(), Height.ToPoints(), Info, _style,
        useDataSnapshot: true, preparedStream: _preparedStream);
}
