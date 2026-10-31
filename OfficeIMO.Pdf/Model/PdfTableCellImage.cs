using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents an image rendered inside a table cell.
/// </summary>
public sealed class PdfTableCellImage {
    private readonly byte[] _data;
    private readonly PdfWriter.PdfImageStream? _preparedStream;

    /// <summary>Creates a table-cell image from raster bytes supported by OfficeIMO.Drawing.</summary>
    public PdfTableCellImage(byte[] data, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNullOrEmpty(data, nameof(data));
        Guard.Positive(width.ToPoints(), nameof(width));
        Guard.Positive(height.ToPoints(), nameof(height));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        if (linkContents != null && linkUri == null) {
            throw new ArgumentException("Table cell image link contents require a link URI.", nameof(linkContents));
        }

        if (linkContents != null) {
            Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
        }

        PdfImageStyle? imageStyle = style?.Clone();
        if (imageStyle != null) {
            PdfDocument.ValidateImageStyleForBox(imageStyle, width.ToPoints(), height.ToPoints(), nameof(style));
        }

        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(data);
        if (imageStyle != null) {
            PdfDocument.ValidateImageFitDimensions(prepared.Info, imageStyle.Fit, nameof(style));
        }

        _data = prepared.Data;
        _preparedStream = prepared.PreparedStream;
        Width = width;
        Height = height;
        Info = prepared.Info;
        Style = imageStyle;
        LinkUri = linkUri;
        LinkContents = linkUri == null ? null : linkContents ?? "Table cell image";
    }

    /// <summary>Image bytes.</summary>
    public byte[] Data => (byte[])_data.Clone();

    /// <summary>Target image width in PDF points.</summary>
    public global::OfficeIMO.OfficeLength Width { get; }

    /// <summary>Target image height in PDF points.</summary>
    public global::OfficeIMO.OfficeLength Height { get; }

    /// <summary>Detected source image metadata.</summary>
    public OfficeImageInfo Info { get; }

    /// <summary>Optional image style. When omitted, the table cell alignment is used.</summary>
    public PdfImageStyle? Style { get; }

    /// <summary>Optional absolute URI or catalog-base-relative URI linked from the image rectangle.</summary>
    public string? LinkUri { get; }

    /// <summary>Optional PDF annotation contents metadata for the image link.</summary>
    public string? LinkContents { get; }

    private PdfTableCellImage(PdfTableCellImage source) {
        // Preparation owns this immutable snapshot; Data always returns a caller-owned copy.
        // Cell/style snapshots can share its payload, just as they share the prepared stream.
        _data = source._data;
        _preparedStream = source._preparedStream;
        Width = source.Width;
        Height = source.Height;
        Info = source.Info;
        Style = source.Style?.Clone();
        LinkUri = source.LinkUri;
        LinkContents = source.LinkContents;
    }

    internal PdfTableCellImage Clone() => new PdfTableCellImage(this);

    internal ImageBlock ToImageBlock(PdfAlign fallbackAlign) {
        PdfImageStyle style = Style?.Clone() ?? new PdfImageStyle {
            Align = fallbackAlign
        };
        return new ImageBlock(
            _data,
            Width.ToPoints(),
            Height.ToPoints(),
            Info,
            style,
            LinkUri,
            LinkContents,
            useDataSnapshot: true,
            preparedStream: _preparedStream);
    }
}
