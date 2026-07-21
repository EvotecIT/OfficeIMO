using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents an image rendered inside a table cell.
/// </summary>
public sealed class PdfTableCellImage {
    private readonly byte[] _data;

    /// <summary>Creates a table-cell image from raster bytes supported by OfficeIMO.Drawing.</summary>
    public PdfTableCellImage(byte[] data, double width, double height, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNullOrEmpty(data, nameof(data));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        if (linkContents != null && linkUri == null) {
            throw new ArgumentException("Table cell image link contents require a link URI.", nameof(linkContents));
        }

        if (linkContents != null) {
            Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
        }

        PdfImageStyle? imageStyle = style?.Clone();
        if (imageStyle != null) {
            PdfDocument.ValidateImageStyleForBox(imageStyle, width, height, nameof(style));
        }

        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(data);
        if (imageStyle != null) {
            PdfDocument.ValidateImageFitDimensions(prepared.Info, imageStyle.Fit, nameof(style));
        }

        _data = prepared.Data;
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
    public double Width { get; }

    /// <summary>Target image height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Detected source image metadata.</summary>
    public OfficeImageInfo Info { get; }

    /// <summary>Optional image style. When omitted, the table cell alignment is used.</summary>
    public PdfImageStyle? Style { get; }

    /// <summary>Optional absolute URI or catalog-base-relative URI linked from the image rectangle.</summary>
    public string? LinkUri { get; }

    /// <summary>Optional PDF annotation contents metadata for the image link.</summary>
    public string? LinkContents { get; }

    internal PdfTableCellImage Clone() => new PdfTableCellImage(_data, Width, Height, Style, LinkUri, LinkContents);

    internal ImageBlock ToImageBlock(PdfAlign fallbackAlign) {
        PdfImageStyle style = Style?.Clone() ?? new PdfImageStyle {
            Align = fallbackAlign
        };
        return new ImageBlock(_data, Width, Height, Info, style, LinkUri, LinkContents, useDataSnapshot: true);
    }
}
