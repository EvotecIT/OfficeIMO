using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Represents an image rendered inside a table cell.
/// </summary>
public sealed class PdfTableCellImage {
    /// <summary>Creates a supported table-cell image. JPEG and simple PNG images, including Adam7 interlace, indexed-color palettes, and alpha soft masks, are supported.</summary>
    public PdfTableCellImage(byte[] data, double width, double height, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNullOrEmpty(data, nameof(data));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
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

        OfficeImageInfo info = PdfDocument.ValidateImageBytes(data);
        if (imageStyle != null) {
            PdfDocument.ValidateImageFitDimensions(info, imageStyle.Fit, nameof(style));
        }

        Data = (byte[])data.Clone();
        Width = width;
        Height = height;
        Info = info;
        Style = imageStyle;
        LinkUri = linkUri;
        LinkContents = linkUri == null ? null : linkContents ?? "Table cell image";
    }

    /// <summary>Image bytes.</summary>
    public byte[] Data { get; }

    /// <summary>Target image width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Target image height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Detected source image metadata.</summary>
    public OfficeImageInfo Info { get; }

    /// <summary>Optional image style. When omitted, the table cell alignment is used.</summary>
    public PdfImageStyle? Style { get; }

    /// <summary>Optional absolute URI linked from the image rectangle.</summary>
    public string? LinkUri { get; }

    /// <summary>Optional PDF annotation contents metadata for the image link.</summary>
    public string? LinkContents { get; }

    internal PdfTableCellImage Clone() => new PdfTableCellImage(Data, Width, Height, Style, LinkUri, LinkContents);

    internal ImageBlock ToImageBlock(PdfAlign fallbackAlign) {
        PdfImageStyle style = Style?.Clone() ?? new PdfImageStyle {
            Align = fallbackAlign
        };
        return new ImageBlock(Data, Width, Height, Info, style, LinkUri, LinkContents);
    }
}
