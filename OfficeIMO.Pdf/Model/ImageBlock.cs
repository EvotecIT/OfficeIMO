using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class ImageBlock : IPdfBlock {
    public byte[] Data { get; }
    public double Width { get; }
    public double Height { get; }
    public OfficeImageInfo Info { get; }
    public PdfWriter.PdfImageStream? PreparedStream { get; }
    /// <summary>The parsed drawing of an SVG image drawn as vectors; <c>null</c> for raster images.</summary>
    public OfficeDrawing? SvgDrawing { get; }
    public PdfImageStyle? Style { get; }
    public string? LinkUri { get; }
    public string? LinkContents { get; }
    public OfficeAlignment Align => (Style ?? new PdfImageStyle()).Align;
    public OfficeClipPath? ClipPath => Style?.ClipPath;
    public OfficeImageFit Fit => (Style ?? new PdfImageStyle()).Fit;
    public string? AltText => Style?.AltText;

    public ImageBlock(byte[] data, double width, double height, OfficeImageInfo info, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null, bool useDataSnapshot = false, PdfWriter.PdfImageStream? preparedStream = null, OfficeDrawing? svgDrawing = null) {
        Guard.NotNullOrEmpty(data, nameof(data));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(info, nameof(info));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        if (linkContents != null && linkUri == null) {
            throw new ArgumentException("Image link contents require a link URI.", nameof(linkContents));
        }

        if (linkContents != null) {
            Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
        }

        Data = useDataSnapshot ? data : (byte[])data.Clone();
        Width = width;
        Height = height;
        Info = info;
        PreparedStream = preparedStream;
        SvgDrawing = svgDrawing;
        Style = style?.Clone();
        LinkUri = linkUri;
        LinkContents = linkUri == null ? null : linkContents ?? "Image";
    }

    public ImageBlock(byte[] data, double width, double height, OfficeAlignment align, OfficeImageInfo info, OfficeClipPath? clipPath = null, OfficeImageFit fit = OfficeImageFit.Stretch, string? linkUri = null, string? linkContents = null)
        : this(data, width, height, info, new PdfImageStyle {
            Align = align,
            ClipPath = clipPath,
            Fit = fit
        }, linkUri, linkContents) {
    }
}

