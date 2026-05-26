using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class ImageBlock : IPdfBlock {
    public byte[] Data { get; }
    public double Width { get; }
    public double Height { get; }
    public OfficeImageInfo Info { get; }
    public PdfImageStyle? Style { get; }
    public string? LinkUri { get; }
    public string? LinkContents { get; }
    public PdfAlign Align => (Style ?? new PdfImageStyle()).Align;
    public OfficeClipPath? ClipPath => Style?.ClipPath;
    public OfficeImageFit Fit => (Style ?? new PdfImageStyle()).Fit;

    public ImageBlock(byte[] data, double width, double height, OfficeImageInfo info, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNullOrEmpty(data, nameof(data));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.NotNull(info, nameof(info));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        if (linkContents != null && linkUri == null) {
            throw new ArgumentException("Image link contents require a link URI.", nameof(linkContents));
        }

        if (linkContents != null) {
            Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
        }

        Data = (byte[])data.Clone();
        Width = width;
        Height = height;
        Info = info;
        Style = style?.Clone();
        LinkUri = linkUri;
        LinkContents = linkUri == null ? null : linkContents ?? "Image";
    }

    public ImageBlock(byte[] data, double width, double height, PdfAlign align, OfficeImageInfo info, OfficeClipPath? clipPath = null, OfficeImageFit fit = OfficeImageFit.Stretch, string? linkUri = null, string? linkContents = null)
        : this(data, width, height, info, new PdfImageStyle {
            Align = align,
            ClipPath = clipPath,
            Fit = fit
        }, linkUri, linkContents) {
    }
}

