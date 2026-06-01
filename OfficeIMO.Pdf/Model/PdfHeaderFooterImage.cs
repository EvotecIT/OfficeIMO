using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Describes a simple image rendered in a page header or footer.
/// </summary>
public sealed class PdfHeaderFooterImage {
    /// <summary>Creates a header/footer image.</summary>
    public PdfHeaderFooterImage(byte[] data, double width, double height, PdfAlign align = PdfAlign.Left, OfficeImageFit fit = OfficeImageFit.Stretch) {
        Guard.NotNullOrEmpty(data, nameof(data));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.LeftCenterRightAlign(align, nameof(align), "PDF header/footer image");
        PdfDoc.ValidateImageFit(fit, nameof(fit));

        OfficeImageInfo imageInfo = PdfDoc.ValidateImageBytes(data);
        PdfDoc.ValidateImageFitDimensions(imageInfo, fit, nameof(fit));

        Data = (byte[])data.Clone();
        Width = width;
        Height = height;
        Align = align;
        Fit = fit;
        Info = imageInfo;
    }

    /// <summary>Image payload.</summary>
    public byte[] Data { get; }

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

    internal PdfHeaderFooterImage Clone() => new PdfHeaderFooterImage(Data, Width, Height, Align, Fit);

    internal ImageBlock ToImageBlock() => new ImageBlock(Data, Width, Height, Info, new PdfImageStyle {
        Align = Align,
        Fit = Fit
    });
}
