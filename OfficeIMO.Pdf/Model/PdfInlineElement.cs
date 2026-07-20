using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Base class for a fixed-size visual that participates in rich-paragraph wrapping.
/// </summary>
public abstract class PdfInlineElement {
    /// <summary>Width of the inline element in points.</summary>
    public double Width { get; }

    /// <summary>Height of the inline element in points.</summary>
    public double Height { get; }

    /// <summary>Vertical offset of the element bottom from the text baseline, in points.</summary>
    public double BaselineOffset { get; }

    /// <summary>Alternative text used when the element is emitted into a tagged document.</summary>
    public string? AlternativeText { get; }

    private protected PdfInlineElement(double width, double height, double baselineOffset, string? alternativeText) {
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        if (double.IsNaN(baselineOffset) || double.IsInfinity(baselineOffset)) {
            throw new ArgumentException("Inline element baseline offset must be finite.", nameof(baselineOffset));
        }

        if (alternativeText != null) {
            Guard.NotNullOrWhiteSpace(alternativeText, nameof(alternativeText));
        }

        Width = width;
        Height = height;
        BaselineOffset = baselineOffset;
        AlternativeText = alternativeText;
    }
}

/// <summary>
/// Raster image that flows between text runs as one indivisible inline element.
/// </summary>
public sealed class PdfInlineImage : PdfInlineElement {
    internal ImageBlock Block { get; }

    /// <summary>Image fitting mode inside the requested inline frame.</summary>
    public OfficeImageFit Fit { get; }

    /// <summary>Creates an inline image from first-party supported image bytes.</summary>
    public PdfInlineImage(
        byte[] imageBytes,
        double width,
        double height,
        string? alternativeText = null,
        OfficeImageFit fit = OfficeImageFit.Contain,
        double baselineOffset = 0D)
        : base(width, height, baselineOffset, alternativeText) {
        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(imageBytes);
        var style = new PdfImageStyle {
            Fit = fit,
            AlternativeText = alternativeText
        };
        PdfDocument.ValidateImageFitDimensions(prepared.Info, fit, nameof(fit));
        Block = new ImageBlock(prepared.Data, width, height, prepared.Info, style, useDataSnapshot: true);
        Fit = fit;
    }
}

/// <summary>
/// Fixed-size filled and/or bordered box that flows between text runs.
/// </summary>
public sealed class PdfInlineBox : PdfInlineElement {
    /// <summary>Optional fill color.</summary>
    public PdfColor? Background { get; }

    /// <summary>Optional border color.</summary>
    public PdfColor? BorderColor { get; }

    /// <summary>Border width in points.</summary>
    public double BorderWidth { get; }

    /// <summary>Creates a fixed-size inline box.</summary>
    public PdfInlineBox(
        double width,
        double height,
        PdfColor? background = null,
        PdfColor? borderColor = null,
        double borderWidth = 0.5D,
        string? alternativeText = null,
        double baselineOffset = 0D)
        : base(width, height, baselineOffset, alternativeText) {
        if (borderWidth < 0D || double.IsNaN(borderWidth) || double.IsInfinity(borderWidth)) {
            throw new ArgumentException("Inline box border width must be a non-negative finite value.", nameof(borderWidth));
        }

        Background = background;
        BorderColor = borderColor;
        BorderWidth = borderWidth;
    }
}
