using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Base class for a visual that participates in rich-paragraph wrapping.
/// </summary>
public abstract class PdfInlineElement {
    /// <summary>Intrinsic width of the inline element in points. Flexible rules resolve their zero intrinsic width during wrapping.</summary>
    public global::OfficeIMO.OfficeLength Width { get; }

    /// <summary>Height of the inline element in points.</summary>
    public global::OfficeIMO.OfficeLength Height { get; }

    /// <summary>Vertical offset of the element bottom from the text baseline, in points.</summary>
    public global::OfficeIMO.OfficeLength BaselineOffset { get; }

    /// <summary>Alternative text used when the element is emitted into a tagged document.</summary>
    public string? AlternativeText { get; }

    private protected PdfInlineElement(global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeLength baselineOffset, string? alternativeText)
        : this(width, height, baselineOffset, alternativeText, flexibleWidth: false) { }

    private protected PdfInlineElement(global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeLength baselineOffset, string? alternativeText, bool flexibleWidth) {
        if (!flexibleWidth || width.ToPoints() != 0D) Guard.Positive(width.ToPoints(), nameof(width));
        Guard.Positive(height.ToPoints(), nameof(height));
        if (double.IsNaN(baselineOffset.ToPoints()) || double.IsInfinity(baselineOffset.ToPoints())) {
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
    internal int TableTextRotation { get; }
    /// <summary>Imported flow pictures retain their authored frame beyond paragraph bounds.</summary>
    internal bool AllowFrameOverflow { get; init; }
    /// <summary>Clips imported exact-height picture paint to its paragraph fragment without changing frame admission.</summary>
    internal bool ClipToParagraphFragment { get; init; }

    /// <summary>Projects the inline frame onto turned text axes without decoding or changing the image payload.</summary>
    internal PdfInlineImage ForTableTextRotation(int rotation) => rotation == 0 ? this : new PdfInlineImage(this, rotation);

    private PdfInlineImage(PdfInlineImage source, int rotation)
        : base(source.Height, source.Width, source.BaselineOffset, source.AlternativeText) {
        Block = source.Block;
        Fit = source.Fit;
        TableTextRotation = rotation;
        AllowFrameOverflow = source.AllowFrameOverflow;
        ClipToParagraphFragment = source.ClipToParagraphFragment;
    }

    /// <summary>Image fitting mode inside the requested inline frame.</summary>
    public OfficeImageFit Fit { get; }

    /// <summary>Creates an inline image from first-party supported image bytes.</summary>
    public PdfInlineImage(
        byte[] imageBytes,
        global::OfficeIMO.OfficeLength width,
        global::OfficeIMO.OfficeLength height,
        string? alternativeText = null,
        OfficeImageFit fit = OfficeImageFit.Contain,
        global::OfficeIMO.OfficeLength baselineOffset = default)
        : this(imageBytes, width, height, new PdfImageStyle {
            Fit = fit,
            AlternativeText = alternativeText
        }, baselineOffset) {
    }

    /// <summary>Creates an inline image with source crop, rotation, mirrors and alternate text from a reusable image style.</summary>
    /// <remarks>The style is copied. Width and height describe the authored inline frame; rotation does not enlarge the line box.</remarks>
    public PdfInlineImage(byte[] imageBytes, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, PdfImageStyle style, global::OfficeIMO.OfficeLength baselineOffset)
        : base(width, height, baselineOffset, style?.AlternativeText) {
        Guard.NotNull(style, nameof(style));
        PdfImageStyle imageStyle = style!;
        PdfDocument.ValidateImageStyleForBox(imageStyle, width.ToPoints(), height.ToPoints(), nameof(style));
        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(imageBytes);
        PdfDocument.ValidateImageFitDimensions(prepared.Info, imageStyle.Fit, nameof(style));
        Block = new ImageBlock(
            prepared.Data,
            width.ToPoints(),
            height.ToPoints(),
            prepared.Info,
            imageStyle,
            useDataSnapshot: true,
            preparedStream: prepared.PreparedStream);
        Fit = imageStyle.Fit;
    }
}

/// <summary>
/// Fixed-size filled and/or bordered box that flows between text runs.
/// </summary>
public sealed class PdfInlineBox : PdfInlineElement {
    // Imported text spacers belong to the adjacent word when finding the
    // minimum automatic table width; ordinary inline visuals can wrap alone.
    internal bool IsTextSpacer { get; set; }

    // Resolved flexible rules paint a narrow band but retain the surrounding
    // run's font baseline and line-box policy.
    internal bool UsesRunTextMetrics => SourceRule != null;
    internal PdfInlineRule? SourceRule { get; private set; }

    internal static PdfInlineBox FromRule(PdfInlineRule rule, double width) =>
        new PdfInlineBox(global::OfficeIMO.OfficeLength.Points(width), rule.Height, width > 0D ? rule.Color : null,
            null, global::OfficeIMO.OfficeLength.Points(0D), null, rule.BaselineOffset, flexibleWidth: true) { SourceRule = rule };

    /// <summary>Optional fill color.</summary>
    public global::OfficeIMO.OfficeColor? Background { get; }

    /// <summary>Optional border color.</summary>
    public global::OfficeIMO.OfficeColor? BorderColor { get; }

    /// <summary>Border width in points.</summary>
    public global::OfficeIMO.OfficeLength BorderWidth { get; }

    /// <summary>Creates a fixed-size inline box.</summary>
    public PdfInlineBox(
        global::OfficeIMO.OfficeLength width,
        global::OfficeIMO.OfficeLength height,
        global::OfficeIMO.OfficeColor? background = null,
        global::OfficeIMO.OfficeColor? borderColor = null,
        global::OfficeIMO.OfficeLength? borderWidth = null,
        string? alternativeText = null,
        global::OfficeIMO.OfficeLength baselineOffset = default)
        : this(width, height, background, borderColor, (borderWidth ?? global::OfficeIMO.OfficeLength.Points(0.5D)), alternativeText, baselineOffset, flexibleWidth: false) { }

    private PdfInlineBox(global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeColor? background, global::OfficeIMO.OfficeColor? borderColor,
        global::OfficeIMO.OfficeLength borderWidth, string? alternativeText, global::OfficeIMO.OfficeLength baselineOffset, bool flexibleWidth)
        : base(width, height, baselineOffset, alternativeText, flexibleWidth) {
        if (borderWidth.ToPoints() < 0D || double.IsNaN(borderWidth.ToPoints()) || double.IsInfinity(borderWidth.ToPoints())) {
            throw new ArgumentException("Inline box border width must be a non-negative finite value.", nameof(borderWidth));
        }

        Background = background;
        BorderColor = borderColor;
        BorderWidth = borderWidth;
    }
}
