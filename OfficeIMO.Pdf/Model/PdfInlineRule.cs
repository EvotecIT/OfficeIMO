namespace OfficeIMO.Pdf;

/// <summary>
/// Horizontal ink band that fills the remaining rich-text line, optionally limited to a maximum width.
/// </summary>
/// <remarks>
/// The rule has zero intrinsic width. Each layout resolves a fresh band against
/// its actual remaining frame, preserving the source run's font line metrics.
/// Text following a rule wraps normally when the rule consumes the line.
/// </remarks>
public sealed class PdfInlineRule : PdfInlineElement {
    private readonly InkBand[] _bands;
    internal IReadOnlyList<InkBand> Bands { get; }

    internal readonly struct InkBand {
        internal double Height { get; }
        internal PdfColor Color { get; }
        internal double BaselineOffset { get; }

        internal InkBand(double height, PdfColor color, double baselineOffset) {
            Guard.Positive(height, nameof(height));
            if (double.IsNaN(baselineOffset) || double.IsInfinity(baselineOffset))
                throw new ArgumentException("Inline rule band baseline offset must be finite.", nameof(baselineOffset));
            Height = height;
            Color = color;
            BaselineOffset = baselineOffset;
        }
    }

    /// <summary>Color of the rule's filled band.</summary>
    public PdfColor Color { get; }

    /// <summary>Optional positive maximum width in points. Null fills all remaining line width.</summary>
    public double? MaximumWidth { get; }

    /// <summary>Creates a flexible inline rule with explicit band height and baseline offset, in points.</summary>
    public PdfInlineRule(double height, PdfColor color, double? maximumWidth = null, double baselineOffset = 0D)
        : base(0D, height, baselineOffset, null, flexibleWidth: true) {
        if (maximumWidth.HasValue) Guard.Positive(maximumWidth.Value, nameof(maximumWidth));
        Color = color;
        MaximumWidth = maximumWidth;
        _bands = new[] { new InkBand(height, color, baselineOffset) };
        Bands = Array.AsReadOnly(_bands);
    }

    private PdfInlineRule(PdfInlineRule source, InkBand additionalBand)
        : base(0D, source.Height, source.BaselineOffset, null, flexibleWidth: true) {
        Color = source.Color;
        MaximumWidth = source.MaximumWidth;
        _bands = new InkBand[source._bands.Length + 1];
        Array.Copy(source._bands, _bands, source._bands.Length);
        _bands[_bands.Length - 1] = additionalBand;
        Bands = Array.AsReadOnly(_bands);
    }

    /// <summary>Returns a new immutable rule with another ink band sharing the resolved width.</summary>
    internal PdfInlineRule WithAdditionalBand(double height, PdfColor color, double baselineOffset) =>
        new(this, new InkBand(height, color, baselineOffset));
}
