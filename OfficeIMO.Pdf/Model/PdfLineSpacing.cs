namespace OfficeIMO.Pdf;

/// <summary>Immutable paragraph line spacing with explicit relative or point units.</summary>
public sealed class PdfLineSpacing {
    private PdfLineSpacing(PdfLineSpacingRule rule, double value, double naturalMultiplier) {
        if (value < 0D || (value == 0D && rule != PdfLineSpacingRule.AtLeast) || double.IsNaN(value) || double.IsInfinity(value))
            throw new System.ArgumentOutOfRangeException(nameof(value), "Line spacing must be finite and non-negative; proportional and exact spacing must be positive.");
        if (naturalMultiplier <= 0D || double.IsNaN(naturalMultiplier) || double.IsInfinity(naturalMultiplier))
            throw new System.ArgumentOutOfRangeException(nameof(naturalMultiplier), "Natural line height must be a positive finite multiplier.");
        Rule = rule;
        Value = value;
        NaturalMultiplier = naturalMultiplier;
    }

    /// <summary>The spacing rule that determines how <see cref="Value"/> is interpreted.</summary>
    public PdfLineSpacingRule Rule { get; }

    /// <summary>The multiplier for <see cref="PdfLineSpacingRule.Multiple"/>, or points for the other rules.</summary>
    public double Value { get; }

    /// <summary>Natural line advance divided by font size. Used by minimum spacing and equal to <see cref="Value"/> for multiple spacing.</summary>
    public double NaturalMultiplier { get; }

    /// <summary>Creates spacing proportional to each line's font size.</summary>
    /// <param name="multiplier">Line advance divided by font size.</param>
    public static PdfLineSpacing Multiple(double multiplier) => new(PdfLineSpacingRule.Multiple, multiplier, multiplier);

    /// <summary>Creates a fixed line advance in points, independent of font and inline element sizes.</summary>
    /// <param name="points">Line advance in points. Text may extend beyond the line box.</param>
    public static PdfLineSpacing Exactly(double points) => new(PdfLineSpacingRule.Exact, points, 1D);

    /// <summary>Creates a minimum line advance in points that expands for larger text or inline elements.</summary>
    /// <param name="points">Non-negative minimum advance in points. Zero uses the natural line height.</param>
    /// <param name="naturalMultiplier">Natural line advance divided by font size, used when text exceeds the minimum.</param>
    public static PdfLineSpacing AtLeast(double points, double naturalMultiplier = 1.4D) => new(PdfLineSpacingRule.AtLeast, points, naturalMultiplier);

    internal bool IsExact => Rule == PdfLineSpacingRule.Exact;

    // Document formats may position a baseline in the font's natural line box
    // rather than align its ascender to the top. Keep this independent of the
    // requested multiple, which changes advance but not that initial box.
    internal double? FontLineBoxMultiplier { get; private init; }

    // Some document formats place exact-height text at a fixed point offset
    // within the authored line box, independent of the visible run sizes.
    internal double? FixedLineBoxBaselineOffset { get; private init; }

    internal PdfLineSpacing WithFixedLineBoxBaseline(double baselineOffset) {
        if (!IsExact || baselineOffset < 0D || double.IsNaN(baselineOffset) || double.IsInfinity(baselineOffset))
            throw new System.ArgumentOutOfRangeException(nameof(baselineOffset), "A fixed line-box baseline requires exact spacing and a finite non-negative offset.");
        return new(Rule, Value, NaturalMultiplier) {
            FontLineBoxMultiplier = FontLineBoxMultiplier, FixedLineBoxBaselineOffset = baselineOffset
        };
    }

    internal PdfLineSpacing WithFontLineBoxBaseline(double naturalMultiplier) =>
        new(Rule, Value, NaturalMultiplier) {
            FontLineBoxMultiplier = naturalMultiplier, FixedLineBoxBaselineOffset = FixedLineBoxBaselineOffset
        };

    internal double GetAdvance(double fontSize) => Rule switch {
        PdfLineSpacingRule.Exact => Value,
        PdfLineSpacingRule.AtLeast => System.Math.Max(Value, fontSize * NaturalMultiplier),
        _ => fontSize * Value
    };
}
