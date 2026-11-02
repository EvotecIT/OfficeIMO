namespace OfficeIMO.Pdf;

/// <summary>Immutable paragraph line spacing with explicit relative or point units.</summary>
public sealed class PdfLineSpacing {
    private readonly double _multiple;
    private readonly global::OfficeIMO.OfficeLength _height;

    private PdfLineSpacing(PdfLineSpacingRule rule, double multiple, global::OfficeIMO.OfficeLength height, double naturalMultiplier) {
        if (rule == PdfLineSpacingRule.Multiple) {
            if (multiple <= 0D || double.IsNaN(multiple) || double.IsInfinity(multiple))
                throw new System.ArgumentOutOfRangeException(nameof(multiple), "Line spacing must be a positive finite multiple.");
        } else if (height < global::OfficeIMO.OfficeLength.Zero || (height == global::OfficeIMO.OfficeLength.Zero && rule == PdfLineSpacingRule.Exact)) {
            throw new System.ArgumentOutOfRangeException(nameof(height), "Minimum line height must be non-negative; exact line height must be positive.");
        }
        if (naturalMultiplier <= 0D || double.IsNaN(naturalMultiplier) || double.IsInfinity(naturalMultiplier))
            throw new System.ArgumentOutOfRangeException(nameof(naturalMultiplier), "Natural line height must be a positive finite multiplier.");
        Rule = rule;
        _multiple = multiple;
        _height = height;
        NaturalMultiplier = naturalMultiplier;
    }

    /// <summary>The spacing rule, assigned together with its physical height or line multiple.</summary>
    public PdfLineSpacingRule Rule { get; }

    /// <summary>Gets the line multiple for proportional spacing, otherwise null.</summary>
    public double? MultipleValue => Rule == PdfLineSpacingRule.Multiple ? _multiple : (double?)null;

    /// <summary>Gets the physical height for exact or minimum spacing, otherwise null.</summary>
    public global::OfficeIMO.OfficeLength? Height => Rule == PdfLineSpacingRule.Multiple ? null : _height;

    /// <summary>Natural line advance divided by font size. For proportional spacing, equal to the line multiple.</summary>
    public double NaturalMultiplier { get; }

    /// <summary>Creates spacing proportional to each line's font size.</summary>
    /// <param name="multiplier">Positive finite line advance divided by font size.</param>
    public static PdfLineSpacing Multiple(double multiplier) => new(PdfLineSpacingRule.Multiple, multiplier, default, multiplier);

    /// <summary>Creates a fixed physical line advance, independent of font and inline element sizes.</summary>
    /// <param name="height">Positive line height. Text may extend beyond the line box.</param>
    public static PdfLineSpacing Exact(global::OfficeIMO.OfficeLength height) => new(PdfLineSpacingRule.Exact, 0D, height, 1D);

    /// <summary>Creates a minimum physical line advance that expands for larger text or inline elements.</summary>
    /// <param name="height">Non-negative minimum height. Zero uses the natural line height.</param>
    /// <param name="naturalMultiplier">Natural line advance divided by font size, used when text exceeds the minimum.</param>
    public static PdfLineSpacing AtLeast(global::OfficeIMO.OfficeLength height, double naturalMultiplier = 1.4D) => new(PdfLineSpacingRule.AtLeast, 0D, height, naturalMultiplier);

    internal bool IsExact => Rule == PdfLineSpacingRule.Exact;

    // Document formats may position a baseline in the font's natural line box
    // rather than align its ascender to the top. Keep this independent of the
    // requested multiple, which changes advance but not that initial box.
    internal double? FontLineBoxMultiplier { get; private init; }

    // Some document formats place exact-height text at a fixed point offset
    // within the authored line box, independent of the visible run sizes.
    internal double? FixedLineBoxBaselineOffset { get; private init; }

    // A document line grid constrains natural line boxes before applying
    // multiple or minimum spacing. Exact spacing bypasses the grid.
    internal double? LineGridPitch { get; private init; }
    internal PdfFontLineAlignment FontAlignment { get; private init; }
    internal double FontAlignmentPointsPerPixel { get; private init; }

    internal PdfLineSpacing WithFontAlignment(PdfFontLineAlignment alignment, double pointsPerPixel) {
        if (pointsPerPixel <= 0D || double.IsNaN(pointsPerPixel) || double.IsInfinity(pointsPerPixel))
            throw new System.ArgumentOutOfRangeException(nameof(pointsPerPixel));
        return new(Rule, _multiple, _height, NaturalMultiplier) {
            FontLineBoxMultiplier = FontLineBoxMultiplier, FixedLineBoxBaselineOffset = FixedLineBoxBaselineOffset,
            LineGridPitch = LineGridPitch, FontAlignment = alignment, FontAlignmentPointsPerPixel = pointsPerPixel
        };
    }

    internal PdfLineSpacing WithLineGrid(double pitch) {
        if (pitch <= 0D || double.IsNaN(pitch) || double.IsInfinity(pitch))
            throw new System.ArgumentOutOfRangeException(nameof(pitch));
        return new(Rule, _multiple, _height, NaturalMultiplier) {
            FontLineBoxMultiplier = FontLineBoxMultiplier, FixedLineBoxBaselineOffset = FixedLineBoxBaselineOffset,
            LineGridPitch = IsExact ? null : pitch,
            FontAlignment = FontAlignment, FontAlignmentPointsPerPixel = FontAlignmentPointsPerPixel
        };
    }

    internal PdfLineSpacing WithFixedLineBoxBaseline(double baselineOffset) {
        if (!IsExact || baselineOffset < 0D || double.IsNaN(baselineOffset) || double.IsInfinity(baselineOffset))
            throw new System.ArgumentOutOfRangeException(nameof(baselineOffset), "A fixed line-box baseline requires exact spacing and a finite non-negative offset.");
        return new(Rule, _multiple, _height, NaturalMultiplier) {
            FontLineBoxMultiplier = FontLineBoxMultiplier, FixedLineBoxBaselineOffset = baselineOffset, LineGridPitch = LineGridPitch,
            FontAlignment = FontAlignment, FontAlignmentPointsPerPixel = FontAlignmentPointsPerPixel
        };
    }

    internal PdfLineSpacing WithFontLineBoxBaseline(double naturalMultiplier) =>
        new(Rule, _multiple, _height, NaturalMultiplier) {
            FontLineBoxMultiplier = naturalMultiplier, FixedLineBoxBaselineOffset = FixedLineBoxBaselineOffset, LineGridPitch = LineGridPitch,
            FontAlignment = FontAlignment, FontAlignmentPointsPerPixel = FontAlignmentPointsPerPixel
        };

    internal double GetAdvance(double fontSize) => LineGridPitch.HasValue && !IsExact
        ? GetGridAdvance(fontSize * (FontLineBoxMultiplier ?? NaturalMultiplier)) : Rule switch {
        PdfLineSpacingRule.Exact => _height.ToPoints(),
        PdfLineSpacingRule.AtLeast => System.Math.Max(_height.ToPoints(), fontSize * NaturalMultiplier),
        _ => fontSize * _multiple
    };

    internal double GetGridNaturalHeight(double naturalHeight) => LineGridPitch is double pitch
        ? System.Math.Ceiling(System.Math.Max(0D, naturalHeight) / pitch - 1e-9D) * pitch : naturalHeight;

    internal double GetGridAdvance(double naturalHeight) => System.Math.Max(GetGridNaturalHeight(naturalHeight),
        Rule == PdfLineSpacingRule.AtLeast ? _height.ToPoints() : LineGridPitch!.Value * _multiple / (FontLineBoxMultiplier ?? NaturalMultiplier));

    internal double GetGridBaselineExpansion(double naturalHeight) {
        double gridHeight = GetGridNaturalHeight(naturalHeight);
        return Rule == PdfLineSpacingRule.AtLeast
            ? (gridHeight - naturalHeight) / 2D + System.Math.Max(0D, _height.ToPoints() - gridHeight)
            : (GetGridAdvance(naturalHeight) - naturalHeight) / 2D;
    }
}
