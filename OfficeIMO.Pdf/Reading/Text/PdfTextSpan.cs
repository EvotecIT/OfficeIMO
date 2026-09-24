using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// A piece of text extracted from a PDF page with basic font and position info.
/// Coordinates are in user space units (points) as emitted by content stream Tm/Td.
/// </summary>
public sealed class PdfTextSpan {
    /// <summary>Text content of the span.</summary>
    public string Text { get; }
    /// <summary>Font resource name from the page resources (e.g., F1).</summary>
    public string FontResource { get; }
    /// <summary>PDF base font name resolved from the active resource dictionary, when available.</summary>
    public string? BaseFont { get; }
    /// <summary>Font descriptor weight from 1 through 1000, when declared by the PDF.</summary>
    public int? FontWeight { get; }
    /// <summary>Raw PDF font descriptor flags, when declared by the PDF.</summary>
    public int? FontDescriptorFlags { get; }
    /// <summary>True when declared weight or legacy font-name evidence identifies a bold face.</summary>
    public bool IsBold => PdfFontStyleEvidence.IsBold(BaseFont, FontWeight);
    /// <summary>True when descriptor flags or legacy font-name evidence identifies an italic face.</summary>
    public bool IsItalic => PdfFontStyleEvidence.IsItalic(BaseFont, FontDescriptorFlags);
    /// <summary>Font size in points.</summary>
    public double FontSize { get; }
    /// <summary>X position (points) in page user space.</summary>
    public double X { get; }
    /// <summary>Y position (points) in page user space.</summary>
    public double Y { get; }
    /// <summary>Advance width in user space for this span (includes spacing and hscale).</summary>
    public double Advance { get; }
    /// <summary>Fill color used for visual text rendering, when it could be read from the content stream.</summary>
    public OfficeColor? Color { get; }
    /// <summary>True when the text span should be painted during visual rendering.</summary>
    public bool IsVisible { get; }
    /// <summary>Baseline rotation in PDF user-space degrees, where positive angles are counter-clockwise.</summary>
    public double RotationDegrees { get; }
    /// <summary>Active content-stream-scoped marked-content identifier from a tagged PDF, when present.</summary>
    public int? MarkedContentId { get; }
    /// <summary>Form XObject or other content-stream object owning the MCID, or null for page content.</summary>
    public int? ContentStreamObjectNumber { get; }
    /// <summary>Optional visual clipping path in page top-left coordinates.</summary>
    internal PdfPageClipPath? ClipPath { get; }
    internal string? DrawingFontFamily { get; }
    internal double PaintOrder { get; }
    internal PdfContentOrderKey? ContentOrderKey { get; }
    internal PdfContentOrderKey? TextObjectOrderKey { get; }
    internal int LogicalLineBreaksBefore { get; }
    internal bool LogicalLeadingSpace { get; }
    internal bool LogicalTrailingSpace { get; }
    internal IReadOnlyList<int>? EmbeddedLineBreakCounts { get; }
    internal IReadOnlyList<double>? CharacterAdvances { get; }
    internal IReadOnlyList<int>? GlyphCharacterLengths { get; }
    internal IReadOnlyList<byte[]>? GlyphBytes { get; }
    internal IReadOnlyList<double>? GlyphPaintedAdvances { get; }
    internal double CharacterAdvanceDirection { get; }
    /// <summary>PDF text rendering mode active for this span (0 through 7).</summary>
    public int TextRenderingMode { get; }
    internal bool CanRestamp { get; }
    internal bool CanScaleAggregateAdvance { get; }
    internal double RestampFontSize { get; }
    internal string RestampText { get; }
    internal Matrix2D? TextToPageTransform { get; }
    internal string? VisualPaintIdentity { get; }
    internal bool HasActualText { get; }
    internal bool IsType3Font { get; }
    internal bool GlyphSequenceProgressesLeftToRight { get; }
    internal bool IsPaintedGlyphProjection { get; private set; }
    internal string? LogicalDrawingText { get; private set; }
    internal void MarkPaintedGlyphProjection() => IsPaintedGlyphProjection = true;
    /// <summary>True when this span came from PDF content explicitly marked as an /Artifact.</summary>
    public bool IsArtifactContent { get; }
    /// <summary>Creates a new text span.</summary>
    public PdfTextSpan(
        string text,
        string fontResource,
        double fontSize,
        double x,
        double y,
        double advance = 0,
        OfficeColor? color = null,
        bool isVisible = true,
        double rotationDegrees = 0D,
        string? baseFont = null)
        : this(
            text,
            fontResource,
            fontSize,
            x,
            y,
            advance,
            color,
            isVisible,
            rotationDegrees,
            baseFont,
            null,
            fontWeight: null,
            fontDescriptorFlags: null) {
    }

    /// <summary>Creates a new text span with explicit PDF font descriptor style evidence.</summary>
    public PdfTextSpan(
        string text,
        string fontResource,
        double fontSize,
        double x,
        double y,
        double advance,
        OfficeColor? color,
        bool isVisible,
        double rotationDegrees,
        string? baseFont,
        int? fontWeight,
        int? fontDescriptorFlags)
        : this(
            text,
            fontResource,
            fontSize,
            x,
            y,
            advance,
            color,
            isVisible,
            rotationDegrees,
            baseFont,
            null,
            fontWeight: fontWeight,
            fontDescriptorFlags: fontDescriptorFlags) {
    }

    internal PdfTextSpan(
        string text,
        string fontResource,
        double fontSize,
        double x,
        double y,
        double advance,
        OfficeColor? color,
        bool isVisible,
        double rotationDegrees,
        string? baseFont,
        PdfPageClipPath? clipPath,
        double paintOrder = 0D,
        string? drawingFontFamily = null,
        int logicalLineBreaksBefore = 0,
        bool logicalLeadingSpace = false,
        bool logicalTrailingSpace = false,
        PdfContentOrderKey? contentOrderKey = null,
        IReadOnlyList<double>? characterAdvances = null,
        int textRenderingMode = 0,
        bool canRestamp = true,
        double? restampFontSize = null,
        string? restampText = null,
        bool canScaleAggregateAdvance = true,
        int? markedContentId = null,
        int? contentStreamObjectNumber = null,
        PdfContentOrderKey? textObjectOrderKey = null,
        Matrix2D? textToPageTransform = null,
        string? visualPaintIdentity = null,
        IReadOnlyList<int>? glyphCharacterLengths = null,
        IReadOnlyList<byte[]>? glyphBytes = null,
        IReadOnlyList<double>? glyphPaintedAdvances = null,
        double characterAdvanceDirection = 0D,
        bool hasActualText = false,
        bool isType3Font = false,
        bool glyphSequenceProgressesLeftToRight = false,
        bool isArtifactContent = false,
        int? fontWeight = null,
        int? fontDescriptorFlags = null,
        IReadOnlyList<int>? embeddedLineBreakCounts = null) {
        if (fontWeight.HasValue && (fontWeight.Value < 1 || fontWeight.Value > 1000)) {
            throw new ArgumentOutOfRangeException(nameof(fontWeight));
        }
        if (fontDescriptorFlags < 0) throw new ArgumentOutOfRangeException(nameof(fontDescriptorFlags));
        Text = text;
        FontResource = fontResource;
        BaseFont = baseFont;
        FontWeight = fontWeight;
        FontDescriptorFlags = fontDescriptorFlags;
        FontSize = fontSize;
        X = x;
        Y = y;
        Advance = advance;
        Color = color;
        IsVisible = isVisible;
        RotationDegrees = rotationDegrees;
        MarkedContentId = markedContentId;
        ContentStreamObjectNumber = contentStreamObjectNumber;
        ClipPath = clipPath;
        PaintOrder = paintOrder;
        DrawingFontFamily = drawingFontFamily;
        LogicalLineBreaksBefore = logicalLineBreaksBefore;
        LogicalLeadingSpace = logicalLeadingSpace;
        LogicalTrailingSpace = logicalTrailingSpace;
        EmbeddedLineBreakCounts = embeddedLineBreakCounts;
        ContentOrderKey = contentOrderKey;
        TextObjectOrderKey = textObjectOrderKey;
        CharacterAdvances = characterAdvances?.ToArray();
        GlyphCharacterLengths = glyphCharacterLengths?.ToArray();
        GlyphBytes = glyphBytes?.Select(static bytes => bytes.ToArray()).ToArray();
        GlyphPaintedAdvances = glyphPaintedAdvances?.ToArray();
        CharacterAdvanceDirection = characterAdvanceDirection;
        TextRenderingMode = textRenderingMode;
        CanRestamp = canRestamp;
        CanScaleAggregateAdvance = canScaleAggregateAdvance;
        RestampFontSize = restampFontSize ?? fontSize;
        RestampText = restampText ?? text;
        TextToPageTransform = textToPageTransform;
        VisualPaintIdentity = visualPaintIdentity;
        HasActualText = hasActualText;
        IsType3Font = isType3Font;
        GlyphSequenceProgressesLeftToRight = glyphSequenceProgressesLeftToRight;
        IsArtifactContent = isArtifactContent;
    }

    private bool _completeTextProjectionSuppressed;

    internal PdfTextSpan WithCanRestamp(bool canRestamp) => new PdfTextSpan(
        Text, FontResource, FontSize, X, Y, Advance, Color, IsVisible, RotationDegrees, BaseFont, ClipPath,
        PaintOrder, DrawingFontFamily, LogicalLineBreaksBefore, LogicalLeadingSpace, LogicalTrailingSpace,
        ContentOrderKey, CharacterAdvances, TextRenderingMode, canRestamp, RestampFontSize, RestampText, CanScaleAggregateAdvance, MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, TextToPageTransform, VisualPaintIdentity, GlyphCharacterLengths, GlyphBytes, GlyphPaintedAdvances, CharacterAdvanceDirection, HasActualText, IsType3Font, GlyphSequenceProgressesLeftToRight, IsArtifactContent, FontWeight, FontDescriptorFlags, EmbeddedLineBreakCounts) { _completeTextProjectionSuppressed = this._completeTextProjectionSuppressed, LogicalDrawingText = LogicalDrawingText, IsPaintedGlyphProjection = IsPaintedGlyphProjection };

    internal PdfTextSpan WithOffset(double deltaX, double deltaY) => new PdfTextSpan(
        Text, FontResource, FontSize, X + deltaX, Y + deltaY, Advance, Color, IsVisible, RotationDegrees, BaseFont, ClipPath,
        PaintOrder, DrawingFontFamily, LogicalLineBreaksBefore, LogicalLeadingSpace, LogicalTrailingSpace,
        ContentOrderKey, CharacterAdvances, TextRenderingMode, CanRestamp, RestampFontSize, RestampText, CanScaleAggregateAdvance, MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, TextToPageTransform, VisualPaintIdentity, GlyphCharacterLengths, GlyphBytes, GlyphPaintedAdvances, CharacterAdvanceDirection, HasActualText, IsType3Font, GlyphSequenceProgressesLeftToRight, IsArtifactContent, FontWeight, FontDescriptorFlags, EmbeddedLineBreakCounts) { _completeTextProjectionSuppressed = this._completeTextProjectionSuppressed, LogicalDrawingText = LogicalDrawingText, IsPaintedGlyphProjection = IsPaintedGlyphProjection };

    internal PdfTextSpan WithVisualFontSize(double fontSize) => new PdfTextSpan(
        Text, FontResource, fontSize, X, Y, Advance, Color, IsVisible, RotationDegrees, BaseFont, ClipPath,
        PaintOrder, DrawingFontFamily, LogicalLineBreaksBefore, LogicalLeadingSpace, LogicalTrailingSpace,
        ContentOrderKey, CharacterAdvances, TextRenderingMode, CanRestamp, RestampFontSize, RestampText, CanScaleAggregateAdvance, MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, TextToPageTransform, VisualPaintIdentity, GlyphCharacterLengths, GlyphBytes, GlyphPaintedAdvances, CharacterAdvanceDirection, HasActualText, IsType3Font, GlyphSequenceProgressesLeftToRight, IsArtifactContent, FontWeight, FontDescriptorFlags, EmbeddedLineBreakCounts) { _completeTextProjectionSuppressed = this._completeTextProjectionSuppressed, LogicalDrawingText = LogicalDrawingText, IsPaintedGlyphProjection = IsPaintedGlyphProjection };

    /// <summary>Replaces displayed text with its painted form while retaining editable logical text.</summary>
    internal PdfTextSpan WithVisualText(string text) => new PdfTextSpan(
        text, FontResource, FontSize, X, Y, Advance, Color, IsVisible, RotationDegrees, BaseFont, ClipPath,
        PaintOrder, DrawingFontFamily, LogicalLineBreaksBefore, LogicalLeadingSpace, LogicalTrailingSpace,
        ContentOrderKey, CharacterAdvances, TextRenderingMode, CanRestamp, RestampFontSize, RestampText, CanScaleAggregateAdvance, MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, TextToPageTransform, VisualPaintIdentity, GlyphCharacterLengths, GlyphBytes, GlyphPaintedAdvances, CharacterAdvanceDirection, HasActualText, IsType3Font, GlyphSequenceProgressesLeftToRight, IsArtifactContent, FontWeight, FontDescriptorFlags, text.Length == Text.Length ? EmbeddedLineBreakCounts : null) { _completeTextProjectionSuppressed = this._completeTextProjectionSuppressed, LogicalDrawingText = LogicalDrawingText ?? Text, IsPaintedGlyphProjection = IsPaintedGlyphProjection };

    /// <summary>
    /// Returns one painted glyph of this run at its own origin. The glyph advance is its painted
    /// width, so the slice can be scaled to its advance even when the run used character spacing.
    /// </summary>
    internal PdfTextSpan WithPaintedGlyph(int glyphIndex, int characterOffset, double x, double y) {
        int characterLength = GlyphCharacterLengths![glyphIndex];
        string text = Text.Substring(characterOffset, characterLength);
        double paintedAdvance = GlyphPaintedAdvances![glyphIndex];
        var characterAdvances = new double[characterLength];
        for (int index = 0; index < characterLength; index++) characterAdvances[index] = CharacterAdvances![characterOffset + index];
        return new PdfTextSpan(
            text, FontResource, FontSize, x, y, paintedAdvance, Color, IsVisible, RotationDegrees, BaseFont, ClipPath,
            PaintOrder, DrawingFontFamily, 0, false, false,
            ContentOrderKey, characterAdvances, TextRenderingMode, false, RestampFontSize, text, true, MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, null, VisualPaintIdentity,
            new[] { characterLength }, new[] { GlyphBytes![glyphIndex] }, new[] { paintedAdvance }, CharacterAdvanceDirection, HasActualText, IsType3Font, false, IsArtifactContent, FontWeight, FontDescriptorFlags) {
            _completeTextProjectionSuppressed = this._completeTextProjectionSuppressed,
            LogicalDrawingText = LogicalDrawingText?.Substring(characterOffset, characterLength),
            IsPaintedGlyphProjection = IsPaintedGlyphProjection
        };
    }

    /// <summary>Displays a single-glyph run as one Unicode scalar, such as an alias for a ligature glyph.</summary>
    internal PdfTextSpan WithVisualGlyph(int glyphScalar) {
        string glyphText = char.ConvertFromUtf32(glyphScalar);
        double[]? advances = CharacterAdvances == null ? null : glyphText.Length == 1
            ? new[] { CharacterAdvances.Sum() } : new[] { CharacterAdvances.Sum(), 0D };
        return new PdfTextSpan(
            glyphText, FontResource, FontSize, X, Y, Advance, Color, IsVisible, RotationDegrees, BaseFont, ClipPath,
            PaintOrder, DrawingFontFamily, LogicalLineBreaksBefore, LogicalLeadingSpace, LogicalTrailingSpace,
            ContentOrderKey, advances, TextRenderingMode, CanRestamp, RestampFontSize, RestampText, CanScaleAggregateAdvance, MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, TextToPageTransform, VisualPaintIdentity, new[] { glyphText.Length }, GlyphBytes, GlyphPaintedAdvances, CharacterAdvanceDirection, HasActualText, IsType3Font, GlyphSequenceProgressesLeftToRight, IsArtifactContent, FontWeight, FontDescriptorFlags, glyphText.Length == Text.Length ? EmbeddedLineBreakCounts : null) { _completeTextProjectionSuppressed = this._completeTextProjectionSuppressed, LogicalDrawingText = LogicalDrawingText ?? Text, IsPaintedGlyphProjection = IsPaintedGlyphProjection };
    }

    // Layout and rendering measure the transformed glyph height. Raw extraction still exposes
    // the authored Tf operand, which can be 1 when a producer puts scaling in Tm or cm.
    internal PdfTextSpan WithPageFontSize() =>
        RestampFontSize > 0D && !double.IsNaN(RestampFontSize) && !double.IsInfinity(RestampFontSize) &&
        Math.Abs(RestampFontSize - FontSize) > 0.000001D
            ? WithVisualFontSize(RestampFontSize)
            : this;

    internal PdfTextSpan WithLayoutGeometry(double x, double y, double angle, double sourcePageHeight) => new PdfTextSpan(
        Text, FontResource, FontSize, x, y, Advance, Color, IsVisible, angle, BaseFont, null,
        PaintOrder, DrawingFontFamily, LogicalLineBreaksBefore, LogicalLeadingSpace, LogicalTrailingSpace,
        ContentOrderKey, CharacterAdvances, TextRenderingMode, false, RestampFontSize, RestampText, CanScaleAggregateAdvance,
        MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, null, VisualPaintIdentity, GlyphCharacterLengths,
        GlyphBytes, GlyphPaintedAdvances, CharacterAdvanceDirection, HasActualText, IsType3Font,
        GlyphSequenceProgressesLeftToRight, IsArtifactContent, FontWeight, FontDescriptorFlags, EmbeddedLineBreakCounts) {
        // The internal layout proxy has no source-space clip. Retain the original eligibility
        // decision so removing that coordinate-dependent clip cannot expose complete cell text.
        _completeTextProjectionSuppressed = !CanProjectCompleteText(sourcePageHeight),
        LogicalDrawingText = LogicalDrawingText,
        IsPaintedGlyphProjection = IsPaintedGlyphProjection
    };

    internal bool CanProjectCompleteText(double? pageHeight) {
        if (_completeTextProjectionSuppressed || !IsVisible || Color?.A <= 3 || string.IsNullOrEmpty(Text)) return false;
        if (!ClipPath.HasValue) return true;
        if (!pageHeight.HasValue || Math.Abs(RotationDegrees) > 0.01D) return false;

        PdfPageClipPath clip = ClipPath.Value;
        if (!clip.IsRectangle || !clip.IsExact || clip.ContainsTextClipping) return false;
        // Extracted spans expose a baseline rather than font ascent/descent metrics. Use the painted
        // glyph box approximation here so tight producer cell clips do not look like partial text.
        const double approximateAscentFactor = 0.8D;
        double width = Advance > 0D ? Advance : PdfUnicodeScalarAnalysis.CountScalars(Text) * FontSize * 0.55D;
        double height = Math.Max(1D, FontSize);
        double left = X;
        double top = pageHeight.Value - Y - FontSize * approximateAscentFactor;
        const double tolerance = 0.05D;
        return left + tolerance >= clip.X &&
               top + tolerance >= clip.Y &&
               left + width <= clip.X + clip.Width + tolerance &&
               top + height <= clip.Y + clip.Height + tolerance;
    }

    internal bool CanProjectCompleteText(PdfLogicalPage page) {
        return CanProjectCompleteText(page.Height) && IntersectsPageBoundary(page);
    }

    internal bool CanProjectPositionedHtmlText(PdfLogicalPage page) {
        if (!IsVisible || Color?.A <= 3 || string.IsNullOrEmpty(Text) ||
            !IntersectsPageBoundary(page)) return false;
        if (!ClipPath.HasValue) return !_completeTextProjectionSuppressed;
        if (Math.Abs(RotationDegrees) > 0.01D) return false;

        PdfPageClipPath clip = ClipPath.Value;
        if (!clip.IsRectangle || !clip.IsExact || clip.ContainsTextClipping) return false;
        // Positioned review uses a browser font. The cap-height approximation keeps
        // legitimately visible canvas text inside tight producer rectangles.
        double width = Advance > 0D ? Advance : PdfUnicodeScalarAnalysis.CountScalars(Text) * FontSize * 0.55D;
        double top = page.Height - Y - FontSize * 0.7D;
        const double tolerance = 0.05D;
        return X + tolerance >= clip.X &&
            top + tolerance >= clip.Y &&
            X + width <= clip.X + clip.Width + tolerance &&
            top + Math.Max(1D, FontSize) <= clip.Y + clip.Height + tolerance;
    }

    internal bool IsCompletelyWithinPageBoundary(PdfLogicalPage page) {
        PdfPageBox? boundary = page.Geometry.EffectiveBox;
        if (boundary == null) return false;
        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(this);
        const double tolerance = 0.05D;
        return bounds.Left >= boundary.Left - tolerance &&
            bounds.Right <= boundary.Right + tolerance &&
            bounds.Bottom >= boundary.Bottom - tolerance &&
            bounds.Top <= boundary.Top + tolerance;
    }

    internal bool IntersectsPageBoundary(PdfLogicalPage page) {
        PdfPageBox? boundary = page.Geometry.EffectiveBox;
        if (boundary == null) return false;
        PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(this);
        const double tolerance = 0.05D;
        return bounds.Right > boundary.Left + tolerance &&
            bounds.Left < boundary.Right - tolerance &&
            bounds.Top > boundary.Bottom + tolerance &&
            bounds.Bottom < boundary.Top - tolerance;
    }
}
