using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using SixLabors.Fonts;
using SixLabors.Fonts.Unicode;

namespace OfficeIMO.Drawing.SixLabors;

internal sealed class OfficeSixLaborsFontProgram : IOfficeBoundedFontProgram {
    private const int MaximumCachedSizes = 256;
    private readonly byte[] _fontData;
    private readonly FontFamily _family;
    private readonly FontStyle _style;
    private readonly FontVariation[] _variations;
    private readonly Dictionary<float, Font> _fontsBySize = new();
    private readonly object _fontSync = new();
    private readonly FontMetrics _metrics;

    internal OfficeSixLaborsFontProgram(
        byte[] fontData,
        FontFamily family,
        FontStyle style,
        FontVariation[] variations,
        string displayName,
        bool isOpenTypeCff,
        string fingerprint) {
        _fontData = (byte[])fontData.Clone();
        _family = family;
        _style = style;
        _variations = (FontVariation[])variations.Clone();
        DisplayName = displayName;
        IsOpenTypeCff = isOpenTypeCff;
        Fingerprint = fingerprint ?? throw new ArgumentNullException(nameof(fingerprint));
        _metrics = GetFont(1F).FontMetrics;
    }

    public string Fingerprint { get; }

    public string? DisplayName { get; }

    public int? CollectionIndex => null;

    public int UnitsPerEm => _metrics.UnitsPerEm;

    public bool IsOpenTypeCff { get; }

    public bool ProvidesComplexTextLayout => true;

    public double LineSpacingRatio => _metrics.HorizontalMetrics.LineHeight / (double)UnitsPerEm;

    public byte[] GetFontDataForShaping() => (byte[])_fontData.Clone();

    public bool HasGlyphs(string text) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        Font font = GetFont(1F);
        for (int index = 0; index < text.Length;) {
            int scalar = ReadScalar(text, ref index);
            if (OfficeTextElements.IsIgnorableFontCoverageScalar(scalar)) continue;
            if (!font.TryGetGlyphId(new CodePoint(scalar), out ushort glyphId) || glyphId == 0) return false;
        }
        return true;
    }

    public double Measure(string text, double fontSize) {
        ValidateTextAndSize(text, fontSize);
        if (text.Length == 0) return 0D;
        FontRectangle measured = TextMeasurer.MeasureAdvance(text, CreateOptions(fontSize, 0D, 0D));
        return Math.Abs(measured.Width);
    }

    public IReadOnlyList<double> MeasureTextElements(IReadOnlyList<string> elements, double fontSize) {
        if (elements == null) throw new ArgumentNullException(nameof(elements));
        ValidateSize(fontSize);
        var widths = new double[elements.Count];
        for (int index = 0; index < elements.Count; index++) {
            widths[index] = Measure(elements[index], fontSize);
        }
        return widths;
    }

    public double LineHeight(double fontSize) {
        ValidateSize(fontSize);
        return LineSpacingRatio * fontSize;
    }

    public List<List<OfficePoint>> GetTextContours(
        string text,
        double x,
        double y,
        double fontSize) {
        return GetTextContoursBounded(
            text,
            x,
            y,
            fontSize,
            int.MaxValue,
            CancellationToken.None);
    }

    public List<List<OfficePoint>> GetTextContoursBounded(
        string text,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken) {
        ValidateTextAndSize(text, fontSize);
        if (maximumPointCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPointCount));
        cancellationToken.ThrowIfCancellationRequested();
        var renderer = new OfficeSixLaborsContourRenderer(maximumPointCount, cancellationToken);
        global::SixLabors.Fonts.Rendering.TextRenderer.RenderTo(
            renderer,
            text,
            CreateOptions(fontSize, x, y));
        cancellationToken.ThrowIfCancellationRequested();
        return renderer.GetContours();
    }

    public bool TryGetGlyphMetrics(int scalar, out int glyphId, out int advanceWidth) {
        Font font = GetFont(1F);
        if (!font.TryGetGlyphId(new CodePoint(scalar), out ushort mapped)
            || mapped == 0
            || !_metrics.TryGetGlyphMetrics(
                new CodePoint(scalar),
                TextAttributes.None,
                TextDecorations.None,
                LayoutMode.HorizontalTopBottom,
                ColorFontSupport.None,
                null,
                out FontGlyphMetrics? metrics)
            || metrics == null) {
            glyphId = 0;
            advanceWidth = 0;
            return false;
        }
        glyphId = mapped;
        advanceWidth = metrics.AdvanceWidth;
        return true;
    }

    public double MeasureShapedText(
        string text,
        OfficeTextShapingResult result,
        double fontSize) => Measure(text, fontSize);

    public List<List<OfficePoint>> GetShapedTextContours(
        string text,
        OfficeTextShapingResult result,
        double x,
        double y,
        double fontSize) => GetTextContours(text, x, y, fontSize);

    public List<List<OfficePoint>> GetShapedTextContoursBounded(
        string text,
        OfficeTextShapingResult result,
        double x,
        double y,
        double fontSize,
        int maximumPointCount,
        CancellationToken cancellationToken) => GetTextContoursBounded(
            text,
            x,
            y,
            fontSize,
            maximumPointCount,
            cancellationToken);

    private TextOptions CreateOptions(double fontSize, double x, double y) => new(GetFont(ToSingle(fontSize))) {
        Dpi = 72F,
        Origin = new System.Numerics.Vector2(ToSingle(x), ToSingle(y)),
        WrappingLength = -1F,
        KerningMode = KerningMode.Standard,
        TextDirection = TextDirection.Auto,
        ColorFontSupport = ColorFontSupport.None
    };

    private Font GetFont(float size) {
        lock (_fontSync) {
            if (_fontsBySize.TryGetValue(size, out Font? font)) return font;
            if (_fontsBySize.Count >= MaximumCachedSizes) _fontsBySize.Clear();
            font = _family.CreateFont(size, _style, _variations);
            _fontsBySize.Add(size, font);
            return font;
        }
    }

    private static void ValidateTextAndSize(string text, double fontSize) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        ValidateSize(fontSize);
    }

    private static void ValidateSize(double fontSize) {
        if (fontSize <= 0D || double.IsNaN(fontSize) || double.IsInfinity(fontSize) || fontSize > float.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(fontSize));
        }
    }

    private static float ToSingle(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < -float.MaxValue || value > float.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(value));
        }
        return (float)value;
    }

    private static int ReadScalar(string text, ref int index) {
        char first = text[index++];
        return char.IsHighSurrogate(first)
               && index < text.Length
               && char.IsLowSurrogate(text[index])
            ? char.ConvertToUtf32(first, text[index++])
            : first;
    }
}
