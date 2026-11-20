using System.Globalization;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>One filled outline of a color glyph, as PDF path operators in font units.</summary>
internal readonly struct PdfColorGlyphLayer {
    internal PdfColorGlyphLayer(string path, OfficeColor color, bool foreground) {
        Path = path; Color = color; Foreground = foreground;
    }

    internal string Path { get; }
    internal OfficeColor Color { get; }
    /// <summary>Whether the layer paints with the text's own fill color.</summary>
    internal bool Foreground { get; }
}

/// <summary>Painted artwork of one glyph in a run that contains color glyphs: COLR layers, a PNG bitmap or the plain outline.</summary>
internal sealed class PdfColorGlyphArt {
    internal PdfColorGlyphArt(int unitsPerEm, IReadOnlyList<PdfColorGlyphLayer> layers, OfficeBitmapGlyphImage? bitmap = null) {
        UnitsPerEm = unitsPerEm; Layers = layers; Bitmap = bitmap;
    }

    internal int UnitsPerEm { get; }
    internal IReadOnlyList<PdfColorGlyphLayer> Layers { get; }
    internal OfficeBitmapGlyphImage? Bitmap { get; }

    internal static string ToPathOperators(List<List<OfficePoint>> contours) {
        var path = new StringBuilder();
        foreach (List<OfficePoint> contour in contours) {
            if (contour.Count == 0) continue;
            for (int index = 0; index < contour.Count; index++) {
                path.Append(Number(contour[index].X)).Append(' ').Append(Number(contour[index].Y))
                    .Append(index == 0 ? " m\n" : " l\n");
            }
            path.Append("h\n");
        }
        return path.ToString();
    }

    private static string Number(double value) => Math.Round(value, 2).ToString("0.##", CultureInfo.InvariantCulture);
}

/// <summary>The page that content is written for: it gives PNG color glyphs and translucent layers their resources.</summary>
internal interface IColorGlyphPage {
    /// <summary>Registers a PNG glyph drawn through the given image unit-square transform and returns the token that draws it.</summary>
    string? PlaceImage(OfficeBitmapGlyphImage image, OfficeTransform unitSquare);

    /// <summary>Returns the name of an ExtGState with this fill and stroke alpha, shared by every use of that alpha.</summary>
    string? AlphaState(double alpha);
}

internal sealed partial class ContentStreamBuilder {
    private readonly struct DeferredGlyphArt {
        internal DeferredGlyphArt(double a, double b, double c, double d, double e, double f, PdfColorGlyphArt art, OfficeColor fill) {
            A = a; B = b; C = c; D = d; E = e; F = f; Art = art; Fill = fill;
        }
        internal double A { get; }
        internal double B { get; }
        internal double C { get; }
        internal double D { get; }
        internal double E { get; }
        internal double F { get; }
        internal PdfColorGlyphArt Art { get; }
        internal OfficeColor Fill { get; }
    }

    // The page being written. The layout context and page assembly install it around page content.
    // Without one (annotation appearances), PNG glyphs stay ordinary text and translucent layers paint opaque.
    [ThreadStatic] private static IColorGlyphPage? _colorGlyphPage;
    // 1 minus the opacity the enclosing content already applies, such as a translucent watermark's,
    // stored as the complement so the thread-static default means fully opaque.
    [ThreadStatic] private static double _colorGlyphAlphaComplement;

    /// <summary>Installs the page for this thread and returns the previous one, for the caller to restore.</summary>
    internal static IColorGlyphPage? ExchangeColorGlyphPage(IColorGlyphPage? page) {
        IColorGlyphPage? previous = _colorGlyphPage;
        _colorGlyphPage = page;
        return previous;
    }

    /// <summary>
    /// Scales translucent layer alpha by the opacity an enclosing graphics state already applies. An ExtGState
    /// <c>ca</c> replaces the current alpha rather than multiplying it, so a layer must carry the product.
    /// </summary>
    internal static IDisposable ScaleColorGlyphAlpha(double opacity) {
        double previous = _colorGlyphAlphaComplement;
        _colorGlyphAlphaComplement = 1D - Math.Max(0D, Math.Min(1D, opacity));
        return new AlphaScope(previous);
    }

    private sealed class AlphaScope : IDisposable {
        private readonly double _previous;
        internal AlphaScope(double previous) => _previous = previous;
        public void Dispose() => _colorGlyphAlphaComplement = _previous;
    }

    private List<DeferredGlyphArt>? _deferredGlyphArt;
    private int _textRenderingMode;

    // A run with color glyphs is shown invisibly, so extraction keeps its glyph codes, ToUnicode
    // and ActualText, and every glyph's artwork is painted after the text object at the same pen.
    private bool TryQueueColorGlyphArt(PdfTextShowCommand command, double fontSize, double textRise) {
        IReadOnlyList<PdfColorGlyphArt?>? arts = command.ColorArt;
        if (arts == null || !T._positionKnown || _textRenderingMode is 3 or 7
            || command.IsScalarBidi || command.IsolatedRightToLeft || command.SimpleFontKerning != null) return false;
        IReadOnlyList<PdfGlyphInfo>? glyphs = command.PositionedGlyphs ?? command.VisualGlyphs;
        if (glyphs == null || glyphs.Count != arts.Count || glyphs.Count != command.GlyphCount) return false;
        // A PNG glyph needs a page to become an image of; without one the run keeps its ordinary glyphs.
        for (int index = 0; index < arts.Count; index++)
            if (arts[index]?.Bitmap != null && _colorGlyphPage == null) return false;
        double scale = T._textScale, pen = 0D;
        // Tracking is written as a TJ adjustment after each boundary glyph, as AppendPositionedGlyphs does.
        double tracking = (command.Tracking?.GetAdjustment(fontSize / command.FontMetricScale) ?? 0D)
            * fontSize / command.UnitsPerEm * (command.NegativeTracking ? -1D : 1D);
        _deferredGlyphArt ??= new List<DeferredGlyphArt>();
        for (int index = 0; index < glyphs.Count; index++) {
            PdfGlyphInfo glyph = glyphs[index];
            PdfColorGlyphArt? art = arts[index];
            if (art != null) {
                double unit = fontSize / art.UnitsPerEm;
                double x = (pen + glyph.OffsetX1000 * fontSize / 1000D) * scale;
                double y = textRise + glyph.OffsetY1000 * fontSize / 1000D;
                // Glyph space to text space, then through the tracked text matrix.
                double ga = unit * scale, gd = unit;
                _deferredGlyphArt.Add(new DeferredGlyphArt(
                    ga * T._textA, ga * T._textB, gd * T._textC, gd * T._textD,
                    x * T._textA + y * T._textC + T._textE, x * T._textB + y * T._textD + T._textF,
                    art, _paintFillColor));
            }
            pen += glyph.AdvanceWidth1000 * fontSize / 1000D + T._textCharacterSpacing
                + command.GetJustificationAdvance(index) / scale
                + (command.TrackingBoundaries?[index] == true ? tracking : 0D);
        }
        return true;
    }

    private void FlushDeferredGlyphArt() {
        if (_deferredGlyphArt == null || _deferredGlyphArt.Count == 0) return;
        foreach (DeferredGlyphArt item in _deferredGlyphArt) {
            if (item.Art.Bitmap is { } bitmap) {
                // The image's unit square through its glyph box and the full glyph matrix: rotation, skew and scale.
                var unitSquare = new OfficeTransform(item.A * bitmap.Width, item.B * bitmap.Width, item.C * bitmap.Height, item.D * bitmap.Height,
                    item.A * bitmap.Left + item.C * bitmap.Bottom + item.E, item.B * bitmap.Left + item.D * bitmap.Bottom + item.F);
                string? token = _colorGlyphPage?.PlaceImage(bitmap, unitSquare);
                if (token != null) _sb.Append(token);
                continue;
            }
            // Font units scale by size/unitsPerEm (about 0.005 to 0.05); keep that factor exact.
            _sb.Append("q\n").Append(PdfNumberFormatter.Precise(item.A)).Append(' ').Append(PdfNumberFormatter.Precise(item.B)).Append(' ')
                .Append(PdfNumberFormatter.Precise(item.C)).Append(' ').Append(PdfNumberFormatter.Precise(item.D)).Append(' ')
                .Append(PdfNumberFormatter.Precise(item.E)).Append(' ').Append(PdfNumberFormatter.Precise(item.F)).Append(" cm\n");
            foreach (PdfColorGlyphLayer layer in item.Art.Layers) {
                OfficeColor color = layer.Foreground ? item.Fill : layer.Color;
                if (color.A == 0 || layer.Path.Length == 0) continue;
                string? alphaState = color.A < 255
                    ? _colorGlyphPage?.AlphaState(color.A / 255D * (1D - _colorGlyphAlphaComplement)) : null;
                if (alphaState != null) _sb.Append("q\n/").Append(alphaState).Append(" gs\n");
                _sb.Append(F(color.R / 255D)).Append(' ').Append(F(color.G / 255D)).Append(' ').Append(F(color.B / 255D)).Append(" rg\n")
                    .Append(layer.Path).Append("f\n");
                if (alphaState != null) _sb.Append("Q\n");
            }
            _sb.Append("Q\n");
        }
        _deferredGlyphArt.Clear();
    }
}
