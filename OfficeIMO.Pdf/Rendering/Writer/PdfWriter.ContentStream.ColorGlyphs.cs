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

    // The layout context that owns the page being written places PNG glyphs as page images and returns
    // the content-stream token that draws them. Outside a layout (watermark and appearance Forms) there is
    // no sink, and the preflight reports PNG glyphs instead of leaving a silent blank.
    [ThreadStatic] private static Func<OfficeBitmapGlyphImage, double, double, double, double, string?>? _bitmapGlyphSink;

    [ThreadStatic] private static int _bitmapGlyphPlacementSuppressed;

    /// <summary>Whether text being checked will be written where PNG glyphs cannot become page images.</summary>
    internal static bool BitmapGlyphPlacementSuppressed => _bitmapGlyphPlacementSuppressed > 0;

    /// <summary>Marks the text checked in this scope as written outside the page, such as a watermark Form.</summary>
    internal static IDisposable WithoutBitmapGlyphPlacement() {
        _bitmapGlyphPlacementSuppressed++;
        return new PlacementScope();
    }

    private sealed class PlacementScope : IDisposable {
        private bool _disposed;
        public void Dispose() {
            if (_disposed) return;
            _disposed = true;
            _bitmapGlyphPlacementSuppressed--;
        }
    }

    /// <summary>Installs the PNG glyph sink for this thread and returns the previous one, for the caller to restore.</summary>
    internal static Func<OfficeBitmapGlyphImage, double, double, double, double, string?>? ExchangeBitmapGlyphSink(
        Func<OfficeBitmapGlyphImage, double, double, double, double, string?>? sink) {
        var previous = _bitmapGlyphSink;
        _bitmapGlyphSink = sink;
        return previous;
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
        for (int index = 0; index < arts.Count; index++)
            // Page images are placed upright: a rotated run keeps its ordinary glyphs.
            if (arts[index]?.Bitmap != null && (_bitmapGlyphSink == null || T._textB != 0D || T._textA <= 0D || T._textD <= 0D)) return false;
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
                double x0 = item.A * bitmap.Left + item.C * bitmap.Bottom + item.E;
                double y0 = item.B * bitmap.Left + item.D * bitmap.Bottom + item.F;
                // Upright placement at the glyph origin; synthetic oblique shears only the outline glyphs.
                if (item.B == 0D && item.A > 0D && item.D > 0D) {
                    string? token = _bitmapGlyphSink?.Invoke(bitmap, x0, y0, item.A * bitmap.Width, item.D * bitmap.Height);
                    if (token != null) _sb.Append(token);
                }
                continue;
            }
            // Font units scale by size/unitsPerEm (about 0.005 to 0.05); keep that factor exact.
            _sb.Append("q\n").Append(PdfNumberFormatter.Precise(item.A)).Append(' ').Append(PdfNumberFormatter.Precise(item.B)).Append(' ')
                .Append(PdfNumberFormatter.Precise(item.C)).Append(' ').Append(PdfNumberFormatter.Precise(item.D)).Append(' ')
                .Append(PdfNumberFormatter.Precise(item.E)).Append(' ').Append(PdfNumberFormatter.Precise(item.F)).Append(" cm\n");
            foreach (PdfColorGlyphLayer layer in item.Art.Layers) {
                OfficeColor color = layer.Foreground ? item.Fill : layer.Color;
                // Translucent palette entries are rare in emoji fonts; paint them opaque rather than drop them.
                if (color.A == 0 || layer.Path.Length == 0) continue;
                _sb.Append(F(color.R / 255D)).Append(' ').Append(F(color.G / 255D)).Append(' ').Append(F(color.B / 255D)).Append(" rg\n")
                    .Append(layer.Path).Append("f\n");
            }
            _sb.Append("Q\n");
        }
        _deferredGlyphArt.Clear();
    }
}
