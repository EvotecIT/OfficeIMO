using System;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

// Repeated glyph paint is shared by drawing and ordinary PDF tab layout.
internal static class OfficeTextTabLeaderLayout {
    internal static Paint Create(string glyph, double gap, int maximumCharacters, Func<string, double> measure,
        CancellationToken cancellationToken, int minimumGapGlyphs = 0) {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrWhiteSpace(glyph) || gap <= 0) return default;
        if (double.IsNaN(gap) || double.IsInfinity(gap)) return new Paint(string.Empty, 0, true);
        double glyphWidth = measure(glyph);
        if (glyphWidth <= 0 || double.IsNaN(glyphWidth) || double.IsInfinity(glyphWidth)) return new Paint(string.Empty, 0, true);
        if (gap <= glyphWidth * minimumGapGlyphs) return default;
        double requested = Math.Floor(gap / glyphWidth);
        int maximum = Math.Max(0, maximumCharacters / glyph.Length);
        bool limited = requested > maximum;
        int count = requested >= maximum ? maximum : (int)requested;
        if (count == 0) return new Paint(string.Empty, 0, limited);
        string text = Repeat(count); double width = Measure(text);
        // Whole-run shaping can differ from isolated glyph metrics. Reduce paint
        // rather than letting it cover the following field; searches are bounded.
        if (width < 0 || width > gap || double.IsNaN(width) || double.IsInfinity(width)) {
            int low = 0, high = count - 1;
            while (low < high) {
                int middle = low + (high - low + 1) / 2;
                double candidate = Measure(Repeat(middle));
                if (candidate >= 0 && candidate <= gap) low = middle; else high = middle - 1;
            }
            text = Repeat(low); width = text.Length == 0 ? 0 : Measure(text);
        }
        if (width < 0 || width > gap || double.IsNaN(width) || double.IsInfinity(width)) return new Paint(string.Empty, 0, true);
        return new Paint(text, width, limited);

        string Repeat(int repetitions) {
            cancellationToken.ThrowIfCancellationRequested();
            var result = new StringBuilder(repetitions * glyph.Length);
            for (int i = 0; i < repetitions; i++) {
                if ((i & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                result.Append(glyph);
            }
            return result.ToString();
        }
        double Measure(string value) { cancellationToken.ThrowIfCancellationRequested(); return measure(value); }
    }

    // One generated-paint budget spans every paragraph in a drawing text frame.
    internal sealed class Budget {
        internal Budget(bool reportClipping = true) { ReportClipping = reportClipping; }
        internal bool ReportClipping { get; }
        internal int Remaining = OfficeTextLayoutEngine.MaximumLayoutTextCharacters;
        internal int RemainingVertices = OfficeTextTabLineLeaderLayout.MaximumFrameVertices;
    }

    internal readonly struct Paint {
        internal Paint(string text, double width, bool limited) { Text = text; Width = width; Limited = limited; }
        internal string? Text { get; }
        internal double Width { get; }
        internal bool Limited { get; }
    }
}
