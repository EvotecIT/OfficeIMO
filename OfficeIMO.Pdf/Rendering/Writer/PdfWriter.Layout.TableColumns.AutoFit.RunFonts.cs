using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>
    /// Measures auto-fit text in the faces that paint it: each run's named font,
    /// or the fallback runs the renderer plans for glyphs the row font lacks.
    /// </summary>
    private sealed class AutoFitRunFontMeasure {
        private readonly string _text;
        private readonly int[] _runStarts;
        private readonly PdfStandardFont[] _fonts;
        private readonly PdfNamedFontFace?[] _namedFonts;
        private readonly double[] _sizes;
        private readonly OfficeTextFeatureSettings?[] _features;
        private readonly PdfCreateOptions _options;
        private System.Collections.Generic.Dictionary<int, int>? _firstRunByScalar;

        private AutoFitRunFontMeasure(string text, int[] runStarts, PdfStandardFont[] fonts, PdfNamedFontFace?[] namedFonts,
            double[] sizes, OfficeTextFeatureSettings?[] features, PdfCreateOptions options) {
            _text = text;
            _runStarts = runStarts;
            _fonts = fonts;
            _namedFonts = namedFonts;
            _sizes = sizes;
            _features = features;
            _options = options;
        }

        /// <summary>
        /// Returns null when every run paints in the row's standard face, so the
        /// row-font measurement stays exact for ordinary cells.
        /// </summary>
        internal static AutoFitRunFontMeasure? TryCreate(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs,
            string cellText, PdfStandardFont rowFont, double rowSize, PdfCreateOptions options) {
            if (!NeedsRunFontMeasurement(runs, cellText)) return null;

            System.Collections.Generic.IReadOnlyList<PdfTextRun> normalized = NormalizeFallbackRuns(runs, rowFont, options);
            var builder = new System.Text.StringBuilder();
            var starts = new System.Collections.Generic.List<int>(normalized.Count);
            var fonts = new System.Collections.Generic.List<PdfStandardFont>(normalized.Count);
            var namedFonts = new System.Collections.Generic.List<PdfNamedFontFace?>(normalized.Count);
            var sizes = new System.Collections.Generic.List<double>(normalized.Count);
            var features = new System.Collections.Generic.List<OfficeTextFeatureSettings?>(normalized.Count);
            bool ownFace = false;
            for (int index = 0; index < normalized.Count; index++) {
                PdfTextRun run = normalized[index];
                if (run.InlineElement != null || string.IsNullOrEmpty(run.Text)) continue;
                PdfNamedFontFace? named = options.TryResolveNamedFontFace(run.FontFamily, run.Bold, run.Italic, out PdfNamedFontFace face)
                    ? face : null;
                PdfStandardFont font = ResolvePageTextRunFont(run, rowFont);
                ownFace |= named.HasValue || !IsSameRun(runs, run);
                starts.Add(builder.Length);
                builder.Append(run.Text);
                fonts.Add(font);
                namedFonts.Add(named);
                sizes.Add(run.FontSize?.ToPoints() ?? rowSize);
                features.Add(run.FeatureSettings);
            }

            return ownFace
                ? new AutoFitRunFontMeasure(builder.ToString(), starts.ToArray(), fonts.ToArray(), namedFonts.ToArray(),
                    sizes.ToArray(), features.ToArray(), options)
                : null;
        }

        internal double Measure(string? value) {
            if (string.IsNullOrEmpty(value)) return 0D;
            int start = _text.IndexOf(value!, StringComparison.Ordinal);
            return start >= 0 ? MeasureSlice(start, value!.Length) : MeasureByFirstOccurrence(value!);
        }

        private double MeasureSlice(int start, int length) {
            double width = 0D;
            int end = start + length;
            int run = FindRun(start);
            while (start < end) {
                int runEnd = run + 1 < _runStarts.Length ? _runStarts[run + 1] : _text.Length;
                int pieceEnd = Math.Min(end, runEnd);
                width += MeasurePiece(run, _text.Substring(start, pieceEnd - start));
                start = pieceEnd;
                run++;
            }
            return width;
        }

        // A token the renderer never sees verbatim, such as a trimmed or
        // re-split word, takes each scalar from the run that paints it.
        private double MeasureByFirstOccurrence(string value) {
            _firstRunByScalar ??= BuildFirstRunByScalar();
            double width = 0D;
            int pieceStart = 0;
            int pieceRun = -1;
            for (int index = 0; index < value.Length;) {
                int scalar = ScalarAt(value, index, out int length);
                int run = _firstRunByScalar.TryGetValue(scalar, out int found) ? found : 0;
                if (pieceRun >= 0 && run != pieceRun) {
                    width += MeasurePiece(pieceRun, value.Substring(pieceStart, index - pieceStart));
                    pieceStart = index;
                }
                pieceRun = run;
                index += length;
            }
            return pieceRun < 0 ? 0D : width + MeasurePiece(pieceRun, value.Substring(pieceStart));
        }

        private System.Collections.Generic.Dictionary<int, int> BuildFirstRunByScalar() {
            var map = new System.Collections.Generic.Dictionary<int, int>();
            int run = 0;
            for (int index = 0; index < _text.Length;) {
                while (run + 1 < _runStarts.Length && _runStarts[run + 1] <= index) run++;
                int scalar = ScalarAt(_text, index, out int length);
                if (!map.TryGetValue(scalar, out _)) map[scalar] = run;
                index += length;
            }
            return map;
        }

        private static int ScalarAt(string text, int index, out int length) {
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                length = 2;
                return char.ConvertToUtf32(text[index], text[index + 1]);
            }
            length = 1;
            return text[index];
        }

        private int FindRun(int offset) {
            int run = Array.BinarySearch(_runStarts, offset);
            return run >= 0 ? run : Math.Max(0, ~run - 1);
        }

        private double MeasurePiece(int run, string text) =>
            EstimateSimpleTextWidthForOptions(text, _fonts[run], _namedFonts[run], _sizes[run], _options, _features[run]);

        private static bool NeedsRunFontMeasurement(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, string cellText) {
            for (int index = 0; index < runs.Count; index++)
                if (runs[index].InlineElement == null && runs[index].FontFamily != null) return true;
            // Fallback runs replace only glyphs outside the row font's Latin repertoire.
            foreach (char ch in cellText)
                if (ch > '\u007f') return true;
            return false;
        }

        private static bool IsSameRun(System.Collections.Generic.IReadOnlyList<PdfTextRun> source, PdfTextRun run) {
            for (int index = 0; index < source.Count; index++) {
                PdfTextRun candidate = source[index];
                if (string.Equals(candidate.Text, run.Text, StringComparison.Ordinal) &&
                    candidate.StandardFont == run.StandardFont &&
                    string.Equals(candidate.FontFamily, run.FontFamily, StringComparison.Ordinal))
                    return true;
            }
            return false;
        }
    }
}
