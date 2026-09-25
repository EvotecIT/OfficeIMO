using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Restores Arabic joining context for glyph runs that a PDF paints one letter at a time.
/// </summary>
/// <remarks>
/// PDF producers commonly paint shaped Arabic-script text as separate positioned glyphs, and ToUnicode
/// maps each contextual glyph back to its base letter. A single-letter run carries no joining context,
/// so a renderer can only draw the isolated form. This pass orders painted letters along their
/// baseline from right to left, splits words at gaps, and gives each base letter the presentation
/// form its joining context selects. Producer order and font or size changes inside a word are
/// irrelevant because only geometry is used. Visual projection only; logical text keeps the letters.
/// </remarks>
internal static class PdfArabicPaintedForms {
    private const double LineTolerance = 0.3D;
    private const double RotationTolerance = 0.1D;
    // Joined letters touch; an inter-word space is well above this fraction of the font size.
    private const double WordGap = 0.12D;

    internal static void Apply(List<PdfTextSpan> spans, System.Threading.CancellationToken cancellationToken = default) {
        // A word may change font or size mid-word, so only the baseline direction separates groups.
        var groups = new List<(double Rotation, List<int> Members)>();
        var rotationBuckets = new Dictionary<int, List<int>>();
        const int bucketCount = 3600;
        for (int index = 0; index < spans.Count; index++) {
            PdfTextSpan span = spans[index];
            if (span.IsVisible && span.Color?.A != 0 && span.Text.Length > 1 &&
                span.Text.All(IsPresentationForm)) span.MarkPaintedGlyphProjection();
            // Invisible text, such as an OCR layer, must not interleave with painted letters.
            if (!span.IsVisible || span.Color?.A == 0 || span.Text.Length != 1 || !IsArabicLetterOrForm(span.Text[0]) ||
                !(EffectiveFontSize(span) > 0D) || !(span.Advance > 0D) || double.IsInfinity(span.Advance) ||
                double.IsNaN(span.RotationDegrees) || double.IsInfinity(span.RotationDegrees)) continue;
            double normalizedRotation = (span.RotationDegrees % 360D + 360D) % 360D;
            int bucket = Math.Min(bucketCount - 1, (int)(normalizedRotation / RotationTolerance));
            int groupIndex = -1;
            for (int offset = -1; offset <= 1 && groupIndex < 0; offset++) {
                int neighbor = (bucket + offset + bucketCount) % bucketCount;
                if (!rotationBuckets.TryGetValue(neighbor, out List<int>? candidates)) continue;
                int match = candidates.FindIndex(candidate =>
                    Math.Abs(Math.IEEERemainder(span.RotationDegrees - groups[candidate].Rotation, 360D)) <= RotationTolerance);
                if (match >= 0) groupIndex = candidates[match];
            }
            if (groupIndex >= 0) {
                groups[groupIndex].Members.Add(index);
            } else {
                groupIndex = groups.Count;
                groups.Add((span.RotationDegrees, new List<int> { index }));
                if (!rotationBuckets.TryGetValue(bucket, out List<int>? candidates))
                    rotationBuckets.Add(bucket, candidates = new List<int>());
                candidates.Add(groupIndex);
            }
        }

        foreach ((double rotation, List<int> members) in groups) {
            cancellationToken.ThrowIfCancellationRequested();
            double radians = rotation * Math.PI / 180D;
            double cos = Math.Cos(radians);
            double sin = Math.Sin(radians);
            var letters = members
                .Select(index => (Index: index,
                    Along: spans[index].X * cos + spans[index].Y * sin,
                    Across: -spans[index].X * sin + spans[index].Y * cos,
                    spans[index].Advance,
                    Size: EffectiveFontSize(spans[index])))
                .OrderBy(letter => letter.Across)
                .ToList();

            int lineStart = 0;
            for (int index = 1; index <= letters.Count; index++) {
                if (index < letters.Count && letters[index].Across - letters[lineStart].Across <=
                    Math.Min(letters[index].Size, letters[lineStart].Size) * LineTolerance) continue;
                var line = letters.GetRange(lineStart, index - lineStart);
                line.Sort(static (left, right) => right.Along.CompareTo(left.Along));
                // Shadow, outline, and faux-bold passes can paint the same glyph more than once.
                // Shape one logical letter per position, then apply its form to every paint pass.
                var positions = new List<List<int>>();
                var representative = new List<(double Along, double Advance, double Size)>();
                foreach (var letter in line) {
                    if (positions.Count > 0) {
                        int last = positions.Count - 1;
                        var prior = representative[last];
                        if (Math.Abs(letter.Along - prior.Along) <= Math.Min(letter.Size, prior.Size) * 0.15D &&
                            Math.Abs(letter.Advance - prior.Advance) <= Math.Max(letter.Advance, prior.Advance) * 0.1D &&
                            spans[letter.Index].Text == spans[positions[last][0]].Text) {
                            positions[last].Add(letter.Index);
                            continue;
                        }
                    }
                    positions.Add(new List<int> { letter.Index });
                    representative.Add((letter.Along, letter.Advance, letter.Size));
                }
                int wordStart = 0;
                for (int letter = 1; letter <= positions.Count; letter++) {
                    if (letter < positions.Count) {
                        double size = Math.Max(representative[letter - 1].Size, representative[letter].Size);
                        double gap = representative[letter - 1].Along - (representative[letter].Along + representative[letter].Advance);
                        if (gap <= size * WordGap && gap >= -size) continue;
                    }
                    ApplyWord(spans, positions.GetRange(wordStart, letter - wordStart));
                    wordStart = letter;
                }
                lineStart = index;
            }
        }
    }

    private static void ApplyWord(List<PdfTextSpan> spans, List<List<int>> word) {
        var logical = new char[word.Count];
        for (int index = 0; index < word.Count; index++) {
            char painted = spans[word[index][0]].Text[0];
            // An isolated lam-alef ligature joins only to the preceding letter, like alef.
            logical[index] = IsIsolatedLamAlef(painted) ? '\u0627' : OfficeArabicTextShaper.ToLogicalText(painted.ToString())[0];
        }
        string shaped = OfficeArabicTextShaper.Shape(new string(logical));
        if (shaped.Length != word.Count) return;

        for (int index = 0; index < word.Count; index++) {
            foreach (int paintIndex in word[index]) {
                PdfTextSpan span = spans[paintIndex];
                char painted = span.Text[0];
                char form = painted;
                if (IsBaseLetter(painted)) form = shaped[index];
                else if (IsIsolatedLamAlef(painted) && shaped[index] == '\uFE8E') form = (char)(painted + 1);
                if (form != painted) {
                    PdfTextSpan visual = span.WithVisualText(form.ToString());
                    visual.MarkPaintedGlyphProjection();
                    spans[paintIndex] = visual;
                } else if (IsPresentationForm(painted)) span.MarkPaintedGlyphProjection();
            }
        }
    }

    private static double EffectiveFontSize(PdfTextSpan span) =>
        span.RestampFontSize > 0D && !double.IsNaN(span.RestampFontSize) && !double.IsInfinity(span.RestampFontSize)
            ? span.RestampFontSize : span.FontSize;

    private static bool IsArabicLetterOrForm(char value) =>
        IsBaseLetter(value) || value >= '\uFB50' && value <= '\uFBFF' || value >= '\uFE80' && value <= '\uFEFC';

    // Core letters and the extended letters the shaper covers (Persian, Urdu and related scripts).
    private static bool IsBaseLetter(char value) =>
        value >= '\u0621' && value <= '\u064A' || value >= '\u0671' && value <= '\u06D3';

    private static bool IsIsolatedLamAlef(char value) =>
        value == '\uFEF5' || value == '\uFEF7' || value == '\uFEF9' || value == '\uFEFB';

    internal static bool IsPresentationForm(char value) =>
        value >= '\uFB50' && value <= '\uFDFF' || value >= '\uFE70' && value <= '\uFEFC';
}
