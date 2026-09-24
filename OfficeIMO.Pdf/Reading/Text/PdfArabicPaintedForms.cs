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
    // Joined letters touch; an inter-word space is well above this fraction of the font size.
    private const double WordGap = 0.12D;

    internal static void Apply(List<PdfTextSpan> spans, System.Threading.CancellationToken cancellationToken = default) {
        // A word may change font or size mid-word, so only the baseline direction separates groups.
        var groups = new Dictionary<double, List<int>>();
        for (int index = 0; index < spans.Count; index++) {
            PdfTextSpan span = spans[index];
            // Invisible text, such as an OCR layer, must not interleave with painted letters.
            if (!span.IsVisible || span.Text.Length != 1 || !IsArabicLetterOrForm(span.Text[0]) ||
                !(span.FontSize > 0D) || !(span.Advance > 0D) || double.IsInfinity(span.Advance)) continue;
            double rotation = Math.Round(span.RotationDegrees);
            if (!groups.TryGetValue(rotation, out List<int>? members)) groups.Add(rotation, members = new List<int>());
            members.Add(index);
        }

        foreach (KeyValuePair<double, List<int>> group in groups) {
            cancellationToken.ThrowIfCancellationRequested();
            double radians = group.Key * Math.PI / 180D;
            double cos = Math.Cos(radians);
            double sin = Math.Sin(radians);
            var letters = group.Value
                .Select(index => (Index: index,
                    Along: spans[index].X * cos + spans[index].Y * sin,
                    Across: -spans[index].X * sin + spans[index].Y * cos,
                    spans[index].Advance,
                    Size: spans[index].FontSize))
                .OrderBy(letter => letter.Across)
                .ToList();

            int lineStart = 0;
            for (int index = 1; index <= letters.Count; index++) {
                if (index < letters.Count && letters[index].Across - letters[lineStart].Across <=
                    Math.Min(letters[index].Size, letters[lineStart].Size) * LineTolerance) continue;
                var line = letters.GetRange(lineStart, index - lineStart);
                line.Sort(static (left, right) => right.Along.CompareTo(left.Along));
                int wordStart = 0;
                for (int letter = 1; letter <= line.Count; letter++) {
                    if (letter < line.Count) {
                        double size = Math.Max(line[letter - 1].Size, line[letter].Size);
                        double gap = line[letter - 1].Along - (line[letter].Along + line[letter].Advance);
                        if (gap <= size * WordGap && gap >= -size) continue;
                    }
                    ApplyWord(spans, line.GetRange(wordStart, letter - wordStart).Select(static item => item.Index).ToList());
                    wordStart = letter;
                }
                lineStart = index;
            }
        }
    }

    private static void ApplyWord(List<PdfTextSpan> spans, List<int> word) {
        var logical = new char[word.Count];
        for (int index = 0; index < word.Count; index++) {
            char painted = spans[word[index]].Text[0];
            // An isolated lam-alef ligature joins only to the preceding letter, like alef.
            logical[index] = IsIsolatedLamAlef(painted) ? '\u0627' : OfficeArabicTextShaper.ToLogicalText(painted.ToString())[0];
        }
        string shaped = OfficeArabicTextShaper.Shape(new string(logical));
        if (shaped.Length != word.Count) return;

        for (int index = 0; index < word.Count; index++) {
            PdfTextSpan span = spans[word[index]];
            char painted = span.Text[0];
            char form = painted;
            if (IsBaseLetter(painted)) form = shaped[index];
            else if (IsIsolatedLamAlef(painted) && shaped[index] == '\uFE8E') form = (char)(painted + 1);
            if (form != painted) spans[word[index]] = span.WithVisualText(form.ToString());
        }
    }

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
