using System.Globalization;
using System.Threading;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal sealed partial class VisioRichTextProjection {
    private static VisioRichTextProjection? CreateParagraphProjection(VisioDocument? document,
        XElement text, IReadOnlyDictionary<string, XElement> characters, XElement section,
        VisioTextStyle? style, double scale, double defaultSize, OfficeTextAlignment alignment,
        CancellationToken cancellationToken) {
        XNamespace ns = text.Name.Namespace;
        var rows = new Dictionary<string, XElement>(StringComparer.Ordinal);
        int position = 0;
        foreach (XElement row in section.Elements(ns + "Row")) {
            cancellationToken.ThrowIfCancellationRequested();
            if (rows.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns) return null;
            rows[VisioNativeTextStyleResolver.RowIndex((string?)row.Attribute("IX"), position)] = row;
            position++;
        }
        if (rows.Count == 0) return null;
        var paragraphs = new List<OfficeRichTextParagraph>();
        var paragraphRuns = new List<OfficeRichTextRun>();
        var runs = new List<OfficeRichTextRun>();
        int projectedCharacters = 0;
        string characterIndex = "0", paragraphIndex = "0";
        XElement fallbackCharacter = new(ns + "Row");
        foreach (XNode node in text.DescendantNodes()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (node is XElement marker) {
                if (marker.Name == ns + "cp") characterIndex = VisioNativeTextStyleResolver.RowIndex((string?)marker.Attribute("IX"));
                if (marker.Name == ns + "pp") {
                    // Native paragraph property runs start at a paragraph boundary.
                    if (paragraphRuns.Count > 0) return null;
                    paragraphIndex = VisioNativeTextStyleResolver.RowIndex((string?)marker.Attribute("IX"));
                }
            } else if (node is XText value && value.Value.Length > 0) {
                string normalized = value.Value.Replace("\r\n", "\n").Replace('\r', '\n');
                int start = 0;
                while (start < normalized.Length) {
                    cancellationToken.ThrowIfCancellationRequested();
                    int end = normalized.IndexOf('\n', start);
                    int length = (end < 0 ? normalized.Length : end) - start;
                    if (length > 0 && !AddRun(normalized.Substring(start, length))) return null;
                    if (end < 0) break;
                    if (!FinishParagraph()) return null;
                    start = end + 1;
                }
            }
        }
        // A producer's terminal newline closes the last paragraph; it does not add one.
        if ((paragraphRuns.Count > 0 || paragraphs.Count == 0) && !FinishParagraph()) return null;
        return new VisioRichTextProjection(runs, alignment, paragraphs);

        bool AddRun(string value) {
            if (runs.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns) return false;
            XElement row = fallbackCharacter;
            if (characters.Count > 0 && !characters.TryGetValue(characterIndex, out row!)) return false;
            OfficeRichTextRun run = CreateRun(value, row, document, style, scale, defaultSize);
            if (run.Text.Length > OfficeTextLayoutEngine.MaximumLayoutTextCharacters - projectedCharacters) return false;
            projectedCharacters += run.Text.Length;
            runs.Add(run); paragraphRuns.Add(run);
            return true;
        }

        bool FinishParagraph() {
            if (paragraphs.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns || !rows.TryGetValue(paragraphIndex, out XElement? row)) return false;
            // Empty paragraphs retain their current character size without emitting a glyph.
            if (paragraphRuns.Count == 0 && !AddRun(string.Empty)) return false;
            OfficeRichTextParagraph? paragraph = CreateParagraph(paragraphRuns, row, document, scale, alignment,
                OfficeTextLayoutEngine.MaximumLayoutTextCharacters - projectedCharacters);
            if (paragraph == null) return false;
            if (paragraph.Label != null) {
                if (runs.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns) return false;
                runs.Add(paragraph.Label.Run);
                projectedCharacters += paragraph.Label.Run.Text.Length;
            }
            paragraphs.Add(paragraph); paragraphRuns.Clear();
            return true;
        }
    }

    private static OfficeRichTextParagraph? CreateParagraph(IReadOnlyList<OfficeRichTextRun> runs,
        XElement row, VisioDocument? document, double scale, OfficeTextAlignment fallback, int remainingCharacters) {
        OfficeTextAlignment alignment = fallback;
        if (TryInt(Cell(row, "HorzAlign"), out int align) && align >= 0 && align <= 3)
            alignment = VisioDrawingTextAlignment.ToOfficeTextAlignment((VisioTextHorizontalAlignment)align);
        double first = ParagraphLength(row, "IndFirst", scale);
        double left = ParagraphLength(row, "IndLeft", scale);
        double right = ParagraphLength(row, "IndRight", scale);
        double before = ParagraphLength(row, "SpBefore", scale);
        double after = ParagraphLength(row, "SpAfter", scale);
        // Signed first-line indents become a shared left inset and nonnegative offsets.
        // Hanging text outside the native frame and overlapping paragraph margins
        // remain preserved native content, outside this layout profile.
        double inset = left + Math.Min(0D, first);
        if (!Finite(first) || !Finite(inset) || !Finite(right) || !Finite(before) || !Finite(after)
            || inset < 0D || right < 0D || before < 0D || after < 0D) return null;
        double? height = null, factor = null;
        if (TryParagraphNumber(row, "SpLine", out double spacing)) {
            if (spacing > 0D) height = spacing * scale;
            else factor = spacing == 0D ? 1D : -spacing;
        }
        if (height.HasValue && (!Finite(height.Value) || height.Value <= 0D)) return null;
        var margins = new OfficeTextPadding(inset, before, right, after);
        var indent = new OfficeTextParagraphIndent(Math.Max(0D, first), Math.Max(0D, -first));
        if (!TryParagraphLabel(row, runs[0], document, scale, left + first, remainingCharacters, out OfficeTextParagraphLabel? label)) return null;
        if (label == null) return new OfficeRichTextParagraph(runs, alignment, height, margins, indent, factor);
        if (runs.Count >= OfficeTextLayoutEngine.MaximumLayoutTextRuns) return null;
        return new OfficeRichTextParagraph(runs, label, alignment, height, margins, indent, factor);
    }

    private static double ParagraphLength(XElement row, string name, double scale) =>
        TryParagraphNumber(row, name, out double value) ? value * scale : 0D;

    private static bool TryParagraphNumber(XElement row, string name, out double value) =>
        double.TryParse(Cell(row, name), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
        && !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
