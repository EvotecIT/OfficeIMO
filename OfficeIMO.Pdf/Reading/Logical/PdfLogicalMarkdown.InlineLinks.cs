namespace OfficeIMO.Pdf;

public static partial class PdfLogicalMarkdownExtensions {
    /// <summary>
    /// Places a page's web and mail links on the body text they cover. Each link annotation is placed at most once;
    /// links that cannot be placed, and links to other files, are reported once per target. Links to destinations
    /// inside the document and viewer actions are left out.
    /// </summary>
    private sealed class PdfLogicalMarkdownInlineLinks {
        // Shorter covered text (a footnote mark or a single letter) cannot be matched to one place in a paragraph.
        private const int MinimumMatchedCharacters = 3;

        private readonly List<PdfLogicalLinkAnnotation> _webLinks;
        private readonly bool[] _placed;
        private readonly List<PdfLogicalLinkAnnotation> _otherLinks;

        private PdfLogicalMarkdownInlineLinks(List<PdfLogicalLinkAnnotation> webLinks, List<PdfLogicalLinkAnnotation> otherLinks) {
            _webLinks = webLinks;
            _placed = new bool[webLinks.Count];
            _otherLinks = otherLinks;
        }

        internal readonly struct UnplacedLink {
            internal UnplacedLink(PdfLogicalLinkAnnotation annotation, string label, string target, bool isSafeUri) {
                Annotation = annotation;
                Label = label;
                Target = target;
                IsSafeUri = isSafeUri;
            }

            internal PdfLogicalLinkAnnotation Annotation { get; }
            internal string Label { get; }
            internal string Target { get; }
            internal bool IsSafeUri { get; }
        }

        internal static PdfLogicalMarkdownInlineLinks? Create(PdfLogicalPage page) {
            if (page.Links.Count == 0) return null;
            HashSet<PdfLogicalLinkAnnotation>? contentsLinks = null;
            foreach (PdfLogicalListItem item in page.ListItems) {
                if (item.IsContentsEntry && item.Link is not null) (contentsLinks ??= new HashSet<PdfLogicalLinkAnnotation>()).Add(item.Link);
            }

            var webLinks = new List<PdfLogicalLinkAnnotation>();
            var otherLinks = new List<PdfLogicalLinkAnnotation>();
            foreach (PdfLogicalLinkAnnotation link in page.Links) {
                // A contents entry already renders its own link.
                if (contentsLinks?.Contains(link) == true) continue;
                if (!string.IsNullOrEmpty(link.Uri)) {
                    if (IsSafeLinkUri(link.Uri!)) webLinks.Add(link);
                    else otherLinks.Add(link);
                } else if (!string.IsNullOrEmpty(link.RemoteFile)) {
                    otherLinks.Add(link);
                }
            }
            return webLinks.Count == 0 && otherLinks.Count == 0 ? null : new PdfLogicalMarkdownInlineLinks(webLinks, otherLinks);
        }

        /// <summary>Returns the ordered, non-overlapping ranges of <paramref name="text"/> covered by unplaced web links.</summary>
        internal List<(int Start, int End, string Target)>? Place(string text, IReadOnlyList<PdfLogicalTextBlock> lines) {
            if (string.IsNullOrEmpty(text) || lines.Count == 0) return null;
            List<(int Start, int End, string Target)>? ranges = null;
            string? compact = null;
            int[]? compactToText = null;
            for (int index = 0; index < _webLinks.Count; index++) {
                if (_placed[index]) continue;
                PdfLogicalLinkAnnotation link = _webLinks[index];
                string covered = GetCoveredText(lines, link);
                if (covered.Length < MinimumMatchedCharacters) continue;
                if (compact is null) BuildCompactText(text, out compact, out compactToText);
                // Search after earlier matches of the same text, so repeated words map to successive links.
                int searchFrom = 0;
                while (true) {
                    int found = compact.IndexOf(covered, searchFrom, StringComparison.Ordinal);
                    if (found < 0) break;
                    int start = compactToText![found];
                    int end = compactToText[found + covered.Length - 1] + 1;
                    // A rectangle that stops inside a word still links that word; the link text never splits it.
                    while (start > 0 && !char.IsWhiteSpace(text[start - 1])) start--;
                    while (end < text.Length && !char.IsWhiteSpace(text[end])) end++;
                    if (ranges is null || !ranges.Any(range => start < range.End && end > range.Start)) {
                        (ranges ??= new List<(int Start, int End, string Target)>()).Add((start, end, link.Uri!));
                        _placed[index] = true;
                        break;
                    }
                    searchFrom = found + 1;
                }
            }
            ranges?.Sort(static (left, right) => left.Start.CompareTo(right.Start));
            return ranges;
        }

        internal IEnumerable<UnplacedLink> GetUnplacedLinks() {
            var placedTargets = new HashSet<string>(StringComparer.Ordinal);
            for (int index = 0; index < _webLinks.Count; index++) {
                if (_placed[index]) placedTargets.Add(_webLinks[index].Uri!);
            }
            var listed = new HashSet<string>(StringComparer.Ordinal);
            for (int index = 0; index < _webLinks.Count; index++) {
                if (_placed[index]) continue;
                PdfLogicalLinkAnnotation link = _webLinks[index];
                string target = link.Uri!;
                if (placedTargets.Contains(target) || !listed.Add(target)) continue;
                yield return new UnplacedLink(link, GetLabel(link, target), target, isSafeUri: true);
            }
            foreach (PdfLogicalLinkAnnotation link in _otherLinks) {
                string? target = FormatLinkAnnotationTarget(link);
                if (string.IsNullOrEmpty(target) || !listed.Add(target!)) continue;
                yield return new UnplacedLink(link, GetLabel(link, target!), target!, isSafeUri: false);
            }
        }

        private static string GetLabel(PdfLogicalLinkAnnotation link, string target) =>
            !string.IsNullOrWhiteSpace(link.Contents) ? link.Contents! : target;

        // The covered text, without whitespace, of the unrotated runs whose glyph centres lie inside the link rectangle.
        private static string GetCoveredText(IReadOnlyList<PdfLogicalTextBlock> lines, PdfLogicalLinkAnnotation link) {
            double left = Math.Min(link.X1, link.X2), right = Math.Max(link.X1, link.X2);
            double bottom = Math.Min(link.Y1, link.Y2), top = Math.Max(link.Y1, link.Y2);
            if (!(right > left) || !(top > bottom)) return string.Empty;
            StringBuilder? builder = null;
            foreach (PdfLogicalTextBlock line in lines) {
                foreach (PdfTextSpan span in line.Spans) {
                    if (span.RotationDegrees != 0D || string.IsNullOrEmpty(span.Text)) continue;
                    // Vertical: the middle of the x-height sits inside the rectangle.
                    double middle = span.Y + span.FontSize * 0.3D;
                    if (middle < bottom - 1D || middle > top + 1D) continue;
                    double spanRight = span.X + Math.Max(0D, span.Advance);
                    if (spanRight < left || span.X > right) continue;
                    if (span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs) && glyphs.Count > 0) {
                        foreach (PdfTextGlyph glyph in glyphs) {
                            double center = glyph.X + glyph.Advance / 2D;
                            if (center >= left && center <= right) AppendCompact(ref builder, glyph.Text);
                        }
                    } else {
                        double width = Math.Max(0D, span.Advance) / span.Text.Length;
                        for (int index = 0; index < span.Text.Length; index++) {
                            double center = span.X + width * (index + 0.5D);
                            if (center >= left && center <= right) AppendCompact(ref builder, span.Text[index].ToString());
                        }
                    }
                }
            }
            return builder?.ToString() ?? string.Empty;
        }

        private static void AppendCompact(ref StringBuilder? builder, string text) {
            foreach (char ch in text) {
                if (!char.IsWhiteSpace(ch)) (builder ??= new StringBuilder()).Append(ch);
            }
        }

        // Paragraph text joins its lines with spaces, so matching ignores whitespace and maps back to the original text.
        private static void BuildCompactText(string text, out string compact, out int[] compactToText) {
            var builder = new StringBuilder(text.Length);
            var map = new List<int>(text.Length);
            for (int index = 0; index < text.Length; index++) {
                if (char.IsWhiteSpace(text[index])) continue;
                builder.Append(text[index]);
                map.Add(index);
            }
            compact = builder.ToString();
            compactToText = map.ToArray();
        }
    }
}
