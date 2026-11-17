namespace OfficeIMO.Pdf;

public static partial class PdfLogicalMarkdownExtensions {
    /// <summary>
    /// Places a page's web and mail links on the body text they cover. Each link annotation is placed at most once;
    /// links that cannot be placed, and links to other files, are reported once per target. Links to destinations
    /// inside the document and viewer actions are left out.
    /// </summary>
    private sealed class PdfLogicalMarkdownInlineLinks {
        // A run may be found this many characters after the previous one; paragraph text can drop a hyphen or mark
        // between runs, but a larger jump means the run is not where the paragraph text is being read.
        private const int MaximumAlignmentSkip = 2;

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

        // A source run and the index in the rendered text of each of its characters (-1 for whitespace).
        private readonly struct AlignedSpan {
            internal AlignedSpan(PdfTextSpan span, int[] textIndexes) {
                Span = span;
                TextIndexes = textIndexes;
            }

            internal PdfTextSpan Span { get; }
            internal int[] TextIndexes { get; }
        }

        internal static PdfLogicalMarkdownInlineLinks? Create(PdfLogicalPage page) {
            if (page.Links.Count == 0) return null;
            HashSet<PdfLogicalLinkAnnotation>? contentsLinks = null;
            foreach (PdfLogicalListItem item in page.ListItems) {
                // Only a contents entry that writes its link as the title's target already carries it.
                if (item.IsContentsEntry && item.Link is not null && FormatContentsLinkTarget(item.Link) is not null) {
                    (contentsLinks ??= new HashSet<PdfLogicalLinkAnnotation>()).Add(item.Link);
                }
            }

            var webLinks = new List<PdfLogicalLinkAnnotation>();
            var otherLinks = new List<PdfLogicalLinkAnnotation>();
            foreach (PdfLogicalLinkAnnotation link in page.Links) {
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

        /// <summary>
        /// Returns the ordered, non-overlapping ranges of <paramref name="text"/> painted under unplaced web links. Each
        /// source run of <paramref name="lines"/> is aligned to its own place in the text, so a link covers the characters
        /// under its rectangle even when the same words occur elsewhere in the paragraph.
        /// </summary>
        internal List<(int Start, int End, string Target)>? Place(string text, IReadOnlyList<PdfLogicalTextBlock> lines) {
            if (string.IsNullOrEmpty(text) || lines.Count == 0) return null;
            List<(int Start, int End, string Target)>? ranges = null;
            List<AlignedSpan>? aligned = null;
            for (int index = 0; index < _webLinks.Count; index++) {
                if (_placed[index]) continue;
                PdfLogicalLinkAnnotation link = _webLinks[index];
                if (!TryGetRectangle(link, out double left, out double bottom, out double right, out double top) ||
                    !TouchesAnyRun(lines, left, bottom, right, top)) continue;
                aligned ??= AlignRuns(text, lines);
                int start = int.MaxValue, end = -1;
                foreach (AlignedSpan run in aligned) {
                    if (!IsCandidateRun(run.Span, left, bottom, right, top)) continue;
                    foreach (int character in GetCoveredCharacters(run.Span, left, right)) {
                        int textIndex = run.TextIndexes[character];
                        if (textIndex < 0) continue;
                        start = Math.Min(start, textIndex);
                        end = Math.Max(end, textIndex + 1);
                    }
                }
                if (end < 0) continue;
                // A rectangle that stops inside a word still links that word; the link text never splits it.
                // Only letters and digits extend it, so neighbouring punctuation stays outside the link.
                while (start > 0 && char.IsLetterOrDigit(text[start - 1]) && char.IsLetterOrDigit(text[start])) start--;
                while (end < text.Length && char.IsLetterOrDigit(text[end]) && char.IsLetterOrDigit(text[end - 1])) end++;
                if (ranges is not null && ranges.Any(range => start < range.End && end > range.Start)) continue;
                (ranges ??= new List<(int Start, int End, string Target)>()).Add((start, end, link.Uri!));
                _placed[index] = true;
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

        private static bool TryGetRectangle(PdfLogicalLinkAnnotation link, out double left, out double bottom, out double right, out double top) {
            left = Math.Min(link.X1, link.X2);
            right = Math.Max(link.X1, link.X2);
            bottom = Math.Min(link.Y1, link.Y2);
            top = Math.Max(link.Y1, link.Y2);
            return right > left && top > bottom;
        }

        private static bool TouchesAnyRun(IReadOnlyList<PdfLogicalTextBlock> lines, double left, double bottom, double right, double top) {
            foreach (PdfLogicalTextBlock line in lines) {
                foreach (PdfTextSpan span in line.Spans) {
                    if (IsCandidateRun(span, left, bottom, right, top)) return true;
                }
            }
            return false;
        }

        // An unrotated run whose x-height middle lies inside the rectangle and whose extent reaches it horizontally.
        private static bool IsCandidateRun(PdfTextSpan span, double left, double bottom, double right, double top) {
            if (span.RotationDegrees != 0D || string.IsNullOrEmpty(span.Text)) return false;
            double middle = span.Y + span.FontSize * 0.3D;
            if (middle < bottom - 1D || middle > top + 1D) return false;
            double spanRight = span.X + Math.Max(0D, span.Advance);
            return spanRight >= left && span.X <= right;
        }

        // Character indexes of the run whose glyph centres lie between left and right.
        private static IEnumerable<int> GetCoveredCharacters(PdfTextSpan span, double left, double right) {
            if (span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs) && glyphs.Count > 0) {
                foreach (PdfTextGlyph glyph in glyphs) {
                    double center = glyph.X + glyph.Advance / 2D;
                    if (center < left || center > right) continue;
                    for (int character = glyph.TextStart; character < glyph.TextStart + glyph.TextLength && character < span.Text.Length; character++) {
                        yield return character;
                    }
                }
                yield break;
            }
            double width = Math.Max(0D, span.Advance) / span.Text.Length;
            for (int character = 0; character < span.Text.Length; character++) {
                double center = span.X + width * (character + 0.5D);
                if (center >= left && center <= right) yield return character;
            }
        }

        // Walks the runs in line order and finds each one in the text after the previous run, ignoring whitespace
        // differences, so every painted character maps to its own occurrence in the rendered text.
        private static List<AlignedSpan> AlignRuns(string text, IReadOnlyList<PdfLogicalTextBlock> lines) {
            var result = new List<AlignedSpan>();
            int cursor = 0;
            foreach (PdfLogicalTextBlock line in lines) {
                foreach (PdfTextSpan span in line.Spans) {
                    if (string.IsNullOrEmpty(span.Text) ||
                        !PdfLogicalTextBlock.TryFindNormalizedSpan(text, span.Text, cursor, out int index, out int end) ||
                        CountNonWhitespace(text, cursor, index) > MaximumAlignmentSkip) continue;
                    var indexes = new int[span.Text.Length];
                    int target = index;
                    for (int source = 0; source < span.Text.Length; source++) {
                        if (char.IsWhiteSpace(span.Text[source])) {
                            indexes[source] = -1;
                            continue;
                        }
                        while (target < end && char.IsWhiteSpace(text[target])) target++;
                        indexes[source] = target < end ? target++ : -1;
                    }
                    result.Add(new AlignedSpan(span, indexes));
                    cursor = end;
                }
            }
            return result;
        }

        private static int CountNonWhitespace(string text, int start, int end) {
            int count = 0;
            for (int index = start; index < end; index++) {
                if (!char.IsWhiteSpace(text[index])) count++;
            }
            return count;
        }
    }
}
