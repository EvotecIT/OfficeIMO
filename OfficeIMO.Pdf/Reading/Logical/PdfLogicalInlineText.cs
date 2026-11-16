namespace OfficeIMO.Pdf;

/// <summary>A piece of reflowed text with one style and at most one link annotation over it.</summary>
internal readonly struct PdfLogicalInlineSegment {
    internal PdfLogicalInlineSegment(string text, PdfLogicalTextRun? run, PdfLogicalLinkAnnotation? link) {
        Text = text;
        Run = run;
        Link = link;
    }

    internal string Text { get; }

    /// <summary>Source style run; null for the space that joins two reflowed lines.</summary>
    internal PdfLogicalTextRun? Run { get; }

    internal PdfLogicalLinkAnnotation? Link { get; }

    internal bool IsBold => Run?.IsBold == true;

    internal bool IsItalic => Run?.IsItalic == true;
}

/// <summary>
/// Splits the style runs of reflowed lines where link annotations begin and end, so each link covers only the
/// words drawn under its rectangle instead of the whole line or paragraph.
/// </summary>
internal static class PdfLogicalInlineText {
    private const double VerticalTolerance = 2D;

    internal static bool Overlaps(PdfLogicalLinkAnnotation link, PdfLogicalTextBlock line) {
        bool yOverlaps = line.BaselineY >= link.Y1 - VerticalTolerance && line.BaselineY <= link.Y2 + VerticalTolerance;
        bool xOverlaps = Math.Min(link.X2, line.XEnd) - Math.Max(link.X1, line.XStart) > VerticalTolerance;
        return yOverlaps && xOverlaps;
    }

    /// <summary>
    /// Returns the segments of <paramref name="lines"/> joined by single spaces. Adjacent segments with the same
    /// link and the same visible style are merged.
    /// </summary>
    internal static List<PdfLogicalInlineSegment> Segment(
        IReadOnlyList<PdfLogicalTextBlock> lines,
        IReadOnlyList<PdfLogicalLinkAnnotation> links) {
        var segments = new List<PdfLogicalInlineSegment>();
        for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
            PdfLogicalTextBlock line = lines[lineIndex];
            if (line.Text.Length == 0) continue;
            PdfLogicalLinkAnnotation?[] owners = MapLinks(line, links);
            if (segments.Count > 0) {
                // A link that wraps onto the next line keeps the joining space.
                PdfLogicalLinkAnnotation? previous = segments[segments.Count - 1].Link;
                Add(segments, new PdfLogicalInlineSegment(" ", null, previous != null && ReferenceEquals(previous, FirstOwner(owners)) ? previous : null));
            }

            IReadOnlyList<PdfLogicalTextRun> runs = line.Runs.Count > 0
                ? line.Runs
                : new[] { new PdfLogicalTextRun(line.Text, sourceSpan: null) };
            int position = 0;
            foreach (PdfLogicalTextRun run in runs) {
                int runStart = Math.Min(position, line.Text.Length);
                int runEnd = Math.Min(runStart + run.Text.Length, line.Text.Length);
                position = runEnd;
                int pieceStart = runStart;
                for (int index = runStart + 1; index <= runEnd; index++) {
                    if (index < runEnd && ReferenceEquals(owners[index], owners[pieceStart])) continue;
                    Add(segments, new PdfLogicalInlineSegment(line.Text.Substring(pieceStart, index - pieceStart), run, owners[pieceStart]));
                    pieceStart = index;
                }
            }
        }

        return segments;
    }

    /// <summary>Finds, for each character of the line, the link annotation drawn over it.</summary>
    internal static PdfLogicalLinkAnnotation?[] MapLinks(PdfLogicalTextBlock line, IReadOnlyList<PdfLogicalLinkAnnotation> links) {
        string text = line.Text;
        var owners = new PdfLogicalLinkAnnotation?[text.Length];
        var candidates = new List<PdfLogicalLinkAnnotation>();
        foreach (PdfLogicalLinkAnnotation link in links) {
            if (Overlaps(link, line)) candidates.Add(link);
        }
        if (candidates.Count == 0) return owners;

        int cursor = 0;
        foreach (PdfTextSpan span in line.Spans) {
            if (string.IsNullOrEmpty(span.Text) ||
                !PdfLogicalTextBlock.TryFindNormalizedSpan(text, span.Text, cursor, out int start, out int end)) continue;
            cursor = end;
            int count = end - start;
            if (count <= 0 || Math.Abs(span.RotationDegrees) > 0.5D) continue;
            string spanText = span.Text;
            bool measured = PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries) &&
                boundaries.Length == spanText.Length + 1;
            int source = 0;
            for (int index = start; index < end; index++) {
                if (char.IsWhiteSpace(text[index])) continue;
                // Non-space characters of the line and the span correspond one to one (TryFindNormalizedSpan).
                while (source < spanText.Length && char.IsWhiteSpace(spanText[source])) source++;
                // Without resolved glyph advances, characters share the span advance evenly; a link boundary falls
                // between words, so the estimate only needs to land on the right side of a space.
                double center = measured && source < spanText.Length
                    ? span.X + (boundaries[source] + boundaries[source + 1]) / 2D
                    : span.X + span.Advance * ((index - start) + 0.5D) / count;
                source++;
                foreach (PdfLogicalLinkAnnotation link in candidates) {
                    if (center >= link.X1 - 1D && center <= link.X2 + 1D) {
                        owners[index] = link;
                        break;
                    }
                }
            }
        }

        // Spaces between two words of the same link belong to it; spaces at its edges do not.
        int previousOwned = -1;
        for (int index = 0; index < text.Length; index++) {
            if (char.IsWhiteSpace(text[index])) continue;
            if (previousOwned >= 0 && owners[index] != null && ReferenceEquals(owners[index], owners[previousOwned])) {
                for (int gap = previousOwned + 1; gap < index; gap++) owners[gap] = owners[index];
            }
            previousOwned = index;
        }

        return owners;
    }

    private static PdfLogicalLinkAnnotation? FirstOwner(PdfLogicalLinkAnnotation?[] owners) {
        foreach (PdfLogicalLinkAnnotation? owner in owners) {
            if (owner != null) return owner;
        }
        return null;
    }

    private static void Add(List<PdfLogicalInlineSegment> segments, PdfLogicalInlineSegment segment) {
        if (segment.Text.Length == 0) return;
        if (segments.Count > 0) {
            PdfLogicalInlineSegment previous = segments[segments.Count - 1];
            if (ReferenceEquals(previous.Link, segment.Link) && ((segment.Run is null && IsBlank(segment.Text)) || SameStyle(previous.Run, segment.Run))) {
                segments[segments.Count - 1] = new PdfLogicalInlineSegment(previous.Text + segment.Text, previous.Run ?? segment.Run, previous.Link);
                return;
            }
        }
        segments.Add(segment);
    }

    private static bool IsBlank(string text) {
        foreach (char ch in text) {
            if (!char.IsWhiteSpace(ch)) return false;
        }
        return true;
    }

    /// <summary>True when two runs would be written with the same character formatting.</summary>
    internal static bool SameStyle(PdfLogicalTextRun? left, PdfLogicalTextRun? right) {
        if (ReferenceEquals(left, right)) return true;
        if (left is null || right is null) return false;
        return left.IsBold == right.IsBold &&
            left.IsItalic == right.IsItalic &&
            Math.Abs(left.FontSize - right.FontSize) < 0.1D &&
            string.Equals(left.BaseFont, right.BaseFont, StringComparison.Ordinal) &&
            Nullable.Equals(left.Color, right.Color);
    }
}
