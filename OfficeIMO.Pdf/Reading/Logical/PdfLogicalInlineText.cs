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
        bool unmeasured = line.Spans.Count == 0;
        foreach (PdfTextSpan span in line.Spans) {
            if (string.IsNullOrEmpty(span.Text)) continue;
            if (!PdfLogicalTextBlock.TryFindNormalizedSpan(text, span.Text, cursor, out int start, out int end)) {
                unmeasured = true;
                continue;
            }
            cursor = end;
            int count = end - start;
            if (count <= 0) continue;
            if (Math.Abs(span.RotationDegrees) > 0.5D) {
                unmeasured = true;
                continue;
            }
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

        // Rotated or unmapped text cannot be split by position. A link over such a line keeps the whole line, as
        // before splitting, rather than being written again on its own.
        if (unmeasured && Array.TrueForAll(owners, static owner => owner is null)) {
            int first = 0;
            int last = text.Length - 1;
            while (first <= last && char.IsWhiteSpace(text[first])) first++;
            while (last >= first && char.IsWhiteSpace(text[last])) last--;
            for (int index = first; index <= last; index++) owners[index] = candidates[0];
            return owners;
        }

        // Link rectangles are drawn by hand and can end inside a word; a word belongs to the link covering most of
        // it. Punctuation at its edges ("archive." or "(page)") keeps its own measured position. A word that two
        // links share ("Alpha/Beta" with a link over each part) is decided part by part.
        var measuredOwners = (PdfLogicalLinkAnnotation?[])owners.Clone();
        for (int wordStart = 0; wordStart < text.Length;) {
            if (!char.IsLetterOrDigit(text[wordStart])) {
                wordStart++;
                continue;
            }
            int wordEnd = wordStart;
            while (wordEnd < text.Length && !char.IsWhiteSpace(text[wordEnd])) wordEnd++;
            while (wordEnd > wordStart + 1 && !char.IsLetterOrDigit(text[wordEnd - 1])) wordEnd--;
            if (HasSeveralLinks(measuredOwners, wordStart, wordEnd)) {
                for (int partStart = wordStart; partStart < wordEnd;) {
                    if (!char.IsLetterOrDigit(text[partStart])) {
                        partStart++;
                        continue;
                    }
                    int partEnd = partStart;
                    while (partEnd < wordEnd && char.IsLetterOrDigit(text[partEnd])) partEnd++;
                    // Two links inside one run of letters ("Open" and "AI" of "OpenAI") keep their drawn boundaries.
                    if (!HasSeveralLinks(measuredOwners, partStart, partEnd)) AssignMostCommonOwner(owners, measuredOwners, partStart, partEnd);
                    partStart = partEnd;
                }
            } else {
                AssignMostCommonOwner(owners, measuredOwners, wordStart, wordEnd);
            }
            wordStart = wordEnd;
        }

        // A link that covers only part of a word ("AI" in "OpenAI") would lose all its text to the word around it
        // and then be written again on its own; it keeps the characters drawn under it instead.
        foreach (PdfLogicalLinkAnnotation link in candidates) {
            if (Array.IndexOf(owners, link) >= 0 || Array.IndexOf(measuredOwners, link) < 0) continue;
            for (int index = 0; index < text.Length; index++) {
                if (owners[index] is null && ReferenceEquals(measuredOwners[index], link)) owners[index] = link;
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

    /// <summary>Finds where the character at <paramref name="index"/> of the line text begins on the page.</summary>
    internal static bool TryGetCharacterStart(PdfLogicalTextBlock line, int index, out double x) {
        x = 0D;
        string text = line.Text;
        int cursor = 0;
        foreach (PdfTextSpan span in line.Spans) {
            if (string.IsNullOrEmpty(span.Text)) continue;
            if (!PdfLogicalTextBlock.TryFindNormalizedSpan(text, span.Text, cursor, out int start, out int end)) continue;
            cursor = end;
            if (index < start || index >= end) continue;
            if (Math.Abs(span.RotationDegrees) > 0.5D) return false;
            string spanText = span.Text;
            if (PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries) &&
                boundaries.Length == spanText.Length + 1) {
                // Non-space characters of the line and the span correspond one to one (TryFindNormalizedSpan).
                int preceding = 0;
                for (int position = start; position < index; position++) {
                    if (!char.IsWhiteSpace(text[position])) preceding++;
                }
                int source = 0;
                while (source < spanText.Length && (char.IsWhiteSpace(spanText[source]) || preceding-- > 0)) source++;
                x = span.X + boundaries[source];
            } else {
                x = span.X + span.Advance * (index - start) / Math.Max(1, end - start);
            }
            return true;
        }
        return false;
    }

    private static void AssignMostCommonOwner(PdfLogicalLinkAnnotation?[] owners, PdfLogicalLinkAnnotation?[] measured, int start, int end) {
        PdfLogicalLinkAnnotation? owner = MostCommonOwner(measured, start, end);
        for (int index = start; index < end; index++) owners[index] = owner;
    }

    private static bool HasSeveralLinks(PdfLogicalLinkAnnotation?[] owners, int start, int end) {
        PdfLogicalLinkAnnotation? first = null;
        for (int index = start; index < end; index++) {
            PdfLogicalLinkAnnotation? owner = owners[index];
            if (owner is null) continue;
            if (first is null) first = owner;
            else if (!ReferenceEquals(first, owner)) return true;
        }
        return false;
    }

    /// <summary>The link drawn over most characters of a word; null when most of it is outside every link.</summary>
    private static PdfLogicalLinkAnnotation? MostCommonOwner(PdfLogicalLinkAnnotation?[] owners, int start, int end) {
        // One pass; a line has only the few links that overlap it, so the tally stays short.
        var links = new List<PdfLogicalLinkAnnotation>(2);
        var counts = new List<int>(2);
        int unowned = 0;
        for (int index = start; index < end; index++) {
            PdfLogicalLinkAnnotation? owner = owners[index];
            if (owner is null) {
                unowned++;
                continue;
            }
            int slot = 0;
            while (slot < links.Count && !ReferenceEquals(links[slot], owner)) slot++;
            if (slot == links.Count) {
                links.Add(owner);
                counts.Add(0);
            }
            counts[slot]++;
        }

        PdfLogicalLinkAnnotation? best = null;
        int bestCount = 0;
        for (int slot = 0; slot < links.Count; slot++) {
            if (counts[slot] > bestCount) {
                best = links[slot];
                bestCount = counts[slot];
            }
        }
        return bestCount > unowned ? best : null;
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
