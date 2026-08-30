namespace OfficeIMO.Pdf;

/// <summary>Controls conservative recovery of paragraphs split across adjacent PDF pages.</summary>
public sealed class PdfLogicalParagraphContinuationOptions {
    /// <summary>Whether adjacent page-edge paragraph segments may be merged. Default: true.</summary>
    public bool MergePageContinuations { get; init; } = true;

    /// <summary>
    /// Whether a likely discretionary line-ending hyphen may be removed while joining segments.
    /// Default: false, because PDF layout evidence alone cannot distinguish every authored hyphen.
    /// </summary>
    public bool RejoinLineEndingHyphens { get; init; }

    /// <summary>Maximum adjacent page segments in one recovered paragraph. Default: 16.</summary>
    public int MaximumSegmentsPerParagraph { get; init; } = 16;

    /// <summary>Maximum horizontal geometry difference in visual PDF points. Default: 24.</summary>
    public double GeometryTolerancePoints { get; init; } = 24D;

    /// <summary>Minimum normalized continuation confidence required for a merge. Default: 0.75.</summary>
    public double MinimumConfidence { get; init; } = 0.75D;

    internal static PdfLogicalParagraphContinuationOptions Resolve(PdfLogicalParagraphContinuationOptions? options) {
        PdfLogicalParagraphContinuationOptions effective = options ?? new PdfLogicalParagraphContinuationOptions();
        if (effective.MaximumSegmentsPerParagraph <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MaximumSegmentsPerParagraph, "Maximum paragraph segments must be positive.");
        }
        if (double.IsNaN(effective.GeometryTolerancePoints) ||
            double.IsInfinity(effective.GeometryTolerancePoints) ||
            effective.GeometryTolerancePoints < 0D) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.GeometryTolerancePoints, "Paragraph geometry tolerance must be finite and nonnegative.");
        }
        if (double.IsNaN(effective.MinimumConfidence) ||
            double.IsInfinity(effective.MinimumConfidence) ||
            effective.MinimumConfidence < 0D ||
            effective.MinimumConfidence > 1D) {
            throw new ArgumentOutOfRangeException(nameof(options), effective.MinimumConfidence, "Paragraph continuation confidence must be between zero and one.");
        }
        return effective;
    }
}

/// <summary>Evidence supporting one or more recovered cross-page paragraph boundaries.</summary>
[Flags]
public enum PdfLogicalParagraphContinuationEvidence {
    /// <summary>No cross-page continuation was inferred.</summary>
    None = 0,
    /// <summary>The segments came from adjacent source pages.</summary>
    AdjacentPages = 1,
    /// <summary>The segments were the last and first page paragraphs.</summary>
    BoundaryParagraphs = 2,
    /// <summary>The segments were positioned near the bottom and top page edges.</summary>
    PageEdges = 4,
    /// <summary>The segments had compatible horizontal geometry.</summary>
    CompatibleGeometry = 8,
    /// <summary>The boundary lines had compatible font sizes.</summary>
    CompatibleTypography = 16,
    /// <summary>The preceding segment did not end with strong terminal punctuation.</summary>
    IncompleteTerminal = 32,
    /// <summary>The following segment began with a lower-case letter.</summary>
    LowercaseContinuation = 64,
    /// <summary>The preceding segment ended with a likely discretionary hyphen.</summary>
    HyphenatedBreak = 128
}

/// <summary>One logical paragraph reconstructed from one or more page-level paragraph segments.</summary>
public sealed class PdfLogicalParagraphContinuationGroup {
    internal PdfLogicalParagraphContinuationGroup(
        IReadOnlyList<PdfLogicalParagraph> segments,
        string text,
        double confidence,
        PdfLogicalParagraphContinuationEvidence evidence,
        int rejoinedHyphenCount) {
        Segments = segments;
        Text = text;
        Confidence = confidence;
        Evidence = evidence;
        RejoinedHyphenCount = rejoinedHyphenCount;
    }

    /// <summary>Page-level paragraph segments contributing to this logical paragraph.</summary>
    public IReadOnlyList<PdfLogicalParagraph> Segments { get; }

    /// <summary>Recovered paragraph text.</summary>
    public string Text { get; }

    /// <summary>Lowest normalized confidence across recovered page boundaries, or 1 for an unmerged paragraph.</summary>
    public double Confidence { get; }

    /// <summary>Combined evidence supporting the recovered page boundaries.</summary>
    public PdfLogicalParagraphContinuationEvidence Evidence { get; }

    /// <summary>Number of strongly evidenced line-ending hyphens removed while joining segments.</summary>
    public int RejoinedHyphenCount { get; }

    /// <summary>True when this logical paragraph combines more than one page-level segment.</summary>
    public bool SpansPages => Segments.Count > 1;

    /// <summary>One-based source page number of the first segment.</summary>
    public int FirstPageNumber => Segments[0].PageNumber;

    /// <summary>One-based source page number of the last segment.</summary>
    public int LastPageNumber => Segments[Segments.Count - 1].PageNumber;
}

/// <summary>Conservative, bounded cross-page paragraph continuation analysis.</summary>
public static class PdfLogicalParagraphContinuations {
    /// <summary>
    /// Returns every paragraph in document order, grouping only adjacent page-edge segments whose
    /// geometry, typography, and text-boundary evidence meet the configured confidence threshold.
    /// </summary>
    public static IReadOnlyList<PdfLogicalParagraphContinuationGroup> Group(
        PdfLogicalDocument document,
        PdfLogicalParagraphContinuationOptions? options = null) {
        Guard.NotNull(document, nameof(document));
        PdfLogicalParagraphContinuationOptions effective = PdfLogicalParagraphContinuationOptions.Resolve(options);
        var groups = new List<PdfLogicalParagraphContinuationGroup>();
        var current = new GroupBuilder();

        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            PdfLogicalPage page = document.Pages[pageIndex];
            for (int paragraphIndex = 0; paragraphIndex < page.Paragraphs.Count; paragraphIndex++) {
                PdfLogicalParagraph paragraph = page.Paragraphs[paragraphIndex];
                if (current.Count > 0 &&
                    effective.MergePageContinuations &&
                    current.Count < effective.MaximumSegmentsPerParagraph &&
                    TryContinue(
                        document,
                        current.Last!,
                        paragraph,
                        pageIndex,
                        paragraphIndex,
                        effective,
                        out ContinuationBoundary boundary)) {
                    current.Add(paragraph, boundary);
                    continue;
                }

                if (current.Count > 0) groups.Add(current.Build(effective.RejoinLineEndingHyphens));
                current = new GroupBuilder();
                current.Add(paragraph);
            }
        }

        if (current.Count > 0) groups.Add(current.Build(effective.RejoinLineEndingHyphens));
        return groups.AsReadOnly();
    }

    private static bool TryContinue(
        PdfLogicalDocument document,
        PdfLogicalParagraph previous,
        PdfLogicalParagraph current,
        int currentPageIndex,
        int currentParagraphIndex,
        PdfLogicalParagraphContinuationOptions options,
        out ContinuationBoundary boundary) {
        boundary = default;
        if (currentPageIndex <= 0 || currentParagraphIndex != 0) return false;
        PdfLogicalPage previousPage = document.Pages[currentPageIndex - 1];
        PdfLogicalPage currentPage = document.Pages[currentPageIndex];
        if (previousPage.Paragraphs.Count == 0 ||
            !ReferenceEquals(previousPage.Paragraphs[previousPage.Paragraphs.Count - 1], previous) ||
            current.PageNumber != previous.PageNumber + 1 ||
            currentPage.PageNumber != previousPage.PageNumber + 1) return false;
        if (!TryGetVisualBounds(previous, previousPage, out PdfVisualBounds previousBounds) ||
            !TryGetVisualBounds(current, currentPage, out PdfVisualBounds currentBounds)) return false;
        (_, double previousPageHeight) = previousPage.GetVisualPageSize();
        (_, double currentPageHeight) = currentPage.GetVisualPageSize();
        if (previousBounds.Bottom < previousPageHeight * 0.78D ||
            currentBounds.Top > Math.Max(18D, currentPageHeight * 0.22D)) return false;

        if (Math.Abs(previousBounds.Left - currentBounds.Left) > options.GeometryTolerancePoints) return false;

        double previousFontSize = BoundaryFontSize(previous, useLastLine: true);
        double currentFontSize = BoundaryFontSize(current, useLastLine: false);
        double fontDifference = Math.Abs(previousFontSize - currentFontSize);
        if (fontDifference > Math.Max(1.5D, Math.Max(previousFontSize, currentFontSize) * 0.2D)) return false;

        string previousText = previous.Text.TrimEnd();
        string currentText = current.Text.TrimStart();
        if (previousText.Length == 0 || currentText.Length == 0 || HasStrongTerminal(previousText)) return false;

        PdfLogicalParagraphContinuationEvidence evidence =
            PdfLogicalParagraphContinuationEvidence.AdjacentPages |
            PdfLogicalParagraphContinuationEvidence.BoundaryParagraphs |
            PdfLogicalParagraphContinuationEvidence.PageEdges |
            PdfLogicalParagraphContinuationEvidence.CompatibleGeometry |
            PdfLogicalParagraphContinuationEvidence.CompatibleTypography |
            PdfLogicalParagraphContinuationEvidence.IncompleteTerminal;
        double confidence = 0.65D;
        char first = FirstMeaningfulCharacter(currentText);
        bool startsLowercase = char.IsLetter(first) && char.IsLower(first);
        if (startsLowercase) {
            evidence |= PdfLogicalParagraphContinuationEvidence.LowercaseContinuation;
            confidence += 0.15D;
        }

        bool rejoinHyphen = startsLowercase && HasLikelyDiscretionaryHyphen(previousText);
        if (rejoinHyphen) {
            evidence |= PdfLogicalParagraphContinuationEvidence.HyphenatedBreak;
            confidence += 0.1D;
        }
        if (Math.Abs(previousBounds.Left - currentBounds.Left) <= options.GeometryTolerancePoints * 0.5D) confidence += 0.05D;
        if (fontDifference <= 0.5D) confidence += 0.05D;
        confidence = Math.Min(1D, confidence);
        if (confidence < options.MinimumConfidence) return false;

        boundary = new ContinuationBoundary(confidence, evidence, rejoinHyphen);
        return true;
    }

    private static bool TryGetVisualBounds(
        PdfLogicalParagraph paragraph,
        PdfLogicalPage page,
        out PdfVisualBounds bounds) {
        if (paragraph.Lines.Count == 0 || paragraph.XEnd <= paragraph.XStart) {
            bounds = default;
            return false;
        }
        double fontSize = Math.Max(1D, paragraph.Lines.Max(static line => line.FontSize));
        double bottom = Math.Min(paragraph.YBottom, paragraph.YTop) - fontSize * 0.25D;
        double top = Math.Max(paragraph.YBottom, paragraph.YTop) + fontSize;
        bounds = page.TransformBoundsToVisual(paragraph.XStart, bottom, paragraph.XEnd, top);
        return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
    }

    private static double BoundaryFontSize(PdfLogicalParagraph paragraph, bool useLastLine) {
        if (paragraph.Lines.Count == 0) return 0D;
        return useLastLine
            ? paragraph.Lines[paragraph.Lines.Count - 1].FontSize
            : paragraph.Lines[0].FontSize;
    }

    private static char FirstMeaningfulCharacter(string text) {
        for (int i = 0; i < text.Length; i++) {
            char value = text[i];
            if (char.IsLetterOrDigit(value)) return value;
        }
        return '\0';
    }

    private static bool HasStrongTerminal(string text) {
        int index = text.Length - 1;
        while (index >= 0 && (char.IsWhiteSpace(text[index]) || text[index] == '"' || text[index] == '\'' || text[index] == ')' || text[index] == ']' || text[index] == '}')) index--;
        if (index < 0) return true;
        char value = text[index];
        return value == '.' || value == '!' || value == '?' || value == ':';
    }

    private static bool HasLikelyDiscretionaryHyphen(string text) {
        if (text.Length < 2 || text[text.Length - 1] != '-') return false;
        return char.IsLetter(text[text.Length - 2]) && (text.Length < 3 || !char.IsWhiteSpace(text[text.Length - 2]));
    }

    private readonly struct ContinuationBoundary {
        internal ContinuationBoundary(
            double confidence,
            PdfLogicalParagraphContinuationEvidence evidence,
            bool rejoinHyphen) {
            Confidence = confidence;
            Evidence = evidence;
            RejoinHyphen = rejoinHyphen;
        }

        internal double Confidence { get; }
        internal PdfLogicalParagraphContinuationEvidence Evidence { get; }
        internal bool RejoinHyphen { get; }
    }

    private sealed class GroupBuilder {
        private readonly List<PdfLogicalParagraph> _segments = new();
        private readonly List<ContinuationBoundary> _boundaries = new();

        internal int Count => _segments.Count;
        internal PdfLogicalParagraph? Last => _segments.Count == 0 ? null : _segments[_segments.Count - 1];

        internal void Add(PdfLogicalParagraph paragraph) => _segments.Add(paragraph);

        internal void Add(PdfLogicalParagraph paragraph, ContinuationBoundary boundary) {
            _segments.Add(paragraph);
            _boundaries.Add(boundary);
        }

        internal PdfLogicalParagraphContinuationGroup Build(bool rejoinLineEndingHyphens) {
            var text = new StringBuilder();
            int rejoinedHyphens = 0;
            PdfLogicalParagraphContinuationEvidence evidence = PdfLogicalParagraphContinuationEvidence.None;
            double confidence = 1D;
            for (int index = 0; index < _segments.Count; index++) {
                string segmentText = _segments[index].Text.Trim();
                if (index == 0) {
                    text.Append(segmentText);
                    continue;
                }
                ContinuationBoundary boundary = _boundaries[index - 1];
                confidence = Math.Min(confidence, boundary.Confidence);
                evidence |= boundary.Evidence;
                if (rejoinLineEndingHyphens && boundary.RejoinHyphen && text.Length > 0 && text[text.Length - 1] == '-') {
                    text.Length--;
                    rejoinedHyphens++;
                } else if (text.Length > 0 && segmentText.Length > 0) {
                    text.Append(' ');
                }
                text.Append(segmentText);
            }

            return new PdfLogicalParagraphContinuationGroup(
                _segments.ToArray(),
                text.ToString(),
                confidence,
                evidence,
                rejoinedHyphens);
        }
    }
}

public sealed partial class PdfLogicalDocument {
    /// <summary>Returns conservative cross-page paragraph continuation groups in document order.</summary>
    public IReadOnlyList<PdfLogicalParagraphContinuationGroup> GetParagraphContinuationGroups(
        PdfLogicalParagraphContinuationOptions? options = null) =>
        PdfLogicalParagraphContinuations.Group(this, options);
}
