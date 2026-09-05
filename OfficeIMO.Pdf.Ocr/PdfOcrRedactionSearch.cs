using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>One privacy-conscious OCR-derived redaction candidate.</summary>
public sealed class PdfOcrRedactionCandidate {
    internal PdfOcrRedactionCandidate(PdfRedactionArea area, string criterion, double minimumConfidence, string? provider, string? model, string? language) {
        Area = area;
        Criterion = criterion;
        MinimumConfidence = minimumConfidence;
        Provider = provider;
        Model = model;
        Language = language;
    }

    /// <summary>Canonical PDF user-space rectangle covering every OCR word in the match.</summary>
    public PdfRedactionArea Area { get; }
    /// <summary>Stable criterion description. The recognized text itself is intentionally omitted.</summary>
    public string Criterion { get; }
    /// <summary>Lowest OCR confidence among the words contributing to this candidate.</summary>
    public double MinimumConfidence { get; }
    /// <summary>OCR provider identifier, when reported.</summary>
    public string? Provider { get; }
    /// <summary>OCR model identifier, when reported.</summary>
    public string? Model { get; }
    /// <summary>Detected or requested language, when reported.</summary>
    public string? Language { get; }
}

/// <summary>OCR merge evidence and redaction rectangles derived from bounded search criteria.</summary>
public sealed class PdfOcrRedactionSearchResult {
    internal PdfOcrRedactionSearchResult(PdfOcrMergeResult ocr, IReadOnlyList<PdfOcrRedactionCandidate> candidates) {
        Ocr = ocr;
        Candidates = candidates;
    }

    /// <summary>OCR execution and merge evidence.</summary>
    public PdfOcrMergeResult Ocr { get; }
    /// <summary>Privacy-conscious OCR-derived candidates.</summary>
    public IReadOnlyList<PdfOcrRedactionCandidate> Candidates { get; }
}

/// <summary>OCR-assisted redaction candidate discovery.</summary>
public static class PdfOcrRedactionExtensions {
    /// <summary>
    /// Runs the canonical OCR adapter and maps literal and bounded-regex matches back to PDF user-space areas.
    /// Form fields and logical element kinds remain native-parser criteria and are ignored by this OCR pass.
    /// </summary>
    public static async Task<PdfOcrRedactionSearchResult> SearchRedactionCandidatesWithOcrAsync(
        this PdfDocument document,
        IOcrEngine engine,
        PdfRedactionSearchOptions search,
        PdfOcrMergeOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (engine == null) throw new ArgumentNullException(nameof(engine));
        if (search == null) throw new ArgumentNullException(nameof(search));
        if (search.RegexTimeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(search), "Regex timeout must be positive.");
        if (search.MaximumCandidates <= 0) throw new ArgumentOutOfRangeException(nameof(search), "Maximum candidates must be positive.");
        if (search.LiteralText.Count == 0 && search.RegularExpressions.Count == 0) {
            throw new ArgumentException("OCR redaction search requires at least one literal or regular-expression criterion.", nameof(search));
        }
        using var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, search.CancellationToken);
        CancellationToken effectiveCancellation = linkedCancellation.Token;
        effectiveCancellation.ThrowIfCancellationRequested();

        Regex[] expressions = search.RegularExpressions
            .Select(pattern => new Regex(pattern, search.RegexOptions, search.RegexTimeout))
            .ToArray();
        PdfOcrMergeResult ocr = await document.ReadWithOcrAsync(engine, options, effectiveCancellation).ConfigureAwait(false);
        var candidates = new List<PdfOcrRedactionCandidate>();
        for (int pageIndex = 0; pageIndex < ocr.Pages.Count; pageIndex++) {
            effectiveCancellation.ThrowIfCancellationRequested();
            PdfOcrPageMergeResult page = ocr.Pages[pageIndex];
            if (page.Words.Count == 0) continue;
            PdfLogicalPage logicalPage = ocr.Document.Pages.First(item => item.PageNumber == page.PageNumber);
            foreach (WordTextMap map in WordTextMap.CreateLines(page.Words)) {
                for (int literalIndex = 0; literalIndex < search.LiteralText.Count; literalIndex++) {
                    string literal = search.LiteralText[literalIndex];
                    int start = 0;
                    while (start <= map.Text.Length - literal.Length) {
                        effectiveCancellation.ThrowIfCancellationRequested();
                        int found = map.Text.IndexOf(literal, start, search.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase);
                        if (found < 0) break;
                        AddCandidate(candidates, map, found, literal.Length, page, logicalPage, "literal:" + literalIndex, search.MaximumCandidates);
                        start = found + Math.Max(1, literal.Length);
                    }
                }
                for (int expressionIndex = 0; expressionIndex < expressions.Length; expressionIndex++) {
                    foreach (Match match in expressions[expressionIndex].Matches(map.Text)) {
                        effectiveCancellation.ThrowIfCancellationRequested();
                        if (match.Length > 0) AddCandidate(candidates, map, match.Index, match.Length, page, logicalPage, "regex:" + expressionIndex, search.MaximumCandidates);
                    }
                }
            }
        }
        PdfOcrRedactionCandidate[] distinct = candidates
            .GroupBy(static item => string.Join("|", item.Area.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture), item.Area.X.ToString("R", System.Globalization.CultureInfo.InvariantCulture), item.Area.Y.ToString("R", System.Globalization.CultureInfo.InvariantCulture), item.Area.Width.ToString("R", System.Globalization.CultureInfo.InvariantCulture), item.Area.Height.ToString("R", System.Globalization.CultureInfo.InvariantCulture), item.Criterion), StringComparer.Ordinal)
            .Select(static group => group.First())
            .ToArray();
        return new PdfOcrRedactionSearchResult(ocr, Array.AsReadOnly(distinct));
    }

    private static void AddCandidate(List<PdfOcrRedactionCandidate> target, WordTextMap map, int start, int length, PdfOcrPageMergeResult page, PdfLogicalPage logicalPage, string criterion, int maximumCandidates) {
        PdfRecognizedWord[] words = map.GetWords(start, length);
        if (words.Length == 0) return;
        if (target.Count >= maximumCandidates) throw new InvalidOperationException("OCR redaction search exceeded the configured candidate limit.");
        double left = words.Min(static word => word.X);
        double top = words.Min(static word => word.Y);
        double right = words.Max(static word => word.X + word.Width);
        double bottom = words.Max(static word => word.Y + word.Height);
        PdfPageRectangle rectangle = logicalPage.MapVisualRectangleToUserSpace(left, top, right, bottom);
        target.Add(new PdfOcrRedactionCandidate(
            new PdfRedactionArea(page.PageNumber, rectangle.Left, rectangle.Bottom, rectangle.Width, rectangle.Height, "ocr:" + criterion),
            criterion,
            words.Min(static word => word.Confidence),
            page.Provider,
            page.Model,
            page.Language));
    }

    private sealed class WordTextMap {
        private readonly PdfRecognizedWord[] _words;
        private readonly int[] _starts;
        private WordTextMap(string text, PdfRecognizedWord[] words, int[] starts) { Text = text; _words = words; _starts = starts; }
        internal string Text { get; }

        internal static IReadOnlyList<WordTextMap> CreateLines(IReadOnlyList<PdfRecognizedWord> words) {
            var result = new List<WordTextMap>();
            var line = new List<PdfRecognizedWord>();
            for (int index = 0; index < words.Count; index++) {
                if (line.Count > 0 && !ContinuesLine(line[line.Count - 1], words[index])) {
                    result.Add(Create(line));
                    line.Clear();
                }
                line.Add(words[index]);
            }
            if (line.Count > 0) result.Add(Create(line));
            return result;
        }

        private static WordTextMap Create(IReadOnlyList<PdfRecognizedWord> words) {
            var text = new StringBuilder();
            var starts = new int[words.Count];
            var snapshot = new PdfRecognizedWord[words.Count];
            for (int index = 0; index < words.Count; index++) {
                if (index > 0) text.Append(' ');
                starts[index] = text.Length;
                snapshot[index] = words[index];
                text.Append(words[index].Text);
            }
            return new WordTextMap(text.ToString(), snapshot, starts);
        }

        private static bool ContinuesLine(PdfRecognizedWord previous, PdfRecognizedWord current) {
            if (!string.IsNullOrEmpty(previous.BlockId) && !string.IsNullOrEmpty(current.BlockId) && !string.Equals(previous.BlockId, current.BlockId, StringComparison.Ordinal)) return false;
            if (!string.IsNullOrEmpty(previous.ParagraphId) && !string.IsNullOrEmpty(current.ParagraphId) && !string.Equals(previous.ParagraphId, current.ParagraphId, StringComparison.Ordinal)) return false;
            if (!string.IsNullOrEmpty(previous.LineId) || !string.IsNullOrEmpty(current.LineId)) return string.Equals(previous.LineId, current.LineId, StringComparison.Ordinal);
            double previousCenter = previous.Y + previous.Height / 2D;
            double currentCenter = current.Y + current.Height / 2D;
            double lineTolerance = Math.Max(previous.Height, current.Height) * 0.6D;
            double horizontalGap = current.X - (previous.X + previous.Width);
            double maximumGap = Math.Max(previous.Height, current.Height) * 4D;
            return Math.Abs(previousCenter - currentCenter) <= lineTolerance && horizontalGap >= -lineTolerance && horizontalGap <= maximumGap;
        }

        internal PdfRecognizedWord[] GetWords(int start, int length) {
            int end = checked(start + length);
            var result = new List<PdfRecognizedWord>();
            for (int index = 0; index < _words.Length; index++) {
                int wordStart = _starts[index];
                int wordEnd = wordStart + _words[index].Text.Length;
                if (wordStart < end && wordEnd > start) result.Add(_words[index]);
            }
            return result.ToArray();
        }
    }
}
