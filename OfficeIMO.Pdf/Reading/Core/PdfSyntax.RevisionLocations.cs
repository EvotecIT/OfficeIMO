using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    // Security markers are observations and may also occur inside payload streams. Definition
    // ownership must use the terminal cross-reference pointer and its parsed /Prev chain.
    private static int[] GetActiveRevisionOffsets(byte[] pdf, string text, Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits, System.Diagnostics.Stopwatch timer, CancellationToken cancellationToken) {
        if (!TryGetLatestStartXrefOffset(text, out int current, cancellationToken)) return Array.Empty<int>();
        var scanBudget = new XrefObjectScanBudget(limits);
        var newestFirst = new List<int>();
        var visited = new HashSet<int>();
        while (visited.Add(current) && newestFirst.Count < limits.MaxRevisions) {
            cancellationToken.ThrowIfCancellationRequested();
            ThrowIfParsingTimeExceeded(timer, limits);
            int? previous;
            if (TryParseClassicXrefTable(text, current, out _, out previous, out _, out _, cancellationToken)) {
                // A hybrid /XRefStm describes the same update; only /Prev changes its owner.
            } else if (TryParseIndirectObjectAt(pdf, text, current, objects, scanBudget, out var indirect, cancellationToken) &&
                indirect.Value is PdfStream stream && stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef") {
                // A previous stream may have been freed or replaced in the final object map.
                var dictionary = stream.Dictionary;
                if (dictionary.Items.TryGetValue("Prev", out var value)) {
                    if (!TryReadXrefInteger(value, out int offset)) return Array.Empty<int>();
                    previous = offset;
                } else previous = null;
            } else return Array.Empty<int>();
            ThrowIfParsingTimeExceeded(timer, limits);
            newestFirst.Add(current);
            if (!previous.HasValue) { newestFirst.Reverse(); return newestFirst.ToArray(); }
            // Incremental updates append their cross-reference data. Invalid cycles or order
            // have no trustworthy chronological projection, even when repair can read objects.
            if (previous.Value >= current) return Array.Empty<int>();
            current = previous.Value;
        }
        return Array.Empty<int>();
    }

    /// <summary>Returns complete earlier PDF updates linked by the active cross-reference chain.</summary>
    internal static int[] GetHistoricalRevisionEnds(byte[] pdf, PdfReadDocument document, CancellationToken cancellationToken) {
        var timer = System.Diagnostics.Stopwatch.StartNew();
        var limits = document.ReadOptions.Limits;
        string text = PdfEncoding.Latin1GetStringCancellable(pdf, cancellationToken);
        int[] offsets = GetActiveRevisionOffsets(pdf, text, document.Objects, limits, timer, cancellationToken);
        var ends = new List<int>();
        for (int index = offsets.Length - 2; index >= 0; index--) {
            cancellationToken.ThrowIfCancellationRequested();
            ThrowIfParsingTimeExceeded(timer, limits);
            int position;
            if (TryParseClassicXrefTable(text, offsets[index], out _, out _, out _, out _, cancellationToken)) {
                int trailer = IndexOfKeywordCancellable(text, "trailer", offsets[index], offsets[index + 1], cancellationToken);
                int dictionary = trailer < 0 ? -1 : SkipWhitespaceAndComments(text, trailer + 7, offsets[index + 1], cancellationToken);
                position = dictionary < 0 ? -1 : FindDictEnd(text, dictionary, offsets[index + 1], cancellationToken);
            } else {
                position = FindObjectEnd(text, offsets[index], maximumIndex: offsets[index + 1], cancellationToken: cancellationToken);
            }
            if (position < 0) throw new InvalidDataException("An earlier PDF update has no complete cross-reference boundary.");
            position = SkipRevisionTrivia(position);
            if (position + 9 > text.Length || string.CompareOrdinal(text, position, "startxref", 0, 9) != 0)
                throw new InvalidDataException("An earlier PDF update has no terminal cross-reference pointer.");
            position += 9;
            position = SkipRevisionTrivia(position);
            int first = position;
            while (position < text.Length && text[position] >= '0' && text[position] <= '9') position++;
            if (!TryParseXrefInteger(text, first, position - first, out int target) || target != offsets[index])
                throw new InvalidDataException("An earlier PDF update's terminal pointer does not match its cross-reference section.");
            position = SkipRevisionTrivia(position);
            if (position + 5 > text.Length || string.CompareOrdinal(text, position, "%%EOF", 0, 5) != 0)
                throw new InvalidDataException("An earlier PDF update has no complete EOF boundary.");
            ends.Add(position + 5);
        }
        ThrowIfParsingTimeExceeded(timer, limits);
        return ends.ToArray();

        int SkipRevisionTrivia(int position) {
            while (position < text.Length) {
                if ((position & 4095) == 0) { cancellationToken.ThrowIfCancellationRequested(); ThrowIfParsingTimeExceeded(timer, limits); }
                if (char.IsWhiteSpace(text[position])) { position++; continue; }
                if (text[position] != '%' || (position + 5 <= text.Length && string.CompareOrdinal(text, position, "%%EOF", 0, 5) == 0)) break;
                while (position < text.Length && text[position] != '\r' && text[position] != '\n') {
                    if ((position & 4095) == 0) { cancellationToken.ThrowIfCancellationRequested(); ThrowIfParsingTimeExceeded(timer, limits); }
                    position++;
                }
            }
            return position;
        }
    }
}
