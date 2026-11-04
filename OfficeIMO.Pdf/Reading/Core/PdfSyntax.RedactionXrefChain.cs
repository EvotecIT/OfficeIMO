using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    /// <summary>Reads newest entries first: table, its hybrid supplement, then the previous section.</summary>
    private static bool TryReadRedactionXrefChain(byte[] pdf, PdfReadDocument document, PdfDecodedStreamBudget budget,
        out Dictionary<int, XrefStreamEntry> winners, out bool hasXrefStreams, CancellationToken cancellationToken) {
        winners = new Dictionary<int, XrefStreamEntry>();
        hasXrefStreams = false;
        var objects = document.Objects;
        PdfReadLimits limits = document.ReadOptions.Limits;
        string text = PdfEncoding.Latin1GetStringCancellable(pdf, cancellationToken);
        if (!TryGetLatestStartXrefOffset(text, out int start, cancellationToken)) return false;
        var pending = new Stack<(int Offset, bool Supplement)>();
        var visited = new HashSet<int>();
        var scanBudget = new XrefObjectScanBudget(limits);
        var streams = new Dictionary<int, PdfStream>();
        foreach (PdfIndirectObject item in objects.Values) {
            if (item.Value is PdfStream stream && stream.Dictionary.Get<PdfName>("Type")?.Name == "XRef")
                streams[item.SourceOffset] = stream;
        }
        pending.Push((start, false));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var section = pending.Pop();
            if (!visited.Add(section.Offset) || visited.Count > limits.MaxRevisions) return false;
            if (TryParseClassicXrefTable(text, section.Offset, out var table, out int? previous, out _, out int? hybrid, requireComplete: true, cancellationToken)) {
                if (section.Supplement) return false;
                var sectionNumbers = new HashSet<int>();
                foreach (var entry in table) {
                    if (!sectionNumbers.Add(entry.ObjectNumber)) return false;
                    if (!winners.ContainsKey(entry.ObjectNumber))
                        winners.Add(entry.ObjectNumber, new XrefStreamEntry(entry.ObjectNumber, entry.InUse ? 1 : 0, entry.Offset, entry.Generation));
                }
                if (previous.HasValue) pending.Push((previous.Value, false));
                if (hybrid.HasValue) pending.Push((hybrid.Value, true));
            } else {
                hasXrefStreams = true;
                if (!streams.TryGetValue(section.Offset, out PdfStream? stream)) {
                    if (!TryParseIndirectObjectAt(pdf, text, section.Offset, objects, scanBudget, out PdfIndirectObject parsed, cancellationToken) ||
                        parsed.Value is not PdfStream parsedStream || parsedStream.Dictionary.Get<PdfName>("Type")?.Name != "XRef") return false;
                    stream = parsedStream;
                    streams.Add(section.Offset, stream);
                }
                if (stream.Dictionary.HasIncompleteSyntax || stream.Dictionary.Items.ContainsKey("XRefStm")) return false;
                bool incomplete = false;
                byte[] data = budget.Decode(stream, objects, limits.MaxDecodedStreamBytes, cancellationToken);
                var sectionNumbers = new HashSet<int>();
                foreach (XrefStreamEntry entry in ReadXrefStreamEntries(stream.Dictionary, data, () => incomplete = true, cancellationToken)) {
                    if (!sectionNumbers.Add(entry.ObjectNumber)) return false;
                    if (!winners.TryGetValue(entry.ObjectNumber, out _)) winners.Add(entry.ObjectNumber, entry);
                    if (winners.Count > limits.MaxIndirectObjects) return false;
                }
                if (incomplete) return false;
                // A hybrid supplement belongs to its table's revision. The table's /Prev,
                // rather than a second chronology in the supplement, owns the next section.
                if (!section.Supplement && stream.Dictionary.Items.TryGetValue("Prev", out PdfObject? prev)) {
                    if (!TryReadXrefInteger(prev, out int offset)) return false;
                    pending.Push((offset, false));
                }
            }
            if (winners.Count > limits.MaxIndirectObjects) return false;
        }
        return winners.Count > 0;
    }
}
