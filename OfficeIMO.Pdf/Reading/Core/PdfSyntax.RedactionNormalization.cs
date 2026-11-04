using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfSyntax {
    /// <summary>
    /// Independently follows the active cross-reference chain and checks that every winning
    /// entry, including deletion, agrees with the objects materialized by the reader.
    /// </summary>
    internal static bool CanNormalizeRedactionObjectGraph(byte[] pdf, PdfReadDocument document, CancellationToken cancellationToken) {
        if (document.RepairReport.HasIncompleteObjectCoverage || !CanSerializeRedactionObjectGraph(document, cancellationToken)) return false;
        if (!RequiresRedactionNormalization(document)) return true;
        if (document.Security.HasEncryption && HasCompressedRedactionStructures(document)) return false;
        var objects = document.Objects;
        var budget = new PdfDecodedStreamBudget(document.ReadOptions.Limits);
        if (!TryReadRedactionXrefChain(pdf, document, budget, out var entries, out bool hasXrefStreams, cancellationToken) ||
            document.Security.HasEncryption && hasXrefStreams) return false;
        var offsets = objects.ToDictionary(static pair => pair.Key, static pair => pair.Value.SourceOffset);
        foreach (PdfIndirectObject item in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!entries.TryGetValue(item.ObjectNumber, out XrefStreamEntry entry)) return false;
            if (entry.Type == 1) {
                if (entry.Field1 != item.SourceOffset || entry.Field2 != item.Generation) return false;
            } else if (entry.Type == 2) {
                if (entry.Field1 > int.MaxValue || entry.Field1 < 1 || entry.Field2 > int.MaxValue || entry.Field2 < 0) return false;
                if (!TryParseObjectFromObjectStream(objects, offsets, (int)entry.Field1, (int)entry.Field2,
                    item.ObjectNumber, document.ReadOptions.Limits, budget, out _, out int selectedOffset, cancellationToken) ||
                    selectedOffset != item.SourceOffset || item.Generation != 0) return false;
            } else return false;
        }
        // Every in-use entry must have materialized; missing members cannot silently disappear.
        return entries.Values.All(entry => entry.Type == 0 || objects.ContainsKey(entry.ObjectNumber));
    }
    /// <summary>Only live document objects can survive; structural containers cannot be copied as document data.</summary>
    private static bool CanSerializeRedactionObjectGraph(PdfReadDocument document, CancellationToken cancellationToken) {
        PdfDictionary? catalog = FindCatalog(document.Objects, document.TrailerRaw, cancellationToken: cancellationToken);
        if (catalog is null) return false;
        var collector = new PdfPageExtractor.ObjectCollector(document.Objects, undefinedReferencesAreNull: true, cancellationToken: cancellationToken);
        try {
            collector.CollectObjectGraph(catalog);
        } catch (InvalidOperationException) {
            // Generation disagreement is not a proven undefined reference.
            return false;
        }
        foreach (int number in collector.ObjectIds) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject value = document.Objects[number].Value;
            if (value is PdfStream stream && stream.Dictionary.Get<PdfName>("Type")?.Name is "ObjStm" or "XRef" ||
                value is PdfDictionary dictionary && dictionary.Items.ContainsKey("Linearized")) return false;
        }
        return true;
    }

    /// <summary>Uses decoded types: ordinary resource names are not structural containers.</summary>
    internal static bool HasCompressedRedactionStructures(PdfReadDocument document) =>
        document.Objects.Values.Any(static item => item.Value is PdfStream stream &&
            stream.Dictionary.Get<PdfName>("Type")?.Name is "ObjStm" or "XRef");

    /// <summary>Includes history whose compressed containers have been freed from the live map.</summary>
    internal static bool RequiresRedactionNormalization(PdfReadDocument document) {
        if (HasCompressedRedactionStructures(document)) return true;
        int index = 0;
        // Unreadable structure is not proof that the source has no history.
        return !TryReadNextTrailerDictionary(document.TrailerRaw, ref index, document.ReadOptions.Limits,
            out PdfDictionary trailer, CancellationToken.None) || trailer.HasIncompleteSyntax ||
            trailer.Items.ContainsKey("Prev") || trailer.Items.ContainsKey("XRefStm");
    }

    /// <summary>Checks the active cross-reference section and typed objects, not raw marker spellings.</summary>
    internal static bool IsClassicRedactionOutput(byte[] pdf, PdfReadDocument document, CancellationToken cancellationToken) {
        if (HasCompressedRedactionStructures(document) || document.RepairReport.HasIncompleteObjectCoverage ||
            document.Objects.Values.Any(static item => item.Value is PdfDictionary dictionary && dictionary.Items.ContainsKey("Linearized"))) return false;
        string text = PdfEncoding.Latin1GetStringCancellable(pdf, cancellationToken);
        return TryGetLatestStartXrefOffset(text, out int offset, cancellationToken) &&
            TryParseClassicXrefTable(text, offset, out _, out int? previous, out _, out int? hybrid, requireComplete: true, cancellationToken) &&
            !previous.HasValue && !hybrid.HasValue;
    }

    /// <summary>Inspects every container member, including members absent from the live xref.</summary>
    internal static IEnumerable<(int Number, PdfObject Value)> ReadRedactionObjectStreamMembers(PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects, PdfReadLimits limits, PdfDecodedStreamBudget budget, CancellationToken cancellationToken) {
        byte[] data = budget.Decode(stream, objects, limits.MaxDecodedStreamBytes, cancellationToken);
        if (!TryReadObjectStreamLayout(stream.Dictionary, data.Length, limits, out int count, out int first))
            throw new InvalidDataException("The object stream layout cannot be verified.");
        var pairs = ParsePairs(PdfEncoding.Latin1GetStringCancellable(CopyBytes(data, 0, first, cancellationToken), cancellationToken),
            count, out bool complete, cancellationToken);
        if (!complete) throw new InvalidDataException("The object stream header cannot be verified.");
        for (int index = 0; index < count; index++) {
            int start = first + pairs[index].Off;
            int end = index + 1 < count ? first + pairs[index + 1].Off : data.Length;
            if (start < first || end > data.Length || end <= start) throw new InvalidDataException("The object stream member cannot be verified.");
            PdfObject? value = ParseTopLevelObject(PdfEncoding.Latin1GetStringCancellable(CopyBytes(data, start, end - start, cancellationToken), cancellationToken),
                limits, trackEncodedStringSourceSpans: false, cancellationToken: cancellationToken);
            if (value is null || value.HasIncompleteSyntax) throw new InvalidDataException("The object stream member cannot be verified.");
            yield return (pairs[index].Obj, value);
        }
    }

}
