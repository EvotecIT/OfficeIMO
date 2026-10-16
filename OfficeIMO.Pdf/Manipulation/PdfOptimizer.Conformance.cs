using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfOptimizer {
    /// <summary>Rejects settings that cannot preserve a declared archival profile before serialization.</summary>
    private static void ValidateConformanceSettings(PdfXmpMetadataInfo? xmp, bool hasMetadata, PdfOptimizationOptions options) {
        if (hasMetadata && (xmp is null || !xmp.IsWellFormedXml) &&
            (options.UseObjectStreams || options.XrefFormat == PdfOptimizationXrefFormat.XrefStream)) {
            throw new NotSupportedException("Object streams and cross-reference streams require readable XMP to establish the declared conformance. Use classic cross references or supply metadata within the supported XML read limits.");
        }
        if (xmp?.PdfAPart == 1 && (options.UseObjectStreams || options.XrefFormat == PdfOptimizationXrefFormat.XrefStream)) {
            throw new NotSupportedException("PDF/A-1 forbids object streams and cross-reference streams. Use Balanced, Archival or Web, or disable UseObjectStreams and select ClassicTable.");
        }
    }

    /// <summary>Retains the encoded payload and dictionary of metadata, attached files and output profiles.</summary>
    private static HashSet<PdfStream> FindPreservedStreams(
        Dictionary<int, PdfIndirectObject> objects, CancellationToken cancellationToken) {
        var preserved = new HashSet<PdfStream>();
        var pending = new Stack<(PdfObject Value, bool Preserve)>();
        var visited = new HashSet<(int ObjectNumber, bool Preserve)>();
        foreach (PdfIndirectObject indirect in objects.Values) pending.Push((indirect.Value, false));
        while (pending.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var (value, preserve) = pending.Pop();
            if (value is PdfReference reference) {
                if (visited.Add((reference.ObjectNumber, preserve)) &&
                    PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                    pending.Push((indirect.Value, preserve));
                }
                continue;
            }
            if (value is PdfArray array) {
                foreach (PdfObject child in array.Items) pending.Push((child, preserve));
                continue;
            }
            PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
            if (dictionary is null) continue;
            if (value is PdfStream payload) {
                string? type = ReadName(dictionary, "Type");
                bool metadata = string.Equals(type, "Metadata", StringComparison.Ordinal);
                if (metadata && dictionary.Items.ContainsKey("Filter")) {
                    throw new NotSupportedException("Optimization requires uncompressed XMP metadata so its packet can remain unchanged. Supply a PDF with an unfiltered Metadata stream.");
                }
                if (preserve || metadata || string.Equals(type, "EmbeddedFile", StringComparison.Ordinal)) preserved.Add(payload);
            }
            foreach (KeyValuePair<string, PdfObject> entry in dictionary.Items) {
                bool protectChild = entry.Key == "Metadata" || entry.Key == "EF" || entry.Key == "DestOutputProfile";
                pending.Push((entry.Value, preserve || protectChild));
            }
        }
        return preserved;
    }
}
