using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Saves captured physical entry payloads and explicitly supported same-length native mutations.</summary>
internal static class IWorkPreservationPackageWriter {
    internal static IWorkNativeSaveResult Write(IWorkSourceDocument source, IReadOnlyDictionary<IWorkArchiveRecord, byte[]> edits,
        IWorkNativeSaveOptions limits, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (edits.Count > limits.MaxEditedRecords) throw new InvalidDataException("The edited-record limit was exceeded.");
        if (edits.Count > 0 && !limits.AllowUnassessedVisualCaches)
            throw new InvalidOperationException("Edited saves retain native visual caches without refreshing them. Set AllowUnassessedVisualCaches to explicitly accept this reported boundary.");
        var entries = new SortedDictionary<string, byte[]>(StringComparer.Ordinal);
        long inputLength = 0;
        foreach (IWorkPackageEntry entry in source.ContainerEntries) {
            token.ThrowIfCancellationRequested(); inputLength += entry.Length;
            if (inputLength > limits.MaxPackageBytes) throw new InvalidDataException("The native save input-entry limit was exceeded.");
            entries.Add(entry.Path, entry.Bytes);
        }
        var paths = edits.Keys.Select(record => record.EntryPath).Distinct(StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal).ToArray();
        long decompressedBytes = 0;
        foreach (string path in paths) {
            token.ThrowIfCancellationRequested();
            if (!entries.TryGetValue(path, out byte[]? original)) throw new NotSupportedException("The edited archive is not a physical source entry.");
            byte[] raw = IWorkSnappy.DecodeIwa(original, source.Options,
                limits.MaxEditedArchiveBytes - decompressedBytes, token);
            decompressedBytes += raw.Length;
            IWorkPreservingArchiveEdit.Rewrite(raw, path, edits, source.Options, token);
            entries[path] = IWorkSnappy.EncodeIwa(raw, limits.MaxPackageBytes, token);
        }
        byte[] bytes = IWorkNativePackageWriter.WriteEntries(entries, limits.MaxPackageBytes, token);
        // Carrier/encoded entry sizes can grow; retain captured decompression and structural budgets.
        IWorkReadOptions verification = source.Options.Snapshot();
        verification.MaxPackageBytes = Math.Max(verification.MaxPackageBytes, bytes.Length);
        long outputEntryBytes = 0;
        foreach (var entry in entries) {
            token.ThrowIfCancellationRequested(); outputEntryBytes += entry.Value.Length;
            verification.MaxEntryBytes = Math.Max(verification.MaxEntryBytes, entry.Value.Length);
            if (entry.Key.EndsWith(".iwa", StringComparison.OrdinalIgnoreCase))
                verification.MaxIwaBytes = Math.Max(verification.MaxIwaBytes, entry.Value.Length);
        }
        verification.MaxTotalEntryBytes = Math.Max(verification.MaxTotalEntryBytes, outputEntryBytes);
        verification.MaxSnappyChunkBytes = Math.Max(verification.MaxSnappyChunkBytes,
            Math.Min(65_536, verification.MaxDecompressedIwaBytes));
        IWorkSourceDocument reopened = IWorkSourceDocument.Open(bytes, source.Kind, verification, token);
        if (reopened.Records.Count != source.Records.Count || reopened.Entries.Count != source.Entries.Count)
            throw new InvalidDataException("The saved package changed source entry or record membership.");
        var diagnostics = new List<IWorkDiagnostic> {
            new(IWorkDiagnosticSeverity.Information, "IWORK_NATIVE_CARRIER_REWRITTEN",
                "Captured file-entry payloads are preserved. ZIP compression, timestamps, extra fields, comments, directory entries and filesystem metadata are not preserved.")
        };
        if (edits.Count > 0) diagnostics.Add(new(IWorkDiagnosticSeverity.Warning, "IWORK_NATIVE_VISUAL_CACHES_UNASSESSED",
            "Native content was changed while previews and other visual caches were retained without refresh. Native rendering/cache regeneration and wider application metadata are unassessed.",
            lossKind: global::OfficeIMO.OfficeConversionLossKind.Unassessed));
        string sourceHash = source.WithCancellation(token).ComputePackageContentHash(), outputHash = reopened.ComputePackageContentHash();
        if (edits.Count == 0 && sourceHash != outputHash) throw new InvalidDataException("An unchanged save did not preserve normalized source content.");
        var report = new IWorkNativeSaveReport(source.Kind, sourceHash, outputHash, entries.Count, reopened.Records.Count,
            edits.Keys.OrderBy(record => record.EntryPath, StringComparer.Ordinal).ThenBy(record => record.Identifier).Select(record => new IWorkObjectIdentity(record)), paths, diagnostics);
        token.ThrowIfCancellationRequested(); return new(bytes, report);
    }
}
