using OfficeIMO.Core.Internal;
using OfficeIMO.Provenance;
using System.Threading;

namespace OfficeIMO.Epub;

internal static partial class EpubReader {
    /// <summary>Reuses package discovery and archive identity checks for loss-aware editing.</summary>
    internal static (string OpfPath, Dictionary<string, byte[]> Entries, IReadOnlyList<EpubEncryptionInfo> Encryption)
        ReadEditablePackage(byte[] bytes, EpubPublicationLoadOptions options, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        OfficeProvenanceZip.ValidateMimetypeEntry(bytes, "application/epub+zip", options.MaxEntries);
        token.ThrowIfCancellationRequested();
        var readOptions = new EpubReadOptions {
            MaxPackageBytes = options.MaxInputBytes, MaxArchiveEntries = options.MaxEntries,
            MaxTotalUncompressedBytes = options.MaxExpandedBytes, MaxPackageMetadataBytes = options.MaxMetadataBytes
        };
        using var input = new MemoryStream(bytes, false);
        using var archive = new ZipArchive(input, ZipArchiveMode.Read);
        OfficeProvenanceZip.GetValidatedEntryNames(bytes, archive);
        var diagnostics = new EpubDiagnosticCollector();
        Dictionary<string, ZipArchiveEntry> index = BuildEntryIndex(archive, readOptions, diagnostics, token);
        if (diagnostics.Items.Count != 0) throw new InvalidDataException("Editing requires unique, safe archive entry paths.");
        foreach (var pair in index) {
            if (pair.Value.FullName != pair.Key) throw new InvalidDataException("Editing requires canonical archive paths.");
        }
        EpubPackage? package = TryReadPackage(index, readOptions, diagnostics, out IReadOnlyList<EpubRootfile> rootfiles, token);
        if (package == null || !rootfiles.Any(rootfile => rootfile.IsSelected)) {
            throw new InvalidDataException("Editing requires a readable, declared package document.");
        }
        var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        long retained = 0;
        foreach (var pair in index) {
            token.ThrowIfCancellationRequested();
            if (pair.Value.Length > options.MaxEntryBytes) throw new InvalidDataException("Entry exceeds MaxEntryBytes: " + pair.Key);
            using Stream entry = pair.Value.Open();
            byte[] payload = OfficeStreamReader.ReadAllBytes(entry, token, Math.Max(1, Math.Min(options.MaxEntryBytes, options.MaxExpandedBytes - retained)));
            if (payload.LongLength > options.MaxExpandedBytes - retained) throw new InvalidDataException("Package exceeds MaxExpandedBytes.");
            retained += payload.LongLength;
            entries.Add(pair.Key, payload);
        }
        return (package.OpfPath, entries, ReadEncryption(index, readOptions, diagnostics, token));
    }
}
