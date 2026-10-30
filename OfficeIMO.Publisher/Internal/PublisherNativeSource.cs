using OfficeIMO.Core.Internal;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Owned native bytes and logical storage. Retention does not establish native edit or repair-free opening qualification.</summary>
internal sealed class PublisherNativeSource {
    private readonly byte[] _bytes;
    private readonly OfficeCompoundFile _compound;
    internal PublisherNativeSource(byte[] bytes, OfficeCompoundFile compound) { _bytes = (byte[])bytes.Clone(); _compound = compound; }
    internal int ByteLength => _bytes.Length;
    internal int StreamCount => _compound.Streams.Count;
    internal byte[] CopyBytes() => (byte[])_bytes.Clone();
    internal byte[] EscherBytes => _compound.Streams["Escher/EscherStm"];
    internal byte[] QuillBytes => _compound.Streams["Quill/QuillSub/CONTENTS"];
    internal byte[] ContentsBytes => _compound.Streams["Contents"];

    internal bool CanRewrite(OfficeLegacyImportLimits limits, CancellationToken token) {
        var bounds = new OfficeCompoundReadOptions(maxDirectoryEntries: limits.MaxRecords,
            maxStreamCount: limits.MaxCompoundStreams, maxStreamBytes: limits.MaxInputBytes,
            maxTotalStreamBytes: limits.MaxInputBytes);
        try {
            using var input = new MemoryStream(_bytes, writable: false);
            using var source = new OfficeCompoundFileReader.ReadSession(input, bounds, leaveOpen: true, token);
            // Compatibility aliases are emitted even for a valid reachable tree.
            // Only the shared structural index can distinguish them from lost entries.
            if (!source.HasCanonicalRootDirectory || source.HasInvalidDirectoryTree) return false;
            return !source.Entries.Any(entry => entry.IsStream && !entry.IsFallback && entry.Size > 0
                && (entry.Path.StartsWith("VBA/", StringComparison.OrdinalIgnoreCase)
                    || entry.Path.StartsWith("Objects/", StringComparison.OrdinalIgnoreCase)
                    || entry.Path.IndexOf("signature", StringComparison.OrdinalIgnoreCase) >= 0
                    || entry.Path.IndexOf("encrypt", StringComparison.OrdinalIgnoreCase) >= 0
                    || entry.Path.StartsWith("\u0006DataSpaces", StringComparison.OrdinalIgnoreCase)));
        } catch (InvalidDataException) {
            return false;
        }
    }

    /// <summary>Removes the stale optional SummaryInformation thumbnail from edited output, retaining other raw property values through Core.</summary>
    internal bool InvalidatePreview(IDictionary<string, byte[]> replacements, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        string name = OfficeOlePropertySetWriter.SummaryInformationStreamName;
        if (!_compound.Streams.TryGetValue(name, out byte[]? source)) return false;
        if (!OfficeOlePropertySetReader.ReadSections(source, token).Any(section =>
            section.FormatId == OfficeOlePropertySetWriter.SummaryInformationFormatId && section.Properties.ContainsKey(17))) return false;
        // A null replacement removes the complete property, including opaque
        // VT_CF clipboard data; it does not reinterpret or render the preview.
        replacements.Add(name, OfficeOlePropertySetEditor.RewriteStrings(source,
            OfficeOlePropertySetWriter.SummaryInformationFormatId, new Dictionary<uint, string?> { [17] = null }, token));
        return true;
    }

    internal byte[] Rewrite(IReadOnlyDictionary<string, byte[]> streams, long maximumBytes, CancellationToken token) =>
        OfficeCompoundFileWriter.Rewrite(_compound, streams,
            maxOutputBytes: maximumBytes, cancellationToken: token);
}
