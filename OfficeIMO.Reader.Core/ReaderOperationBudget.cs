using System.Runtime.CompilerServices;

namespace OfficeIMO.Reader;

/// <summary>Counts operation work without retaining past streamed chunks or decoded inputs.</summary>
internal sealed class ReaderOperationBudget {
    private readonly ReaderResourceLimits _limits;
    private readonly object _gate = new object();
    private readonly ConditionalWeakTable<object, Entry> _seen = new ConditionalWeakTable<object, Entry>();
    private long _chunks, _characters, _blocks, _assets, _assetBytes, _nestedBytes, _nestedDocuments;
    internal ReaderOperationBudget(ReaderResourceLimits limits) { _limits = limits.CloneValidated(); }

    internal long? NestedInputMaximum => _limits.MaxNestedInputBytes;

    internal long? RemainingNestedBytes { get { lock (_gate) return _limits.MaxNestedInputBytes.HasValue
        ? Math.Max(0, _limits.MaxNestedInputBytes.Value - _nestedBytes) : null; } }

    internal void CheckDepth(int depth) {
        if (_limits.MaxNestedDepth.HasValue && depth > _limits.MaxNestedDepth.Value)
            throw new ReaderResourceLimitException(nameof(ReaderResourceLimits.MaxNestedDepth), _limits.MaxNestedDepth.Value);
    }

    internal void AddChunk(ReaderChunk chunk) {
        lock (_gate) {
            long characters = (long)(chunk.Text?.Length ?? 0) + (chunk.Markdown?.Length ?? 0);
            if (_seen.TryGetValue(chunk, out var prior)) {
                Add(ref _characters, characters - prior.Characters, _limits.MaxChunkCharacters, nameof(ReaderResourceLimits.MaxChunkCharacters));
                prior.Characters = characters;
                return;
            }
            Add(ref _chunks, 1, _limits.MaxChunks, nameof(ReaderResourceLimits.MaxChunks));
            Add(ref _characters, characters, _limits.MaxChunkCharacters, nameof(ReaderResourceLimits.MaxChunkCharacters));
            _seen.Add(chunk, new Entry { Characters = characters });
        }
    }

    internal void AddDocument(OfficeDocumentReadResult document) => AddDocument(document, new HashSet<OfficeDocumentReadResult>());

    private void AddDocument(OfficeDocumentReadResult document, HashSet<OfficeDocumentReadResult> ancestors) {
        if (!ancestors.Add(document)) throw new InvalidOperationException("Nested document results cannot contain a cycle.");
        foreach (var nested in document.NestedDocuments ?? Array.Empty<OfficeDocumentNestedResult>()) AddDocument(nested.Document, ancestors);
        foreach (var chunk in document.Chunks ?? Array.Empty<ReaderChunk>()) AddChunk(chunk);
        lock (_gate) {
            foreach (var block in (document.Blocks ?? Array.Empty<OfficeDocumentBlock>()).Concat(
                         (document.Pages ?? Array.Empty<OfficeDocumentPage>()).SelectMany(page => page.Blocks ?? Array.Empty<OfficeDocumentBlock>()))) {
                if (_seen.TryGetValue(block, out _)) continue;
                Add(ref _blocks, 1, _limits.MaxBlocks, nameof(ReaderResourceLimits.MaxBlocks));
                _seen.Add(block, new Entry());
            }
            foreach (var asset in (document.Assets ?? Array.Empty<OfficeDocumentAsset>()).Concat(
                         (document.Pages ?? Array.Empty<OfficeDocumentPage>()).SelectMany(page => page.Assets ?? Array.Empty<OfficeDocumentAsset>()))) {
                var assetEntry = _seen.GetValue(asset, _ => new Entry());
                if (!assetEntry.AssetRecord) {
                    Add(ref _assets, 1, _limits.MaxAssets, nameof(ReaderResourceLimits.MaxAssets));
                    assetEntry.AssetRecord = true;
                }
                long bytes = asset.PayloadBytes?.LongLength ?? Math.Max(0, asset.LengthBytes ?? 0);
                if (asset.PayloadBytes == null) {
                    long representedBytes = Math.Max(0, bytes - assetEntry.MaterializedAssetBytes);
                    Add(ref _assetBytes, Math.Max(0, representedBytes - assetEntry.KnownAssetBytes), _limits.MaxAssetBytes, nameof(ReaderResourceLimits.MaxAssetBytes));
                    assetEntry.KnownAssetBytes = Math.Max(assetEntry.KnownAssetBytes, representedBytes);
                } else {
                    // A descriptor becoming a payload represents the same bytes, not two allocations.
                    Add(ref _assetBytes, -assetEntry.KnownAssetBytes, _limits.MaxAssetBytes, nameof(ReaderResourceLimits.MaxAssetBytes));
                    assetEntry.KnownAssetBytes = 0;
                    assetEntry.MaterializedAssetBytes = Math.Max(assetEntry.MaterializedAssetBytes, bytes);
                    var payloadEntry = _seen.GetValue(asset.PayloadBytes, _ => new Entry());
                    if (!payloadEntry.AssetPayload) {
                        Add(ref _assetBytes, bytes, _limits.MaxAssetBytes, nameof(ReaderResourceLimits.MaxAssetBytes));
                        payloadEntry.AssetPayload = true;
                    }
                }
            }
        }
        ancestors.Remove(document);
    }

    internal void ReserveNestedInput(long bytes) {
        lock (_gate) {
            Add(ref _nestedDocuments, 1, _limits.MaxNestedDocuments, nameof(ReaderResourceLimits.MaxNestedDocuments));
            ReserveNestedBytes(bytes);
        }
    }

    internal void ReserveNestedBytes(long bytes) {
        lock (_gate) Add(ref _nestedBytes, bytes, _limits.MaxNestedInputBytes, nameof(ReaderResourceLimits.MaxNestedInputBytes));
    }

    internal void MarkReservedPayload(byte[] payload) {
        lock (_gate) _seen.GetValue(payload, _ => new Entry()).ReservedInput = true;
    }

    internal bool IsReserved(Stream stream) {
        if (stream is not MemoryStream memory || !memory.TryGetBuffer(out var buffer) || buffer.Array == null) return false;
        lock (_gate) return _seen.TryGetValue(buffer.Array, out var entry) && entry.ReservedInput;
    }

    private static void Add(ref long current, long amount, long? maximum, string name) {
        if (amount > 0 && maximum.HasValue && amount > maximum.Value - current)
            throw new ReaderResourceLimitException(name, maximum.Value);
        current = checked(current + amount);
    }
    private sealed class Entry { internal long Characters; internal bool ReservedInput; internal bool AssetPayload; internal bool AssetRecord; internal long KnownAssetBytes; internal long MaterializedAssetBytes; }
}
