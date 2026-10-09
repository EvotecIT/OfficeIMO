using System.Security.Cryptography;

namespace OfficeIMO.DjVu;

/// <summary>An owned, immutable snapshot of a DjVu document and its explicitly resolved components.</summary>
public sealed class DjVuDocument {
    internal readonly DjVuReadOptions ReadOptions;
    internal readonly Dictionary<string, DjVuComponent> Components;
    private DjVuDocument(byte[] source, DjVuReadOptions options, CancellationToken cancellation) {
        ReadOptions = options;
        SourceLengthBytes = source.LongLength;
        var budget = new DjVuReadBudget(options, cancellation);
        var container = new DjVuContainerReader(budget);
        var components = container.Read(source);
        Components = components.ToDictionary(c => c.Id, StringComparer.Ordinal);
        ValidateIncludes(budget);
        var pages = new List<DjVuPage>();
        foreach (var component in components) {
            if (component.Kind != 1) continue;
            pages.Add(new DjVuPage(this, component, pages.Count + 1, budget));
        }
        Pages = pages.AsReadOnly();
        Bookmarks = DjVuBookmarkReader.Read(container.Root, this, budget, out string? navigationDiagnostic);
        BookmarkDiagnostic = navigationDiagnostic;
        using var algorithm = SHA256.Create();
        SourceSha256 = BitConverter.ToString(algorithm.ComputeHash(source)).Replace("-", string.Empty).ToLowerInvariant();
    }

    /// <summary>Pages in the source directory's order.</summary>
    public IReadOnlyList<DjVuPage> Pages { get; }
    /// <summary>Stored document outline entries in source order.</summary>
    public IReadOnlyList<DjVuBookmark> Bookmarks { get; }
    /// <summary>A malformed outline diagnostic; null for absent or valid outlines.</summary>
    public string? BookmarkDiagnostic { get; }
    /// <summary>SHA-256 of the exact primary input snapshot. Resolved indirect components are separate inputs.</summary>
    public string SourceSha256 { get; }
    /// <summary>Length of the exact primary input snapshot, excluding explicitly resolved components.</summary>
    public long SourceLengthBytes { get; }

    /// <summary>Loads and copies a DjVu file's bytes, applying aggregate document limits.</summary>
    public static DjVuDocument Load(byte[] bytes, DjVuReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        var snapshot = (options ?? new DjVuReadOptions()).Snapshot();
        cancellationToken.ThrowIfCancellationRequested();
        CheckSource(bytes.LongLength, snapshot);
        return new DjVuDocument((byte[])bytes.Clone(), snapshot, cancellationToken);
    }

    /// <summary>Loads from the current position of a readable stream, including non-seekable streams. Leaves the stream open.</summary>
    public static DjVuDocument Load(Stream stream, DjVuReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("DjVu input stream must be readable.", nameof(stream));
        var snapshot = (options ?? new DjVuReadOptions()).Snapshot();
        using var buffer = new MemoryStream();
        byte[] transfer = new byte[81920];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = stream.Read(transfer, 0, (int)Math.Min(transfer.Length, snapshot.MaxSourceBytes - buffer.Length + 1));
            if (count == 0) break;
            CheckSource(buffer.Length + count, snapshot);
            buffer.Write(transfer, 0, count);
        }
        return new DjVuDocument(buffer.ToArray(), snapshot, cancellationToken);
    }

    /// <summary>Loads a file and closes its handle before returning. Component IDs never trigger implicit file access.</summary>
    public static DjVuDocument Load(string path, DjVuReadOptions? options = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        using var stream = File.OpenRead(path);
        return Load(stream, options, cancellationToken);
    }

    private static void CheckSource(long length, DjVuReadOptions options) {
        if (length > options.MaxSourceBytes) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxSourceBytes));
    }

    internal IEnumerable<DjVuChunk> PageChunks(DjVuComponent component, CancellationToken cancellation) {
        int visits = 0;
        var includedComponents = new HashSet<string>(StringComparer.Ordinal);
        return Enumerate(component, 1);
        IEnumerable<DjVuChunk> Enumerate(DjVuComponent current, int depth) {
            cancellation.ThrowIfCancellationRequested();
            if (depth > ReadOptions.MaxDepth) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxDepth));
            if (!includedComponents.Add(current.Id)) yield break;
            foreach (var chunk in current.Form.Children) {
                cancellation.ThrowIfCancellationRequested();
                if (++visits > ReadOptions.MaxChunks) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxChunks));
                if (chunk.Id == "INCL") {
                    string id = DjVuBinary.Utf8.GetString(chunk.Source, chunk.Offset, chunk.Length);
                    foreach (var included in Enumerate(Components[id], depth + 1)) yield return included;
                } else yield return chunk;
            }
        }
    }

    private void ValidateIncludes(DjVuReadBudget budget) {
        var state = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var component in Components.Values) Visit(component, 1);
        void Visit(DjVuComponent component, int depth) {
            budget.Cancellation.ThrowIfCancellationRequested();
            if (depth > ReadOptions.MaxDepth) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxDepth));
            if (state.TryGetValue(component.Id, out int current)) {
                if (current == 1) throw new InvalidDataException("Cyclic DjVu component inclusion.");
                return;
            }
            state[component.Id] = 1;
            foreach (var chunk in component.Form.Children.Where(c => c.Id == "INCL")) {
                string id = DjVuBinary.Utf8.GetString(chunk.Source, chunk.Offset, chunk.Length);
                if (!Components.TryGetValue(id, out var included) || included.Kind != 0) throw new InvalidDataException("DjVu INCL does not identify a shared component.");
                Visit(included, depth + 1);
            }
            state[component.Id] = 2;
        }
    }
}
