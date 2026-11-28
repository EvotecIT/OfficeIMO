namespace OfficeIMO.Notebook;

public sealed partial class NotebookDocument {
    /// <summary>Loads an nbformat 4 or legacy nbformat 3 notebook from a file.</summary>
    public static NotebookDocument Load(string path, NotebookLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        string sourcePath = GetPath(path);
        cancellationToken.ThrowIfCancellationRequested();
        using FileStream source = File.OpenRead(sourcePath);
        return LoadStream(source, sourcePath, (options ?? new()).Snapshot(), cancellationToken);
    }

    /// <summary>Loads a notebook from the current position of a readable stream, leaving it open.</summary>
    public static NotebookDocument Load(Stream stream, NotebookLoadOptions? options = null,
        CancellationToken cancellationToken = default) =>
        LoadStream(stream, null, (options ?? new()).Snapshot(), cancellationToken);

    /// <summary>Loads an independent snapshot of UTF-8 notebook bytes.</summary>
    public static NotebookDocument Load(ReadOnlyMemory<byte> bytes, NotebookLoadOptions? options = null,
        CancellationToken cancellationToken = default) =>
        LoadBytes(bytes, (options ?? new()).Snapshot(), takeOwnership: false, cancellationToken);

    /// <summary>Loads a notebook asynchronously from a file.</summary>
    public static async Task<NotebookDocument> LoadAsync(string path, NotebookLoadOptions? options = null,
        CancellationToken cancellationToken = default) {
        string sourcePath = GetPath(path);
        cancellationToken.ThrowIfCancellationRequested();
        using var source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read,
            bufferSize: 4096, useAsync: true);
        return await LoadStreamAsync(source, sourcePath, (options ?? new()).Snapshot(), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Loads a notebook asynchronously from the current stream position, leaving the stream open.</summary>
    public static Task<NotebookDocument> LoadAsync(Stream stream, NotebookLoadOptions? options = null,
        CancellationToken cancellationToken = default) =>
        LoadStreamAsync(stream, null, (options ?? new()).Snapshot(), cancellationToken);

    // The Reader owns the immutable bytes and projects only the inspected cells and outputs.
    internal static NotebookDocument LoadForReader(byte[] bytes, int maxCells, int maxOutputsPerCell,
        CancellationToken cancellationToken) {
        var options = new NotebookLoadOptions { MaxInputBytes = Math.Max(1L, bytes.LongLength) };
        return LoadBytes(bytes, options, takeOwnership: true, cancellationToken, (maxCells, maxOutputsPerCell));
    }

    private static NotebookDocument LoadStream(Stream stream, string? path, NotebookLoadOptions options,
        CancellationToken cancellationToken) {
        using var bounded = new NotebookInputStream(stream, options.MaxInputBytes, cancellationToken);
        return FinishLoad(Parse(() => JsonDocument.Parse(bounded, JsonOptions(options))), path, options, cancellationToken);
    }

    private static async Task<NotebookDocument> LoadStreamAsync(Stream stream, string? path, NotebookLoadOptions options,
        CancellationToken cancellationToken) {
        using var bounded = new NotebookInputStream(stream, options.MaxInputBytes, cancellationToken);
        JsonDocument json;
        try {
            json = await JsonDocument.ParseAsync(bounded, JsonOptions(options), cancellationToken).ConfigureAwait(false);
        } catch (JsonException error) {
            throw InvalidJson(error);
        }
        return FinishLoad(json, path, options, cancellationToken);
    }

    private static NotebookDocument LoadBytes(ReadOnlyMemory<byte> bytes, NotebookLoadOptions options,
        bool takeOwnership, CancellationToken cancellationToken, (int MaxCells, int MaxOutputsPerCell)? projection = null) {
        cancellationToken.ThrowIfCancellationRequested();
        if (bytes.Length > options.MaxInputBytes) throw NotebookInputStream.InputLimit(options.MaxInputBytes);
        ReadOnlySpan<byte> span = bytes.Span;
        if (span.Length >= 3 && span[0] == 0xEF && span[1] == 0xBB && span[2] == 0xBF) {
            bytes = bytes.Slice(3);
        }
        if (!takeOwnership) bytes = bytes.ToArray();
        return FinishLoad(Parse(() => JsonDocument.Parse(bytes, JsonOptions(options))), null, options, cancellationToken, projection);
    }

    private static JsonDocumentOptions JsonOptions(NotebookLoadOptions options) =>
        new() { MaxDepth = options.MaxJsonDepth };

    private static JsonDocument Parse(Func<JsonDocument> parse) {
        try {
            return parse();
        } catch (JsonException error) {
            throw InvalidJson(error);
        }
    }

    private static FormatException InvalidJson(JsonException error) =>
        new("Notebook input is not valid UTF-8 JSON within the configured depth.", error);

    private static NotebookDocument FinishLoad(JsonDocument json, string? path, NotebookLoadOptions options,
        CancellationToken cancellationToken, (int MaxCells, int MaxOutputsPerCell)? projection = null) {
        try {
            cancellationToken.ThrowIfCancellationRequested();
            JsonElement root = json.RootElement;
            if (root.ValueKind != JsonValueKind.Object) throw new FormatException("Notebook JSON must be an object.");
            int? major = ReadVersion(root, "nbformat");
            int? minor = ReadVersion(root, "nbformat_minor");
            if (major is not (null or 3 or 4)) throw new NotSupportedException($"Notebook version {major} is not supported.");
            var document = new NotebookDocument(path, json, major, minor);
            var diagnostics = new List<OfficeDiagnostic>();
            List<JsonElement> sourceCells = NotebookParser.ReadSourceCells(root, document.NbFormat, projection?.MaxCells ?? options.MaxCells,
                projection.HasValue, cancellationToken, diagnostics, out int sourceCellCount);
            document.SourceCellCount = sourceCellCount;
            document._cells.AddRange(NotebookParser.ReadCells(sourceCells, document.NbFormat, document.NbFormatMinor,
                options, projection?.MaxOutputsPerCell, cancellationToken, diagnostics));
            document.Diagnostics = diagnostics.AsReadOnly();
            return document;
        } catch {
            json.Dispose();
            throw;
        }
    }

    private static int? ReadVersion(JsonElement root, string name) {
        if (!root.TryGetProperty(name, out JsonElement value)) return null;
        if (value.ValueKind != JsonValueKind.Number || !value.TryGetInt32(out int version) || version < 0) {
            throw new FormatException($"Notebook {name} must be a nonnegative integer.");
        }
        return version;
    }
}
