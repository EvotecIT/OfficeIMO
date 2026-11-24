namespace OfficeIMO.Notebook;

/// <summary>An editable, offline Jupyter notebook and its saved outputs.</summary>
/// <remarks>Dispose the document after using its source-backed JSON values. Editing and saving the same document concurrently is unsupported.</remarks>
public sealed partial class NotebookDocument : IDisposable {
    private const int CurrentMinor = 5;
    private readonly List<NotebookCell> _cells = new();
    private readonly Dictionary<string, JsonElement> _metadata;
    private readonly JsonDocument? _json;
    private readonly string? _path;
    private bool _disposed;
    private bool _metadataEdited;

    private NotebookDocument(string? path, JsonDocument? json, int? major, int? minor) {
        _path = path;
        _json = json;
        DeclaredNbFormat = major;
        DeclaredNbFormatMinor = minor;
        NbFormat = major ?? 4;
        NbFormatMinor = minor ?? (json == null ? CurrentMinor : 0);
        _metadata = NotebookJson.ReadMap(NotebookJson.Property(Root, "metadata"));
        Metadata = new ReadOnlyDictionary<string, JsonElement>(_metadata);
        Cells = _cells.AsReadOnly();
    }

    /// <summary>Gets the source major version, or 4 for a new notebook.</summary>
    public int NbFormat { get; }

    /// <summary>Gets the source minor version, or the current writer version for a new notebook.</summary>
    public int NbFormatMinor { get; }

    /// <summary>Gets ordered notebook cells. Legacy worksheets contribute cells in worksheet order.</summary>
    public IReadOnlyList<NotebookCell> Cells { get; }

    /// <summary>Gets notebook metadata. Source values are valid for the document lifetime.</summary>
    public IReadOnlyDictionary<string, JsonElement> Metadata { get; }

    /// <summary>Gets the notebook language from kernel or language metadata, when present.</summary>
    public string? Language {
        get {
            if (Metadata.TryGetValue("kernelspec", out JsonElement kernel)) {
                string? kernelLanguage = NotebookJson.String(kernel, "language");
                if (!string.IsNullOrWhiteSpace(kernelLanguage)) return kernelLanguage;
            }
            return Metadata.TryGetValue("language_info", out JsonElement language) ? NotebookJson.String(language, "name") : null;
        }
    }

    /// <summary>Gets source findings, including conversions applied during loading.</summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics { get; private set; } = Array.Empty<OfficeDiagnostic>();

    /// <summary>Creates an in-memory nbformat 4.5 notebook.</summary>
    public static NotebookDocument Create() => new(null, null, 4, CurrentMinor);

    /// <summary>Creates an nbformat 4.5 notebook whose default save destination is the given path.</summary>
    public static NotebookDocument Create(string path) => new(GetPath(path), null, 4, CurrentMinor);

    /// <summary>Appends Markdown to the notebook and returns the new cell.</summary>
    public NotebookCell AddMarkdownCell(string source) => AddCell(NotebookCellKind.Markdown, "markdown", source);

    /// <summary>Appends code to the notebook and returns the new cell.</summary>
    public NotebookCell AddCodeCell(string source) => AddCell(NotebookCellKind.Code, "code", source);

    /// <summary>Appends a raw cell to the notebook and returns it.</summary>
    public NotebookCell AddRawCell(string source) => AddCell(NotebookCellKind.Raw, "raw", source);

    /// <summary>Sets a notebook metadata value. The document takes an independent snapshot of the value.</summary>
    public void SetMetadata(string name, JsonElement value) {
        CheckDisposed();
        ArgumentNullException.ThrowIfNull(name);
        _metadata[name] = NotebookJson.CloneValue(value);
        _metadataEdited = true;
    }

    /// <summary>Releases the parsed JSON backing store and invalidates source-backed values.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _json?.Dispose();
    }

    internal int? DeclaredNbFormat { get; }
    internal int? DeclaredNbFormatMinor { get; }
    internal int SourceCellCount { get; private set; }

    internal NotebookCell AddImportedCell(NotebookCellKind kind, string source, string? language, string id) {
        string name = kind switch {
            NotebookCellKind.Code => "code",
            NotebookCellKind.Markdown => "markdown",
            NotebookCellKind.Raw => "raw",
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
        return AddCell(kind, name, source, id, language);
    }

    internal void ApplyImportReport(OfficeConversionReport report) {
        CheckDisposed();
        Diagnostics = report.Diagnostics;
    }

    private JsonElement Root => _json?.RootElement ?? default;

    private void CheckDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(NotebookDocument));
    }

    private static string GetPath(string path) {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        return Path.GetFullPath(path);
    }

    private NotebookCell AddCell(NotebookCellKind kind, string name, string source, string? id = null, string? language = null) {
        CheckDisposed();
        ArgumentNullException.ThrowIfNull(source);
        var metadata = new Dictionary<string, JsonElement>(StringComparer.Ordinal);
        if (language != null) metadata.Add("language", NotebookJson.StringValue(language));
        var cell = new NotebookCell(kind, name, id ?? Guid.NewGuid().ToString("N"), SourceCellCount++, default, default,
            source, language, null, metadata, new List<NotebookOutput>());
        _cells.Add(cell);
        return cell;
    }
}
