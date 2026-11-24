namespace OfficeIMO.Notebook;

/// <summary>An ordered cell in a Jupyter notebook.</summary>
public sealed class NotebookCell {
    private readonly Dictionary<string, JsonElement> _metadata;
    private readonly List<NotebookOutput> _outputs;
    private readonly Dictionary<string, IReadOnlyDictionary<string, JsonElement>> _attachments;
    private readonly string? _sourceLanguage;
    private string? _source;
    private string? _language;
    private int? _executionCount;

    internal NotebookCell(NotebookCellKind kind, string originalKind, string id, int sourceIndex,
        JsonElement original, JsonElement source, string? normalizedSource, int? executionCount,
        Dictionary<string, JsonElement> metadata, List<NotebookOutput> outputs) {
        Kind = kind;
        OriginalKind = originalKind;
        Id = id;
        SourceIndex = sourceIndex;
        Original = original;
        SourceElement = source;
        _source = normalizedSource;
        _executionCount = executionCount;
        _metadata = metadata;
        _sourceLanguage = NotebookJson.String(original, "language");
        _language = ResolveLanguage();
        _outputs = outputs;
        _attachments = ReadAttachments(NotebookJson.Property(original, "attachments"));
        Metadata = new ReadOnlyDictionary<string, JsonElement>(_metadata);
        Outputs = _outputs.AsReadOnly();
        Attachments = new ReadOnlyDictionary<string, IReadOnlyDictionary<string, JsonElement>>(_attachments);
    }

    /// <summary>Gets the notebook cell kind.</summary>
    public NotebookCellKind Kind { get; }

    /// <summary>Gets the source cell-type name, including extension names.</summary>
    public string OriginalKind { get; }

    /// <summary>Gets the stable cell identifier.</summary>
    public string Id { get; private set; }

    /// <summary>Gets or sets the cell source, with multiline JSON fragments joined as stored.</summary>
    public string Source {
        get => _source ??= NotebookJson.ReadText(SourceElement);
        set {
            _source = value ?? throw new ArgumentNullException(nameof(value));
            SourceEdited = true;
        }
    }

    /// <summary>Gets the effective cell-specific language from the cell's source and current metadata.</summary>
    public string? Language => _language;

    /// <summary>Gets or sets the saved execution count of a code cell. Set it to <see langword="null"/> to clear the count.</summary>
    public int? ExecutionCount {
        get => _executionCount;
        set {
            if (Kind != NotebookCellKind.Code) throw new InvalidOperationException("Only code cells can have an execution count.");
            if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));
            _executionCount = value;
            ExecutionCountEdited = true;
        }
    }

    /// <summary>Gets the cell metadata. Values from an input are valid for the document lifetime.</summary>
    public IReadOnlyDictionary<string, JsonElement> Metadata { get; }

    /// <summary>Gets the cell outputs in document order.</summary>
    public IReadOnlyList<NotebookOutput> Outputs { get; }

    /// <summary>Gets attachments keyed by their notebook-local names and MIME alternatives.</summary>
    public IReadOnlyDictionary<string, IReadOnlyDictionary<string, JsonElement>> Attachments { get; }

    /// <summary>Sets a cell metadata value. The cell takes an independent snapshot of the value.</summary>
    public void SetMetadata(string name, JsonElement value) {
        ArgumentNullException.ThrowIfNull(name);
        _metadata[name] = NotebookJson.CloneValue(value);
        MetadataEdited = true;
        if (name is "language" or "polyglot_notebook" or "dotnet_interactive" or "vscode") {
            _language = ResolveLanguage();
        }
    }

    /// <summary>Appends a saved output to this code cell.</summary>
    public void AddOutput(NotebookOutput output) {
        ArgumentNullException.ThrowIfNull(output);
        if (Kind != NotebookCellKind.Code) throw new InvalidOperationException("Only code cells can contain outputs.");
        _outputs.Add(output.Detach());
        OutputsEdited = true;
    }

    /// <summary>Clears all saved outputs of a code cell and retains its execution count.</summary>
    public void ClearOutputs() {
        if (Kind != NotebookCellKind.Code) throw new InvalidOperationException("Only code cells can contain outputs.");
        _outputs.Clear();
        OutputsEdited = true;
    }

    /// <summary>Sets a notebook-local attachment on a Markdown or raw cell.</summary>
    public void SetAttachment(string name, IReadOnlyDictionary<string, JsonElement> data) {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentNullException.ThrowIfNull(data);
        if (Kind is not (NotebookCellKind.Markdown or NotebookCellKind.Raw)) {
            throw new InvalidOperationException("Only Markdown and raw cells can contain attachments.");
        }
        _attachments[name] = new ReadOnlyDictionary<string, JsonElement>(NotebookJson.CopyMap(data));
        AttachmentsEdited = true;
    }

    internal void SetImportedId(string id) => Id = id;

    internal int SourceIndex { get; }
    internal int SourceOutputCount => NotebookJson.Property(Original, "outputs") is { ValueKind: JsonValueKind.Array } outputs
        ? outputs.GetArrayLength() : 0;
    internal JsonElement Original { get; }
    internal JsonElement SourceElement { get; }
    internal bool SourceEdited { get; private set; }
    internal bool MetadataEdited { get; private set; }
    internal bool ExecutionCountEdited { get; private set; }
    internal bool OutputsEdited { get; private set; }
    internal bool AttachmentsEdited { get; private set; }

    internal string GetSource(int maxChars, out bool truncated) => _source != null
        ? NotebookJson.LimitText(_source, maxChars, out truncated)
        : NotebookJson.ReadText(SourceElement, maxChars, out truncated);

    private string? ResolveLanguage() {
        string? language = _sourceLanguage ??
            (_metadata.TryGetValue("language", out JsonElement value) && value.ValueKind == JsonValueKind.String
                ? value.GetString() : null) ??
            ReadVendorLanguage("polyglot_notebook", "kernelName") ??
            ReadVendorLanguage("dotnet_interactive", "language");
        if (!string.IsNullOrWhiteSpace(language)) return language;
        string? editorLanguage = ReadVendorLanguage("vscode", "languageId");
        return editorLanguage == "polyglot-notebook" ? null : editorLanguage;
    }

    private string? ReadVendorLanguage(string vendor, string key) =>
        _metadata.TryGetValue(vendor, out JsonElement metadata) ? NotebookJson.String(metadata, key) : null;

    private static Dictionary<string, IReadOnlyDictionary<string, JsonElement>> ReadAttachments(JsonElement value) {
        var attachments = new Dictionary<string, IReadOnlyDictionary<string, JsonElement>>(StringComparer.Ordinal);
        if (value.ValueKind == JsonValueKind.Object) {
            foreach (JsonProperty property in value.EnumerateObject()) {
                attachments[property.Name] = new ReadOnlyDictionary<string, JsonElement>(NotebookJson.ReadMap(property.Value));
            }
        }
        return attachments;
    }
}
