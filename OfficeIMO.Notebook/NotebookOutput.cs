namespace OfficeIMO.Notebook;

/// <summary>A saved output and its presentation alternatives.</summary>
public sealed class NotebookOutput {
    private static readonly IReadOnlyDictionary<string, JsonElement> EmptyMap =
        new ReadOnlyDictionary<string, JsonElement>(new Dictionary<string, JsonElement>(StringComparer.Ordinal));
    private readonly string? _text;
    private readonly JsonElement _textElement;
    private readonly JsonElement _traceback;
    private IReadOnlyList<string>? _tracebackLines;

    internal NotebookOutput(NotebookOutputKind kind, string originalKind, JsonElement original,
        string? name, JsonElement text, string? literal, string? errorName, string? errorValue, JsonElement traceback,
        IReadOnlyList<string>? tracebackLines, int? executionCount,
        Dictionary<string, JsonElement>? data = null, Dictionary<string, JsonElement>? metadata = null,
        bool legacy = false) {
        Kind = kind;
        OriginalKind = originalKind;
        Original = original;
        Name = name;
        _textElement = text;
        _text = literal;
        ErrorName = errorName;
        ErrorValue = errorValue;
        _traceback = traceback;
        _tracebackLines = tracebackLines;
        ExecutionCount = executionCount;
        Data = data is { Count: > 0 } ? new ReadOnlyDictionary<string, JsonElement>(data) : EmptyMap;
        Metadata = metadata is { Count: > 0 } ? new ReadOnlyDictionary<string, JsonElement>(metadata) : EmptyMap;
        Legacy = legacy;
    }

    /// <summary>Gets the output kind.</summary>
    public NotebookOutputKind Kind { get; }

    /// <summary>Gets the original output-type name, including extension names.</summary>
    public string OriginalKind { get; }

    /// <summary>Gets the name of a stream, such as stdout or stderr.</summary>
    public string? Name { get; }

    /// <summary>Gets stream text with multiline fragments joined as stored.</summary>
    public string Text => _text ?? NotebookJson.ReadText(_textElement);

    /// <summary>Gets the error name.</summary>
    public string? ErrorName { get; }

    /// <summary>Gets the error value.</summary>
    public string? ErrorValue { get; }

    /// <summary>Gets the saved traceback entries in source order.</summary>
    public IReadOnlyList<string> Traceback => _tracebackLines ??= Array.AsReadOnly(
        _traceback.ValueKind == JsonValueKind.Array
            ? _traceback.EnumerateArray().Where(static part => part.ValueKind == JsonValueKind.String).Select(static part => part.GetString()!).ToArray()
            : Array.Empty<string>());

    /// <summary>Gets the execution count associated with an execution result.</summary>
    public int? ExecutionCount { get; }

    /// <summary>Gets all MIME alternatives in their stored JSON forms for the document lifetime.</summary>
    public IReadOnlyDictionary<string, JsonElement> Data { get; }

    /// <summary>Gets output metadata for the document lifetime.</summary>
    public IReadOnlyDictionary<string, JsonElement> Metadata { get; }

    /// <summary>Creates a saved output for a standard stream.</summary>
    public static NotebookOutput CreateStream(string name, string text) {
        if (name is not ("stdout" or "stderr")) throw new ArgumentException("Choose stdout or stderr.", nameof(name));
        ArgumentNullException.ThrowIfNull(text);
        return new NotebookOutput(NotebookOutputKind.Stream, "stream", default, name, default, text,
            null, null, default, null, null);
    }

    /// <summary>Creates a display output from a snapshot of the supplied MIME representations.</summary>
    public static NotebookOutput CreateDisplayData(IReadOnlyDictionary<string, JsonElement> data,
        IReadOnlyDictionary<string, JsonElement>? metadata = null) {
        ArgumentNullException.ThrowIfNull(data);
        return new NotebookOutput(NotebookOutputKind.DisplayData, "display_data", default, null, default, null,
            null, null, default, null, null, NotebookJson.CopyMap(data), NotebookJson.CopyMap(metadata));
    }

    /// <summary>Creates an execution result from saved MIME alternatives and an optional execution count.</summary>
    public static NotebookOutput CreateExecuteResult(IReadOnlyDictionary<string, JsonElement> data, int? executionCount,
        IReadOnlyDictionary<string, JsonElement>? metadata = null) {
        ArgumentNullException.ThrowIfNull(data);
        if (executionCount < 0) throw new ArgumentOutOfRangeException(nameof(executionCount));
        return new NotebookOutput(NotebookOutputKind.ExecuteResult, "execute_result", default, null, default, null,
            null, null, default, null, executionCount, NotebookJson.CopyMap(data), NotebookJson.CopyMap(metadata));
    }

    /// <summary>Creates a saved execution error from a snapshot of its traceback lines.</summary>
    public static NotebookOutput CreateError(string errorName, string errorValue, IEnumerable<string> traceback) {
        ArgumentNullException.ThrowIfNull(errorName);
        ArgumentNullException.ThrowIfNull(errorValue);
        ArgumentNullException.ThrowIfNull(traceback);
        string[] entries = traceback.ToArray();
        if (entries.Any(static line => line == null)) throw new ArgumentException("Traceback lines must be strings.", nameof(traceback));
        return new NotebookOutput(NotebookOutputKind.Error, "error", default, null, default, null,
            errorName, errorValue, default, Array.AsReadOnly(entries), null);
    }

    internal JsonElement Original { get; }
    internal bool Legacy { get; }

    internal string GetText(int maxChars, out bool truncated) => _text != null
        ? NotebookJson.LimitText(_text, maxChars, out truncated)
        : NotebookJson.ReadText(_textElement, maxChars, out truncated);

    internal NotebookOutput Detach() => Original.ValueKind == JsonValueKind.Undefined ? this
        : new NotebookOutput(Kind, OriginalKind, Original.Clone(), Name, default, Text, ErrorName, ErrorValue,
            default, Traceback, ExecutionCount, NotebookJson.CopyMap(Data), NotebookJson.CopyMap(Metadata), Legacy);
}
