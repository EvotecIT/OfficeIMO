namespace OfficeIMO.Pdf;

/// <summary>Transactional existing-document AcroForm edit commands.</summary>
public sealed class PdfAcroFormEditSession {
    private readonly List<EditCommand> _commands = new List<EditCommand>();
    private readonly Dictionary<string, JavaScriptContribution> _createdJavaScripts = new Dictionary<string, JavaScriptContribution>(StringComparer.Ordinal);
    private readonly PdfReadLimits _limits;
    /// <summary>Creates a standalone edit session. Documents normally create sessions through <see cref="PdfDocumentForms.Edit(Action{PdfAcroFormEditSession})"/>.</summary>
    public PdfAcroFormEditSession() : this(new PdfReadLimits()) { }
    internal PdfAcroFormEditSession(PdfReadLimits limits) {
        _limits = limits;
    }
    /// <summary>Creates a text, checkbox, choice, radio-button, push-button, or empty signature field.</summary>
    public PdfAcroFormEditSession Create(PdfFormFieldCreateOptions options) {
        Guard.NotNull(options, nameof(options));
        PdfFormFieldCreateOptions snapshot = options.Snapshot();
        byte[]? encodedJavaScript = ValidateJavaScript(snapshot.JavaScript);
        TrackCreatedJavaScript(snapshot, encodedJavaScript);
        _commands.Add(new EditCommand(EditKind.Create, options: snapshot, encodedJavaScript: encodedJavaScript));
        return this;
    }
    /// <summary>Places an empty signature field owned by the signature engine.</summary>
    public PdfAcroFormEditSession PlaceSignatureField(string name, int pageNumber, double x, double y, double width, double height) => Create(new PdfFormFieldCreateOptions { Name = name, Kind = PdfFormFieldCreationKind.Signature, PageNumber = pageNumber, X = x, Y = y, Width = width, Height = height });
    /// <summary>Renames one fully qualified field.</summary>
    public PdfAcroFormEditSession Rename(string name, string newName) { AddNames(EditKind.Rename, name, newName); RenameCreatedJavaScript(name, newName); return this; }
    /// <summary>Removes one field subtree and its widgets.</summary>
    public PdfAcroFormEditSession Remove(string name) { AddName(EditKind.Remove, name); RemoveCreatedJavaScript(name); return this; }
    /// <summary>Moves a single-widget field to a page rectangle.</summary>
    public PdfAcroFormEditSession Move(string name, int pageNumber, double x, double y, double width, double height) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        ValidateRectangle(x, y, width, height);
        _commands.Add(new EditCommand(EditKind.Move, name, pageNumber: pageNumber, rectangle: new[] { x, y, x + width, y + height }));
        return this;
    }
    /// <summary>Sets or clears a field default value.</summary>
    public PdfAcroFormEditSession SetDefaultValue(string name, string? value) { Guard.NotNullOrWhiteSpace(name, nameof(name)); _commands.Add(new EditCommand(EditKind.DefaultValue, name, value: value)); return this; }
    /// <summary>Replaces raw field flags.</summary>
    public PdfAcroFormEditSession SetFlags(string name, int flags) { Guard.NotNullOrWhiteSpace(name, nameof(name)); _commands.Add(new EditCommand(EditKind.Flags, name, number: flags)); return this; }
    /// <summary>Replaces AcroForm calculation order with exact named fields.</summary>
    public PdfAcroFormEditSession SetCalculationOrder(params string[] fieldNames) { Guard.NotNull(fieldNames, nameof(fieldNames)); _commands.Add(new EditCommand(EditKind.CalculationOrder, names: fieldNames.ToArray())); return this; }
    /// <summary>Sets a page /Tabs order hint.</summary>
    public PdfAcroFormEditSession SetTabOrder(int pageNumber, PdfPageTabOrder order) {
        if (order != PdfPageTabOrder.Row &&
            order != PdfPageTabOrder.Column &&
            order != PdfPageTabOrder.Structure &&
            order != PdfPageTabOrder.Annotations) throw new ArgumentOutOfRangeException(nameof(order));
        _commands.Add(new EditCommand(EditKind.TabOrder, pageNumber: pageNumber, number: (int)order));
        return this;
    }
    /// <summary>Marks exact fields for visual flattening after tree edits.</summary>
    public PdfAcroFormEditSession Flatten(params string[] fieldNames) { Guard.NotNull(fieldNames, nameof(fieldNames)); _commands.Add(new EditCommand(EditKind.Flatten, names: fieldNames.ToArray())); return this; }
    internal IReadOnlyList<EditCommand> Commands => _commands.AsReadOnly();
    private byte[]? ValidateJavaScript(string? source) {
        if (source is null) return null;
        if (source.Length == 0) throw new ArgumentException("PDF widget JavaScript cannot be empty.", nameof(source));
        byte[] encoded = PdfJavaScriptStringEncoding.EncodeUnicode(source, nameof(source));
        int maximumBytes = Math.Min(_limits.MaxJavaScriptBytes, _limits.MaxDecodedStreamBytes);
        if (encoded.Length > maximumBytes) throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximumBytes, encoded.Length);
        return encoded;
    }
    private void TrackCreatedJavaScript(PdfFormFieldCreateOptions options, byte[]? encodedJavaScript) {
        if (encodedJavaScript is null) return;
        int count = options.Kind == PdfFormFieldCreationKind.RadioButtonGroup ? options.ChoiceOptions.Count : 1;
        long bytes = checked(encodedJavaScript.LongLength * count);
        _createdJavaScripts[options.Name] = new JavaScriptContribution(count, bytes);
        ValidateCreatedJavaScriptBudget();
    }
    private void RenameCreatedJavaScript(string name, string newName) {
        string descendantPrefix = name + ".";
        foreach (KeyValuePair<string, JavaScriptContribution> item in _createdJavaScripts
                     .Where(item => string.Equals(item.Key, name, StringComparison.Ordinal) || item.Key.StartsWith(descendantPrefix, StringComparison.Ordinal))
                     .ToArray()) {
            _createdJavaScripts.Remove(item.Key);
            _createdJavaScripts[newName + item.Key.Remove(0, name.Length)] = item.Value;
        }
    }
    private void RemoveCreatedJavaScript(string name) {
        string descendantPrefix = name + ".";
        foreach (string createdName in _createdJavaScripts.Keys
                     .Where(candidate => string.Equals(candidate, name, StringComparison.Ordinal) || candidate.StartsWith(descendantPrefix, StringComparison.Ordinal))
                     .ToArray()) {
            _createdJavaScripts.Remove(createdName);
        }
    }
    private void ValidateCreatedJavaScriptBudget() {
        int count = _createdJavaScripts.Values.Sum(static contribution => contribution.Count);
        if (count > _limits.MaxJavaScripts) throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScripts, _limits.MaxJavaScripts, count);
        long bytes = _createdJavaScripts.Values.Sum(static contribution => contribution.Bytes);
        if (bytes > _limits.MaxTotalJavaScriptBytes) throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScriptBytes, _limits.MaxTotalJavaScriptBytes, bytes);
    }
    private static void ValidateRectangle(double x, double y, double width, double height) {
        if (!IsFinite(x)) throw new ArgumentOutOfRangeException(nameof(x), "Field X must be finite.");
        if (!IsFinite(y)) throw new ArgumentOutOfRangeException(nameof(y), "Field Y must be finite.");
        if (!IsFinite(width) || width <= 0D) throw new ArgumentOutOfRangeException(nameof(width), "Field width must be positive and finite.");
        if (!IsFinite(height) || height <= 0D) throw new ArgumentOutOfRangeException(nameof(height), "Field height must be positive and finite.");
        if (!IsFinite(x + width)) throw new ArgumentOutOfRangeException(nameof(width), "Field right edge must be finite.");
        if (!IsFinite(y + height)) throw new ArgumentOutOfRangeException(nameof(height), "Field top edge must be finite.");
    }
    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private void AddName(EditKind kind, string name) { Guard.NotNullOrWhiteSpace(name, nameof(name)); _commands.Add(new EditCommand(kind, name)); }
    private void AddNames(EditKind kind, string name, string value) { Guard.NotNullOrWhiteSpace(name, nameof(name)); Guard.NotNullOrWhiteSpace(value, nameof(value)); _commands.Add(new EditCommand(kind, name, value: value)); }
    internal enum EditKind { Create, Rename, Remove, Move, DefaultValue, Flags, CalculationOrder, TabOrder, Flatten }
    private readonly struct JavaScriptContribution {
        internal JavaScriptContribution(int count, long bytes) { Count = count; Bytes = bytes; }
        internal int Count { get; }
        internal long Bytes { get; }
    }
    internal sealed class EditCommand {
        internal EditCommand(EditKind kind, string? name = null, string? value = null, int pageNumber = 0, double[]? rectangle = null, int number = 0, string[]? names = null, PdfFormFieldCreateOptions? options = null, byte[]? encodedJavaScript = null) { Kind = kind; Name = name; Value = value; PageNumber = pageNumber; Rectangle = rectangle; Number = number; Names = names; Options = options; EncodedJavaScript = encodedJavaScript; }
        internal EditKind Kind { get; } internal string? Name { get; } internal string? Value { get; } internal int PageNumber { get; } internal double[]? Rectangle { get; } internal int Number { get; } internal string[]? Names { get; } internal PdfFormFieldCreateOptions? Options { get; } internal byte[]? EncodedJavaScript { get; }
    }
}
