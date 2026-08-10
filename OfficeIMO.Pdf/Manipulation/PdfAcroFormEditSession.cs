namespace OfficeIMO.Pdf;

/// <summary>Transactional existing-document AcroForm edit commands.</summary>
public sealed class PdfAcroFormEditSession {
    private readonly List<EditCommand> _commands = new List<EditCommand>();
    private readonly PdfReadLimits _limits;
    private int _javaScriptCount;
    private long _javaScriptBytes;
    /// <summary>Creates a standalone edit session. Documents normally create sessions through <see cref="PdfDocumentForms.Edit(Action{PdfAcroFormEditSession})"/>.</summary>
    public PdfAcroFormEditSession() : this(new PdfReadLimits(), 0, 0L) { }
    internal PdfAcroFormEditSession(PdfReadLimits limits, int javaScriptCount, long javaScriptBytes) {
        _limits = limits;
        _javaScriptCount = javaScriptCount;
        _javaScriptBytes = javaScriptBytes;
    }
    /// <summary>Creates a text, checkbox, choice, radio-button, push-button, or empty signature field.</summary>
    public PdfAcroFormEditSession Create(PdfFormFieldCreateOptions options) {
        Guard.NotNull(options, nameof(options));
        PdfFormFieldCreateOptions snapshot = options.Snapshot();
        ReserveJavaScript(snapshot.JavaScript, snapshot.Kind == PdfFormFieldCreationKind.RadioButtonGroup ? snapshot.ChoiceOptions.Count : 1);
        _commands.Add(new EditCommand(EditKind.Create, options: snapshot));
        return this;
    }
    /// <summary>Places an empty signature field owned by the signature engine.</summary>
    public PdfAcroFormEditSession PlaceSignatureField(string name, int pageNumber, double x, double y, double width, double height) => Create(new PdfFormFieldCreateOptions { Name = name, Kind = PdfFormFieldCreationKind.Signature, PageNumber = pageNumber, X = x, Y = y, Width = width, Height = height });
    /// <summary>Renames one fully qualified field.</summary>
    public PdfAcroFormEditSession Rename(string name, string newName) { AddNames(EditKind.Rename, name, newName); return this; }
    /// <summary>Removes one field subtree and its widgets.</summary>
    public PdfAcroFormEditSession Remove(string name) { AddName(EditKind.Remove, name); return this; }
    /// <summary>Moves a single-widget field to a page rectangle.</summary>
    public PdfAcroFormEditSession Move(string name, int pageNumber, double x, double y, double width, double height) { Guard.NotNullOrWhiteSpace(name, nameof(name)); _commands.Add(new EditCommand(EditKind.Move, name, pageNumber: pageNumber, rectangle: new[] { x, y, x + width, y + height })); return this; }
    /// <summary>Sets or clears a field default value.</summary>
    public PdfAcroFormEditSession SetDefaultValue(string name, string? value) { Guard.NotNullOrWhiteSpace(name, nameof(name)); _commands.Add(new EditCommand(EditKind.DefaultValue, name, value: value)); return this; }
    /// <summary>Replaces raw field flags.</summary>
    public PdfAcroFormEditSession SetFlags(string name, int flags) { Guard.NotNullOrWhiteSpace(name, nameof(name)); _commands.Add(new EditCommand(EditKind.Flags, name, number: flags)); return this; }
    /// <summary>Replaces AcroForm calculation order with exact named fields.</summary>
    public PdfAcroFormEditSession SetCalculationOrder(params string[] fieldNames) { Guard.NotNull(fieldNames, nameof(fieldNames)); _commands.Add(new EditCommand(EditKind.CalculationOrder, names: fieldNames.ToArray())); return this; }
    /// <summary>Sets a page /Tabs order hint.</summary>
    public PdfAcroFormEditSession SetTabOrder(int pageNumber, PdfPageTabOrder order) { _commands.Add(new EditCommand(EditKind.TabOrder, pageNumber: pageNumber, number: (int)order)); return this; }
    /// <summary>Marks exact fields for visual flattening after tree edits.</summary>
    public PdfAcroFormEditSession Flatten(params string[] fieldNames) { Guard.NotNull(fieldNames, nameof(fieldNames)); _commands.Add(new EditCommand(EditKind.Flatten, names: fieldNames.ToArray())); return this; }
    internal IReadOnlyList<EditCommand> Commands => _commands.AsReadOnly();
    private void ReserveJavaScript(string? source, int actionCount) {
        if (source is null) return;
        if (source.Length == 0) throw new ArgumentException("PDF widget JavaScript cannot be empty.", nameof(source));
        byte[] encoded = PdfJavaScriptStringEncoding.EncodeUnicode(source, nameof(source));
        int maximumBytes = Math.Min(_limits.MaxJavaScriptBytes, _limits.MaxDecodedStreamBytes);
        if (encoded.Length > maximumBytes) throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximumBytes, encoded.Length);
        int nextCount = checked(_javaScriptCount + actionCount);
        if (nextCount > _limits.MaxJavaScripts) throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScripts, _limits.MaxJavaScripts, nextCount);
        long nextBytes = checked(_javaScriptBytes + encoded.LongLength * actionCount);
        if (nextBytes > _limits.MaxTotalJavaScriptBytes) throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScriptBytes, _limits.MaxTotalJavaScriptBytes, nextBytes);
        _javaScriptCount = nextCount;
        _javaScriptBytes = nextBytes;
    }
    private void AddName(EditKind kind, string name) { Guard.NotNullOrWhiteSpace(name, nameof(name)); _commands.Add(new EditCommand(kind, name)); }
    private void AddNames(EditKind kind, string name, string value) { Guard.NotNullOrWhiteSpace(name, nameof(name)); Guard.NotNullOrWhiteSpace(value, nameof(value)); _commands.Add(new EditCommand(kind, name, value: value)); }
    internal enum EditKind { Create, Rename, Remove, Move, DefaultValue, Flags, CalculationOrder, TabOrder, Flatten }
    internal sealed class EditCommand {
        internal EditCommand(EditKind kind, string? name = null, string? value = null, int pageNumber = 0, double[]? rectangle = null, int number = 0, string[]? names = null, PdfFormFieldCreateOptions? options = null) { Kind = kind; Name = name; Value = value; PageNumber = pageNumber; Rectangle = rectangle; Number = number; Names = names; Options = options; }
        internal EditKind Kind { get; } internal string? Name { get; } internal string? Value { get; } internal int PageNumber { get; } internal double[]? Rectangle { get; } internal int Number { get; } internal string[]? Names { get; } internal PdfFormFieldCreateOptions? Options { get; }
    }
}
