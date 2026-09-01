namespace OfficeIMO.Pdf;

/// <summary>Transactional collection editor for named document-level JavaScript actions.</summary>
public sealed class PdfJavaScriptEditSession {
    private readonly List<EditCommand> _commands = new List<EditCommand>();
    private readonly List<string> _operations = new List<string>();
    private readonly int _maxJavaScriptBytes;
    private readonly int _maxJavaScripts;
    private readonly long _maxTotalJavaScriptBytes;
    private long _commandJavaScriptBytes;

    internal PdfJavaScriptEditSession(PdfReadLimits limits) {
        _maxJavaScriptBytes = limits.MaxJavaScriptBytes;
        _maxJavaScripts = limits.MaxJavaScripts;
        _maxTotalJavaScriptBytes = limits.MaxTotalJavaScriptBytes;
    }

    /// <summary>Adds a named script or replaces the source of an existing script with the same name.</summary>
    public PdfJavaScriptEditSession AddOrReplace(string name, string script) {
        Guard.NotNull(name, nameof(name));
        Guard.NotNull(script, nameof(script));
        if (name.Length == 0) throw new ArgumentException("JavaScript name cannot be empty.", nameof(name));
        if (script.Length == 0) throw new ArgumentException("JavaScript source cannot be empty.", nameof(script));
        byte[] encodedName = PdfJavaScriptStringEncoding.EncodeUnicode(name, nameof(name));
        byte[] encodedScript = PdfJavaScriptStringEncoding.EncodeUnicode(script, nameof(script));
        int byteCount = encodedScript.Length;
        if (byteCount > _maxJavaScriptBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, _maxJavaScriptBytes, byteCount);
        }
        EnsureCommandBudget(byteCount);
        _commands.Add(new EditCommand(EditKind.AddOrReplace, name, script, encodedName, encodedScript));
        _operations.Add("AddOrReplace:" + name);
        return this;
    }

    /// <summary>Removes a named script when it exists.</summary>
    public PdfJavaScriptEditSession Remove(string name) {
        Guard.NotNull(name, nameof(name));
        if (name.Length == 0) throw new ArgumentException("JavaScript name cannot be empty.", nameof(name));
        byte[] encodedName = PdfJavaScriptStringEncoding.EncodeUnicode(name, nameof(name));
        EnsureCommandBudget(0L);
        _commands.Add(new EditCommand(EditKind.Remove, name, null, encodedName, null));
        _operations.Add("Remove:" + name);
        return this;
    }

    /// <summary>Removes every named document-level script.</summary>
    public PdfJavaScriptEditSession Clear() {
        EnsureCommandBudget(0L);
        _commands.Add(new EditCommand(EditKind.Clear, null, null, null, null));
        _operations.Add("Clear");
        return this;
    }

    internal IReadOnlyList<string> Operations => _operations.AsReadOnly();
    internal IReadOnlyList<EditCommand> Commands => _commands.AsReadOnly();

    private void EnsureCommandBudget(long additionalBytes) {
        int nextCount = checked(_commands.Count + 1);
        if (nextCount > _maxJavaScripts) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScripts, _maxJavaScripts, nextCount);
        }
        long nextBytes = checked(_commandJavaScriptBytes + additionalBytes);
        if (nextBytes > _maxTotalJavaScriptBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.JavaScriptBytes, _maxTotalJavaScriptBytes, nextBytes);
        }
        _commandJavaScriptBytes = nextBytes;
    }

    internal enum EditKind { AddOrReplace, Remove, Clear }

    internal sealed class EditCommand {
        internal EditCommand(EditKind kind, string? name, string? script, byte[]? encodedName, byte[]? encodedScript) {
            Kind = kind; Name = name; Script = script; EncodedName = encodedName; EncodedScript = encodedScript;
        }
        internal EditKind Kind { get; }
        internal string? Name { get; }
        internal string? Script { get; }
        internal byte[]? EncodedName { get; }
        internal byte[]? EncodedScript { get; }
    }
}

/// <summary>Edited PDF bytes with exact named-script readback and rewrite-preservation proof.</summary>
public sealed class PdfJavaScriptEditResult {
    private readonly byte[] _pdf;
    private readonly PdfLoadOptions _readOptions;

    internal PdfJavaScriptEditResult(
        byte[] pdf,
        PdfMutationPlan mutationPlan,
        PdfRewritePreservationReport preservationReport,
        IReadOnlyList<PdfJavaScript> javaScripts,
        IReadOnlyList<string> operations,
        PdfLoadOptions readOptions) {
        _pdf = (byte[])pdf.Clone();
        _readOptions = readOptions;
        MutationPlan = mutationPlan;
        PreservationReport = preservationReport;
        JavaScripts = javaScripts.Count == 0
            ? Array.Empty<PdfJavaScript>()
            : Array.AsReadOnly(javaScripts.ToArray());
        Operations = operations.Count == 0
            ? Array.Empty<string>()
            : Array.AsReadOnly(operations.ToArray());
    }

    /// <summary>Shared full-rewrite mutation plan.</summary>
    public PdfMutationPlan MutationPlan { get; }

    /// <summary>Proof that structures outside the document JavaScript name tree survived the rewrite.</summary>
    public PdfRewritePreservationReport PreservationReport { get; }

    /// <summary>Named scripts read back from the saved artifact.</summary>
    public IReadOnlyList<PdfJavaScript> JavaScripts { get; }

    /// <summary>Stable operation descriptions applied in transaction order.</summary>
    public IReadOnlyList<string> Operations { get; }

    /// <summary>Returns a defensive copy of the edited PDF.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();

    /// <summary>Opens the edited artifact as a fluent PDF document.</summary>
    public PdfDocument ToDocument() => PdfDocument.Load(_pdf, _readOptions);
}
