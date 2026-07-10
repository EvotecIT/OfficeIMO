namespace OfficeIMO.AsciiDoc;

/// <summary>Context passed to an explicitly registered custom directive processor.</summary>
public sealed class AsciiDocDirectiveContext {
    internal AsciiDocDirectiveContext(
        string name,
        string target,
        string attributeList,
        string originalText,
        string? sourceName,
        int line,
        AsciiDocDocumentAttributes attributes) {
        Name = name;
        Target = target;
        AttributeList = attributeList;
        OriginalText = originalText;
        SourceName = sourceName;
        Line = line;
        Attributes = attributes;
    }

    /// <summary>Directive name.</summary>
    public string Name { get; }

    /// <summary>Raw directive target.</summary>
    public string Target { get; }

    /// <summary>Raw content between square brackets.</summary>
    public string AttributeList { get; }

    /// <summary>Exact directive line including its line ending.</summary>
    public string OriginalText { get; }

    /// <summary>Source identifier, when known.</summary>
    public string? SourceName { get; }

    /// <summary>One-based line number.</summary>
    public int Line { get; }

    /// <summary>Attributes visible at the directive.</summary>
    public AsciiDocDocumentAttributes Attributes { get; }
}

/// <summary>Result from a custom directive processor.</summary>
public sealed class AsciiDocDirectiveResult {
    private AsciiDocDirectiveResult(string? replacement, bool preserveOriginal) {
        Replacement = replacement;
        PreserveOriginal = preserveOriginal;
    }

    /// <summary>Replacement AsciiDoc source.</summary>
    public string? Replacement { get; }

    /// <summary>True when the original directive should remain.</summary>
    public bool PreserveOriginal { get; }

    /// <summary>Replaces the directive with source text.</summary>
    public static AsciiDocDirectiveResult Replace(string source) => new AsciiDocDirectiveResult(source ?? throw new ArgumentNullException(nameof(source)), false);

    /// <summary>Removes the directive.</summary>
    public static AsciiDocDirectiveResult Remove() => new AsciiDocDirectiveResult(string.Empty, false);

    /// <summary>Leaves the original directive untouched.</summary>
    public static AsciiDocDirectiveResult Preserve() => new AsciiDocDirectiveResult(null, true);
}

/// <summary>Custom block-directive processor supplied as a normal .NET instance.</summary>
public interface IAsciiDocDirectiveProcessor {
    /// <summary>Processes one directive.</summary>
    AsciiDocDirectiveResult Process(AsciiDocDirectiveContext context);
}

/// <summary>Explicit registry; documents cannot discover or load processors.</summary>
public sealed class AsciiDocExtensionRegistry {
    private readonly Dictionary<string, IAsciiDocDirectiveProcessor> _directives =
        new Dictionary<string, IAsciiDocDirectiveProcessor>(StringComparer.Ordinal);

    /// <summary>Registers or replaces a custom directive processor.</summary>
    public AsciiDocExtensionRegistry RegisterDirective(string name, IAsciiDocDirectiveProcessor processor) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        if (processor == null) throw new ArgumentNullException(nameof(processor));
        if (!AsciiDocText.IsMacroName(name)) throw new ArgumentException("Invalid directive name.", nameof(name));
        if (IsReserved(name)) throw new ArgumentException("Built-in preprocessing directives cannot be replaced.", nameof(name));
        _directives[name] = processor;
        return this;
    }

    internal bool TryGetDirective(string name, out IAsciiDocDirectiveProcessor processor) =>
        _directives.TryGetValue(name, out processor!);

    private static bool IsReserved(string name) =>
        string.Equals(name, "include", StringComparison.Ordinal) ||
        string.Equals(name, "ifdef", StringComparison.Ordinal) ||
        string.Equals(name, "ifndef", StringComparison.Ordinal) ||
        string.Equals(name, "ifeval", StringComparison.Ordinal) ||
        string.Equals(name, "endif", StringComparison.Ordinal);
}
