using System.Globalization;
using System.Text;
namespace OfficeIMO.Json;

/// <summary>The JSON Schema dialect a schema is read with.</summary>
public enum JsonSchemaDraft : byte {
    /// <summary>Draft 2020-12, the default for schemas without a recognised <c>$schema</c>.</summary>
    Draft202012 = 0,
    /// <summary>Draft 7 (also used for draft 6 schemas): <c>definitions</c>, array-form <c>items</c> with <c>additionalItems</c>, <c>dependencies</c>, and <c>$ref</c> that overrides its sibling keywords.</summary>
    Draft7
}

/// <summary>How <see cref="JsonSchema"/> reads a schema and how far validation may go.</summary>
/// <remarks>
/// The limits bound the work an untrusted schema or document can cause. Every limit reached is reported as a finding,
/// so a result is never silently incomplete.
/// </remarks>
public sealed class JsonSchemaOptions {
    private int _maxErrors = 100;
    private int _maxDepth = 256;
    private long _maxEvaluationSteps = 10_000_000;
    private TimeSpan _regexTimeout = TimeSpan.FromSeconds(1);

    /// <summary>Gets or sets the dialect. When null, the root schema's <c>$schema</c> decides, and draft 2020-12 is used when it is absent or not recognised.</summary>
    public JsonSchemaDraft? Draft { get; set; }

    /// <summary>
    /// Gets or sets the absolute URI the root schema was retrieved from, against which relative <c>$id</c> and <c>$ref</c>
    /// values resolve. When null and the root has no absolute <c>$id</c>, relative references reach
    /// <see cref="ReferenceResolver"/> exactly as written.
    /// </summary>
    public string? BaseUri { get; set; }

    /// <summary>
    /// Gets or sets the callback that supplies schemas for references outside the loaded document. It receives the
    /// resolved URI without its fragment and returns the schema, or null when it cannot. Each URI is requested at most
    /// once per <see cref="JsonSchema"/>. Nothing is ever fetched from the network or the file system without it.
    /// </summary>
    public Func<string, JsonData?>? ReferenceResolver { get; set; }

    /// <summary>
    /// Gets or sets whether <c>format</c> asserts the formats the validator knows (dates, times, durations, e-mail
    /// addresses, host names, IP addresses, URIs, UUIDs, regular expressions and JSON pointers). The default is true.
    /// Formats it does not know are never errors; <see cref="JsonSchemaValidationResult.UncheckedFormats"/> lists them.
    /// </summary>
    public bool ValidateFormats { get; set; } = true;

    /// <summary>Gets or sets the number of findings after which validation stops. The default is 100.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public int MaxErrors {
        get => _maxErrors;
        set => _maxErrors = value >= 1 ? value : throw new ArgumentOutOfRangeException(nameof(value));
    }

    /// <summary>Gets or sets how deeply schemas may nest during validation, counting subschemas and followed references. The default is 256.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public int MaxDepth {
        get => _maxDepth;
        set => _maxDepth = value >= 1 ? value : throw new ArgumentOutOfRangeException(nameof(value));
    }

    /// <summary>Gets or sets how many schema evaluations, counting each <c>uniqueItems</c> comparison of items with equal hashes, one validation may perform before it stops. The default is ten million.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is less than 1.</exception>
    public long MaxEvaluationSteps {
        get => _maxEvaluationSteps;
        set => _maxEvaluationSteps = value >= 1 ? value : throw new ArgumentOutOfRangeException(nameof(value));
    }

    /// <summary>Gets or sets how long one <c>pattern</c> or <c>patternProperties</c> match may run. The default is one second.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The value is not positive.</exception>
    public TimeSpan RegexTimeout {
        get => _regexTimeout;
        set => _regexTimeout = value > TimeSpan.Zero ? value : throw new ArgumentOutOfRangeException(nameof(value));
    }

    internal JsonSchemaOptions Clone() => (JsonSchemaOptions)MemberwiseClone();
}

/// <summary>The error thrown when a schema cannot be read: a keyword has a value of the wrong type, or a pattern is not a valid regular expression.</summary>
public sealed class JsonSchemaException : Exception {
    /// <summary>Creates the exception.</summary>
    public JsonSchemaException(string message) : base(message) {
    }

    /// <summary>Creates the exception with its cause.</summary>
    public JsonSchemaException(string message, Exception innerException) : base(message, innerException) {
    }
}

/// <summary>One reason a value does not match a schema, with where it is in the document and in the schema.</summary>
public sealed class JsonSchemaFinding {
    internal JsonSchemaFinding(string instancePath, string schemaPath, string keyword, string message, JsonData instance) {
        InstancePath = instancePath;
        SchemaPath = schemaPath;
        Keyword = keyword;
        Message = message;
        Line = instance.Line;
        Column = instance.Column;
        Source = instance.Source;
    }

    /// <summary>Gets the JSON Pointer (RFC 6901) of the value in the document, such as <c>/spec/replicas</c>; the root is the empty string.</summary>
    public string InstancePath { get; }

    /// <summary>Gets the JSON Pointer of the failing keyword along the evaluation path, through any <c>$ref</c> followed, such as <c>/properties/spec/$ref/required</c>.</summary>
    public string SchemaPath { get; }

    /// <summary>Gets the failing keyword, such as <c>type</c> or <c>required</c>.</summary>
    public string Keyword { get; }

    /// <summary>Gets a sentence describing the problem.</summary>
    public string Message { get; }

    /// <summary>Gets the one-based line of the value in the source, or 0 when the value has no position.</summary>
    public int Line { get; }

    /// <summary>Gets the one-based column of the value in the source, or 0 when the value has no position.</summary>
    public int Column { get; }

    /// <summary>Gets the source node of the value, such as a <c>YamlNode</c> or <c>TomlNode</c>, or null.</summary>
    public object? Source { get; }

    /// <summary>Formats the finding as <c>line:column path: message</c>.</summary>
    public override string ToString() {
        string path = InstancePath.Length == 0 ? "(root)" : InstancePath;
        return Line > 0
            ? Line.ToString(CultureInfo.InvariantCulture) + ":" + Column.ToString(CultureInfo.InvariantCulture) + " " + path + ": " + Message
            : path + ": " + Message;
    }
}

/// <summary>The outcome of <see cref="JsonSchema.Validate"/>.</summary>
public sealed class JsonSchemaValidationResult {
    internal JsonSchemaValidationResult(IReadOnlyList<JsonSchemaFinding> findings, bool truncated, IReadOnlyList<string> uncheckedFormats) {
        Findings = findings;
        IsTruncated = truncated;
        UncheckedFormats = uncheckedFormats;
    }

    /// <summary>Gets whether the value matches the schema.</summary>
    public bool IsValid => Findings.Count == 0;

    /// <summary>Gets the reasons the value does not match, in document order of discovery; empty when it matches.</summary>
    public IReadOnlyList<JsonSchemaFinding> Findings { get; }

    /// <summary>Gets whether validation stopped at <see cref="JsonSchemaOptions.MaxErrors"/>, so more findings may exist.</summary>
    public bool IsTruncated { get; }

    /// <summary>Gets the <c>format</c> values that applied to a string but were only annotated, because the validator does not know them or format assertion is off.</summary>
    public IReadOnlyList<string> UncheckedFormats { get; }
}
