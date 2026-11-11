namespace OfficeIMO.Yaml;

/// <summary>
/// Raised when YAML input is malformed or cannot be decoded. <see cref="Code"/> is stable across releases and
/// identifies the rule that failed; the message is for people and may change.
/// </summary>
public class YamlException : Exception {
    private readonly string _reason;

    internal YamlException(string code, string message, YamlPosition position, Exception? innerException = null)
        : base(FormatMessage(message, position), innerException) {
        Code = code;
        Position = position;
        _reason = message;
    }

    /// <summary>Gets the stable error code, one of the <see cref="YamlErrorCodes"/> values.</summary>
    public string Code { get; }

    /// <summary>Gets where the problem was detected in the decoded text.</summary>
    public YamlPosition Position { get; }

    /// <summary>Gets the one-based line where the problem was detected.</summary>
    public int Line => Position.Line;

    /// <summary>Gets the one-based column where the problem was detected.</summary>
    public int Column => Position.Column;

    private string? _excerpt;

    /// <summary>
    /// Gets the message, followed by the offending source line with a caret under the column when the source text was
    /// available (input given as a string, memory or file loaded by <see cref="YamlDocument"/>).
    /// </summary>
    public override string Message => _excerpt is null ? base.Message : base.Message + _excerpt;

    /// <summary>
    /// Returns the failure as an error <see cref="OfficeDiagnostic"/>, the form <see cref="YamlDocument.Diagnostics"/>
    /// uses, so a caller can report a file that does not load next to the findings of files that do.
    /// </summary>
    /// <remarks>
    /// <para>The code is the catalog form of <see cref="Code"/>: <c>Yaml.Syntax.InvalidIndentation</c> becomes
    /// <c>yaml-syntax-invalid-indentation</c>, and <see cref="YamlErrorCodes.DuplicateKey"/> becomes
    /// <see cref="OfficeDiagnosticCodes.YamlDuplicateKey"/>, the code a non-strict load reports as a warning.</para>
    /// <para>The location is <c>line:column</c> as in <see cref="YamlDocument.Diagnostics"/>, or <see langword="null"/>
    /// when the failure has no position. The message leaves the position out and ends with the source excerpt when the
    /// source text was available. A host file that embeds the YAML, such as Markdown front matter, adds its own line
    /// offset to <see cref="Line"/>.</para>
    /// </remarks>
    public OfficeDiagnostic ToDiagnostic() =>
        new OfficeDiagnostic(DiagnosticCode(Code), OfficeDiagnosticSeverity.Error, _excerpt is null ? _reason : _reason + _excerpt,
            Position.Line > 0 ? Position.ToString() : null, "YAML");

    /// <summary>Converts a <see cref="YamlErrorCodes"/> value to its catalog code (the <c>Yaml.{area}.{rule}</c> pattern).</summary>
    internal static string DiagnosticCode(string code) {
        if (code == YamlErrorCodes.DuplicateKey) return OfficeDiagnosticCodes.YamlDuplicateKey;
        var text = new StringBuilder(code.Length + 8);
        for (int i = 0; i < code.Length; i++) {
            char ch = code[i];
            if (ch == '.') {
                text.Append('-');
            } else if (char.IsUpper(ch)) {
                if (i > 0 && char.IsLower(code[i - 1])) text.Append('-');
                text.Append(char.ToLowerInvariant(ch));
            } else {
                text.Append(ch);
            }
        }
        return text.ToString();
    }

    /// <summary>Adds the source excerpt; returns false so it can run in an exception filter without catching.</summary>
    internal bool AttachSource(string source) {
        if (_excerpt is null && source.Length > 0 && Position.Line > 0 && Position.Offset <= source.Length) _excerpt = YamlExcerpt.Format(source, (int)Position.Offset);
        return false;
    }

    private static string FormatMessage(string message, YamlPosition position) =>
        position.Line > 0
            ? message + " (line " + position.Line.ToString(CultureInfo.InvariantCulture) + ", column " + position.Column.ToString(CultureInfo.InvariantCulture) + ")"
            : message;
}

/// <summary>
/// Raised when input exceeds one of the limits in <see cref="YamlReaderOptions"/> or <see cref="YamlLoadOptions"/>. The input may be valid YAML;
/// the limit protects the caller from resource exhaustion.
/// </summary>
public sealed class YamlLimitExceededException : YamlException {
    internal YamlLimitExceededException(string limitName, long limit, YamlPosition position)
        : base(YamlErrorCodes.LimitExceeded, "The YAML input exceeds " + limitName + " (" + limit.ToString(CultureInfo.InvariantCulture) + ").", position) {
        LimitName = limitName;
        Limit = limit;
    }

    /// <summary>Gets the name of the exceeded option, such as <c>MaxDepth</c>.</summary>
    public string LimitName { get; }

    /// <summary>Gets the configured value of the exceeded limit.</summary>
    public long Limit { get; }
}

/// <summary>
/// Stable codes carried by <see cref="YamlException.Code"/>. Codes are never reused for a different rule.
/// </summary>
public static class YamlErrorCodes {
    /// <summary>A character that cannot start or continue any YAML token, including control characters.</summary>
    public const string InvalidCharacter = "Yaml.Syntax.InvalidCharacter";
    /// <summary>A token that the grammar does not allow at this point.</summary>
    public const string UnexpectedToken = "Yaml.Syntax.UnexpectedToken";
    /// <summary>Indentation that does not match the enclosing block structure, including tabs used as indentation.</summary>
    public const string InvalidIndentation = "Yaml.Syntax.InvalidIndentation";
    /// <summary>An implicit key that spans lines, exceeds 1024 characters, or is missing its <c>:</c> indicator.</summary>
    public const string InvalidKey = "Yaml.Syntax.InvalidKey";
    /// <summary>A <c>:</c> value indicator where a mapping value cannot start.</summary>
    public const string InvalidValue = "Yaml.Syntax.InvalidValue";
    /// <summary>A quoted scalar without its closing quote.</summary>
    public const string UnterminatedScalar = "Yaml.Syntax.UnterminatedScalar";
    /// <summary>An unknown or malformed escape sequence in a double-quoted scalar.</summary>
    public const string InvalidEscape = "Yaml.Syntax.InvalidEscape";
    /// <summary>A malformed block scalar header or block scalar indentation.</summary>
    public const string InvalidBlockScalar = "Yaml.Syntax.InvalidBlockScalar";
    /// <summary>A malformed, duplicated or misplaced <c>%YAML</c> or <c>%TAG</c> directive, or an unsupported major version.</summary>
    public const string InvalidDirective = "Yaml.Syntax.InvalidDirective";
    /// <summary>A malformed tag, or a tag handle no <c>%TAG</c> directive declares.</summary>
    public const string InvalidTag = "Yaml.Syntax.InvalidTag";
    /// <summary>A malformed anchor or alias name, or properties attached to an alias.</summary>
    public const string InvalidAnchor = "Yaml.Syntax.InvalidAnchor";
    /// <summary>A document marker, directive or content where the stream structure does not allow it.</summary>
    public const string InvalidDocument = "Yaml.Syntax.InvalidDocument";
    /// <summary>Bytes that are not valid in the detected Unicode encoding.</summary>
    public const string InvalidEncoding = "Yaml.Encoding.Invalid";
    /// <summary>An alias (<c>*name</c>) that no preceding anchor in the same document defines.</summary>
    public const string UndefinedAlias = "Yaml.Load.UndefinedAlias";
    /// <summary>
    /// A mapping with two equal keys. <see cref="YamlDocument"/> throws it under <see cref="YamlDuplicateKeyPolicy.Error"/>,
    /// the default of strict loads, and reports it as a warning diagnostic under the other policies.
    /// <see cref="YamlSerializer"/> throws it only when <see cref="YamlReaderOptions.Strict"/> is set; otherwise the last key wins.
    /// </summary>
    public const string DuplicateKey = "Yaml.Load.DuplicateKey";
    /// <summary>An alias inside the collection its anchor names (<c>&amp;a [*a]</c>), which would make the document infinite.</summary>
    public const string RecursiveAlias = "Yaml.Load.RecursiveAlias";
    /// <summary>A limit from <see cref="YamlReaderOptions"/> or <see cref="YamlLoadOptions"/> was exceeded; see <see cref="YamlLimitExceededException"/>.</summary>
    public const string LimitExceeded = "Yaml.Limit.Exceeded";
    /// <summary>
    /// <see cref="YamlSerializer"/> found a node that cannot become the requested .NET type: a mapping where a list is
    /// expected, text that is not a number, a null for a non-nullable value type, or a value out of range.
    /// </summary>
    public const string TypeMismatch = "Yaml.Serialization.TypeMismatch";
    /// <summary>A mapping read by <see cref="YamlSerializer"/> lacks a key for a required member or constructor parameter.</summary>
    public const string MissingRequiredMember = "Yaml.Serialization.MissingRequiredMember";
    /// <summary>A polymorphic value has no discriminator its type accepts, or one that no derived type registers.</summary>
    public const string UnknownDiscriminator = "Yaml.Serialization.UnknownDiscriminator";
    /// <summary>
    /// <see cref="YamlSerializer"/> reached an object that contains itself while
    /// <see cref="YamlSerializerOptions.ReferenceHandling"/> is <see cref="YamlReferenceHandling.None"/>.
    /// </summary>
    public const string ReferenceCycle = "Yaml.Serialization.ReferenceCycle";
}
