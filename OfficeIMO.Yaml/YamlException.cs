namespace OfficeIMO.Yaml;

/// <summary>
/// Raised when YAML input is malformed or cannot be decoded. <see cref="Code"/> is stable across releases and
/// identifies the rule that failed; the message is for people and may change.
/// </summary>
public class YamlException : Exception {
    internal YamlException(string code, string message, YamlPosition position, Exception? innerException = null)
        : base(FormatMessage(message, position), innerException) {
        Code = code;
        Position = position;
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
