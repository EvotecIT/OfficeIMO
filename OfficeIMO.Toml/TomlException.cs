namespace OfficeIMO.Toml;

/// <summary>
/// Raised when TOML input is malformed, breaks a table or key rule, or cannot be decoded. <see cref="Code"/> is stable
/// across releases and identifies the rule that failed; the message is for people and may change.
/// </summary>
public class TomlException : Exception {
    internal TomlException(string code, string message, TomlPosition position, Exception? innerException = null)
        : base(FormatMessage(message, position), innerException) {
        Code = code;
        Position = position;
    }

    /// <summary>Gets the stable error code, one of the <see cref="TomlErrorCodes"/> values.</summary>
    public string Code { get; }

    /// <summary>Gets where the problem was detected in the decoded text.</summary>
    public TomlPosition Position { get; }

    /// <summary>Gets the one-based line where the problem was detected.</summary>
    public int Line => Position.Line;

    /// <summary>Gets the one-based column where the problem was detected.</summary>
    public int Column => Position.Column;

    private static string FormatMessage(string message, TomlPosition position) =>
        position.Line > 0
            ? message + " (line " + position.Line.ToString(CultureInfo.InvariantCulture) + ", column " + position.Column.ToString(CultureInfo.InvariantCulture) + ")"
            : message;
}

/// <summary>
/// Raised when input exceeds one of the limits in <see cref="TomlReaderOptions"/> or <see cref="TomlLoadOptions"/>. The
/// input may be valid TOML; the limit protects the caller from resource exhaustion.
/// </summary>
public sealed class TomlLimitExceededException : TomlException {
    internal TomlLimitExceededException(string limitName, long limit, TomlPosition position)
        : base(TomlErrorCodes.LimitExceeded, "The TOML input exceeds " + limitName + " (" + limit.ToString(CultureInfo.InvariantCulture) + ").", position) {
        LimitName = limitName;
        Limit = limit;
    }

    /// <summary>Gets the name of the exceeded option, such as <c>MaxDepth</c>.</summary>
    public string LimitName { get; }

    /// <summary>Gets the configured value of the exceeded limit.</summary>
    public long Limit { get; }
}

/// <summary>Stable codes carried by <see cref="TomlException.Code"/>. Codes are never reused for a different rule.</summary>
public static class TomlErrorCodes {
    /// <summary>A character that TOML does not allow at this point, including control characters and a carriage return outside <c>\r\n</c>.</summary>
    public const string InvalidCharacter = "Toml.Syntax.InvalidCharacter";
    /// <summary>A token that the grammar does not allow at this point, such as a missing <c>=</c>, comma or newline.</summary>
    public const string UnexpectedToken = "Toml.Syntax.UnexpectedToken";
    /// <summary>A missing, empty or malformed key, including a multi-line string used as a key.</summary>
    public const string InvalidKey = "Toml.Syntax.InvalidKey";
    /// <summary>A string without its closing delimiter.</summary>
    public const string UnterminatedString = "Toml.Syntax.UnterminatedString";
    /// <summary>An unknown escape or an escape that is not a Unicode scalar value.</summary>
    public const string InvalidEscape = "Toml.Syntax.InvalidEscape";
    /// <summary>A malformed integer, or one outside the 64-bit signed range.</summary>
    public const string InvalidInteger = "Toml.Syntax.InvalidInteger";
    /// <summary>A malformed float.</summary>
    public const string InvalidFloat = "Toml.Syntax.InvalidFloat";
    /// <summary>A malformed date or time, or one with an out-of-range field such as February 30.</summary>
    public const string InvalidDateTime = "Toml.Syntax.InvalidDateTime";
    /// <summary>A value that is not a string, number, boolean, date-time, array or inline table.</summary>
    public const string InvalidValue = "Toml.Syntax.InvalidValue";
    /// <summary>Bytes that are not valid UTF-8.</summary>
    public const string InvalidEncoding = "Toml.Encoding.Invalid";
    /// <summary>A key defined twice in the same table.</summary>
    public const string DuplicateKey = "Toml.Load.DuplicateKey";
    /// <summary>
    /// A table defined twice, or extended in a way TOML forbids: a <c>[table]</c> header for a table that dotted keys,
    /// an inline table or an earlier header defined, dotted keys added to a table defined elsewhere, or a header or key
    /// that goes through a value that is not a table.
    /// </summary>
    public const string InvalidTable = "Toml.Load.InvalidTable";
    /// <summary>An array of tables mixed with a static array or a table of the same name.</summary>
    public const string InvalidArrayOfTables = "Toml.Load.InvalidArrayOfTables";
    /// <summary>A limit from <see cref="TomlReaderOptions"/> or <see cref="TomlLoadOptions"/> was exceeded; see <see cref="TomlLimitExceededException"/>.</summary>
    public const string LimitExceeded = "Toml.Limit.Exceeded";
    /// <summary>A value <see cref="TomlSerializer"/> cannot read as the target type, such as a string for an integer property.</summary>
    public const string TypeMismatch = "Toml.Serialization.TypeMismatch";
    /// <summary>A table without a key that the target type requires.</summary>
    public const string MissingRequiredMember = "Toml.Serialization.MissingRequiredMember";
    /// <summary>A discriminator that is missing or not registered with <c>[OfficeDerivedType]</c> for the target type.</summary>
    public const string UnknownDiscriminator = "Toml.Serialization.UnknownDiscriminator";
    /// <summary>An object that contains itself, which TOML cannot write.</summary>
    public const string ReferenceCycle = "Toml.Serialization.ReferenceCycle";
    /// <summary>
    /// A value TOML cannot hold: a null array item, a null property or dictionary value under
    /// <see cref="TomlNullHandling.Error"/>, or an integer outside the 64-bit signed range.
    /// </summary>
    public const string UnrepresentableValue = "Toml.Serialization.UnrepresentableValue";
}
