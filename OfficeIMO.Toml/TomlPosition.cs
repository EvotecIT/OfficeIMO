namespace OfficeIMO.Toml;

/// <summary>
/// A location in the decoded TOML text. Offsets count UTF-16 code units from the start of the text; a byte-order mark
/// consumed while decoding a file or stream is not part of the text. Lines and columns are one-based, and columns count
/// UTF-16 code units from the start of the line. A node that was added by editing and not loaded from text has the
/// default position, whose line is 0.
/// </summary>
public readonly struct TomlPosition : IEquatable<TomlPosition> {
    /// <summary>Creates a position from its zero-based offset and one-based line and column.</summary>
    public TomlPosition(long offset, int line, int column) {
        Offset = offset;
        Line = line;
        Column = column;
    }

    /// <summary>Gets the zero-based offset in UTF-16 code units.</summary>
    public long Offset { get; }

    /// <summary>Gets the one-based line. Each <c>\r\n</c> or <c>\n</c> starts a new line.</summary>
    public int Line { get; }

    /// <summary>Gets the one-based column in UTF-16 code units.</summary>
    public int Column { get; }

    /// <inheritdoc />
    public bool Equals(TomlPosition other) => Offset == other.Offset && Line == other.Line && Column == other.Column;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is TomlPosition other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => unchecked((int)Offset * 397 ^ Line * 31 ^ Column);

    /// <summary>Formats the position as <c>line:column</c> for diagnostics.</summary>
    public override string ToString() => Line.ToString(CultureInfo.InvariantCulture) + ":" + Column.ToString(CultureInfo.InvariantCulture);

    /// <summary>Compares two positions for equality.</summary>
    public static bool operator ==(TomlPosition left, TomlPosition right) => left.Equals(right);

    /// <summary>Compares two positions for inequality.</summary>
    public static bool operator !=(TomlPosition left, TomlPosition right) => !left.Equals(right);
}
