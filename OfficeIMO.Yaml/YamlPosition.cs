namespace OfficeIMO.Yaml;

/// <summary>
/// A location in the decoded YAML text. Offsets count UTF-16 code units from the start of the text; a byte-order
/// mark consumed while decoding a stream is not part of the text. Lines and columns are one-based, and columns
/// count UTF-16 code units from the start of the line.
/// </summary>
public readonly struct YamlPosition : IEquatable<YamlPosition> {
    /// <summary>Creates a position from its zero-based offset and one-based line and column.</summary>
    public YamlPosition(long offset, int line, int column) {
        Offset = offset;
        Line = line;
        Column = column;
    }

    /// <summary>Gets the zero-based offset in UTF-16 code units. Spans between positions of consecutive events reproduce the source exactly.</summary>
    public long Offset { get; }

    /// <summary>Gets the one-based line. Each <c>\r\n</c>, <c>\n</c> or <c>\r</c> starts a new line.</summary>
    public int Line { get; }

    /// <summary>Gets the one-based column in UTF-16 code units.</summary>
    public int Column { get; }

    /// <inheritdoc />
    public bool Equals(YamlPosition other) => Offset == other.Offset && Line == other.Line && Column == other.Column;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is YamlPosition other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => unchecked((int)Offset * 397 ^ Line * 31 ^ Column);

    /// <summary>Formats the position as <c>line:column</c> for diagnostics.</summary>
    public override string ToString() => Line.ToString(CultureInfo.InvariantCulture) + ":" + Column.ToString(CultureInfo.InvariantCulture);

    /// <summary>Compares two positions for equality.</summary>
    public static bool operator ==(YamlPosition left, YamlPosition right) => left.Equals(right);

    /// <summary>Compares two positions for inequality.</summary>
    public static bool operator !=(YamlPosition left, YamlPosition right) => !left.Equals(right);
}
