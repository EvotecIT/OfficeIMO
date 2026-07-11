namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Identifies a character boundary in AsciiDoc source text.
/// </summary>
public readonly struct AsciiDocSourcePosition : IEquatable<AsciiDocSourcePosition> {
    /// <summary>Creates a source position.</summary>
    public AsciiDocSourcePosition(int offset, int line, int column) {
        if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset));
        if (line < 1) throw new ArgumentOutOfRangeException(nameof(line));
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));

        Offset = offset;
        Line = line;
        Column = column;
    }

    /// <summary>Zero-based UTF-16 character offset.</summary>
    public int Offset { get; }

    /// <summary>One-based line number.</summary>
    public int Line { get; }

    /// <summary>One-based UTF-16 column number.</summary>
    public int Column { get; }

    /// <inheritdoc />
    public bool Equals(AsciiDocSourcePosition other) =>
        Offset == other.Offset && Line == other.Line && Column == other.Column;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is AsciiDocSourcePosition other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = Offset;
            hash = (hash * 397) ^ Line;
            hash = (hash * 397) ^ Column;
            return hash;
        }
    }

    /// <inheritdoc />
    public override string ToString() => $"L{Line}:C{Column} ({Offset})";
}
