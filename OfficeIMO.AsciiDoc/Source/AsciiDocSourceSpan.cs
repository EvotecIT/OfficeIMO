namespace OfficeIMO.AsciiDoc;

/// <summary>
/// Half-open source range. <see cref="Start"/> is inclusive and <see cref="End"/> is exclusive.
/// </summary>
public readonly struct AsciiDocSourceSpan : IEquatable<AsciiDocSourceSpan> {
    /// <summary>Creates a source span.</summary>
    public AsciiDocSourceSpan(AsciiDocSourcePosition start, AsciiDocSourcePosition end) {
        if (end.Offset < start.Offset) throw new ArgumentException("End must not precede start.", nameof(end));
        Start = start;
        End = end;
    }

    /// <summary>Inclusive start boundary.</summary>
    public AsciiDocSourcePosition Start { get; }

    /// <summary>Exclusive end boundary.</summary>
    public AsciiDocSourcePosition End { get; }

    /// <summary>Number of UTF-16 characters in the span.</summary>
    public int Length => End.Offset - Start.Offset;

    /// <summary>Returns whether this span fully contains another half-open span.</summary>
    public bool Contains(AsciiDocSourceSpan other) =>
        Start.Offset <= other.Start.Offset && End.Offset >= other.End.Offset;

    /// <summary>Returns whether the zero-based source offset lies inside the span.</summary>
    public bool ContainsOffset(int offset) => offset >= Start.Offset && offset < End.Offset;

    /// <summary>Extracts this span from the source text.</summary>
    public string Slice(string source) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (End.Offset > source.Length) throw new ArgumentOutOfRangeException(nameof(source), "The source is shorter than the span.");
        return source.Substring(Start.Offset, Length);
    }

    /// <inheritdoc />
    public bool Equals(AsciiDocSourceSpan other) => Start.Equals(other.Start) && End.Equals(other.End);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is AsciiDocSourceSpan other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            return (Start.GetHashCode() * 397) ^ End.GetHashCode();
        }
    }

    /// <inheritdoc />
    public override string ToString() => $"{Start.Line}:{Start.Column}-{End.Line}:{End.Column}";
}
