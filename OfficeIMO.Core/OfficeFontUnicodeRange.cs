using System;

namespace OfficeIMO.Drawing;

/// <summary>Inclusive Unicode scalar range used to constrain a scoped font face.</summary>
public readonly struct OfficeFontUnicodeRange : IEquatable<OfficeFontUnicodeRange> {
    /// <summary>Creates an inclusive Unicode scalar range.</summary>
    public OfficeFontUnicodeRange(int start, int end) {
        if (start < 0 || start > 0x10FFFF) throw new ArgumentOutOfRangeException(nameof(start));
        if (end < start || end > 0x10FFFF) throw new ArgumentOutOfRangeException(nameof(end));
        Start = start;
        End = end;
    }

    /// <summary>First included Unicode scalar.</summary>
    public int Start { get; }

    /// <summary>Last included Unicode scalar.</summary>
    public int End { get; }

    /// <summary>Returns true when the scalar is in this range.</summary>
    public bool Contains(int scalar) => scalar >= Start && scalar <= End;

    /// <inheritdoc />
    public bool Equals(OfficeFontUnicodeRange other) => Start == other.Start && End == other.End;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeFontUnicodeRange other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (Start * 397) ^ End;
}
