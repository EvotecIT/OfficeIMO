using System;

namespace OfficeIMO.Word;

/// <summary>The interpretation of an atomic Word preferred table or cell width.</summary>
public enum WordTableWidthKind {
    /// <summary>A physical width.</summary>
    Length,
    /// <summary>A percentage of the containing width.</summary>
    Percent,
    /// <summary>Automatic width selection.</summary>
    Auto,
    /// <summary>No preferred width; retains the native nil distinction.</summary>
    Nil
}

/// <summary>A Word preferred width expressed as a physical length, human percentage, Auto or Nil.</summary>
/// <remarks>Default is an explicit zero length. This preference is independent of actual grid geometry.
/// Lengths are quantized to twips and percentages to fiftieths of a percent when assigned.</remarks>
public readonly struct WordTableWidth : IEquatable<WordTableWidth> {
    private readonly global::OfficeIMO.OfficeLength _length;
    private readonly double _percent;
    private WordTableWidth(WordTableWidthKind kind, global::OfficeIMO.OfficeLength length, double percent) {
        Kind = kind;
        _length = length;
        _percent = percent;
    }
    /// <summary>Gets the interpretation, which cannot be mutated independently of the payload.</summary>
    public WordTableWidthKind Kind { get; }
    /// <summary>Gets the physical width, or null for a nonphysical preference.</summary>
    public global::OfficeIMO.OfficeLength? LengthValue => Kind == WordTableWidthKind.Length ? _length : null;
    /// <summary>Gets the human percentage (37 means 37%), or null for another interpretation.</summary>
    public double? PercentValue => Kind == WordTableWidthKind.Percent ? _percent : (double?)null;
    /// <summary>Creates a nonnegative physical preference within the native signed twip range.</summary>
    public static WordTableWidth Length(global::OfficeIMO.OfficeLength value) {
        if (value < global::OfficeIMO.OfficeLength.Zero || value.ToTwips() > int.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(value));
        return new WordTableWidth(WordTableWidthKind.Length, value, 0);
    }
    /// <summary>Creates a nonnegative finite percentage. Values above 100 are supported within the native range.</summary>
    public static WordTableWidth Percent(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0 || value > int.MaxValue / 50d)
            throw new ArgumentOutOfRangeException(nameof(value));
        return new WordTableWidth(WordTableWidthKind.Percent, default, value);
    }
    /// <summary>Gets automatic width selection.</summary>
    public static WordTableWidth Auto => new WordTableWidth(WordTableWidthKind.Auto, default, 0);
    /// <summary>Gets an explicit native nil preference.</summary>
    public static WordTableWidth Nil => new WordTableWidth(WordTableWidthKind.Nil, default, 0);
    internal int ToNativeValue() => Kind == WordTableWidthKind.Length ? checked((int)_length.ToTwips())
        : Kind == WordTableWidthKind.Percent ? checked((int)Math.Round(_percent * 50d, MidpointRounding.AwayFromZero)) : 0;
    /// <inheritdoc />
    public bool Equals(WordTableWidth other) => Kind == other.Kind && _length == other._length && _percent.Equals(other._percent);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is WordTableWidth other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => unchecked(((int)Kind * 397 ^ _length.GetHashCode()) * 397 ^ _percent.GetHashCode());
    /// <summary>Compares both interpretation and payload.</summary>
    public static bool operator ==(WordTableWidth left, WordTableWidth right) => left.Equals(right);
    /// <summary>Compares both interpretation and payload.</summary>
    public static bool operator !=(WordTableWidth left, WordTableWidth right) => !left.Equals(right);
}
