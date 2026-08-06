using System;

namespace OfficeIMO.Drawing.Binary;

/// <summary>Represents the packed OfficeArtCOLORREF value defined by MS-ODRAW.</summary>
public readonly struct OfficeArtColorReference : IEquatable<OfficeArtColorReference> {
    /// <summary>Creates a color reference from its packed 32-bit value.</summary>
    public OfficeArtColorReference(uint value) => Value = value;

    /// <summary>Gets the packed 32-bit value.</summary>
    public uint Value { get; }

    /// <summary>Gets the red byte or scheme index.</summary>
    public byte Red => unchecked((byte)Value);

    /// <summary>Gets the green byte.</summary>
    public byte Green => unchecked((byte)(Value >> 8));

    /// <summary>Gets the blue byte.</summary>
    public byte Blue => unchecked((byte)(Value >> 16));

    /// <summary>Gets the OfficeArt color flags.</summary>
    public byte Flags => unchecked((byte)(Value >> 24));

    /// <summary>Gets whether the high byte marks this reference as ignored.</summary>
    public bool IsIgnored => Flags == 0xff;

    /// <summary>Gets whether the color resolves through the current application color scheme.</summary>
    public bool IsSchemeIndex => (Flags & 0x08) != 0;

    /// <summary>Gets whether the value uses a system color index.</summary>
    public bool IsSystemIndex => (Flags & 0x10) != 0;

    /// <summary>Gets whether the value uses a palette index.</summary>
    public bool IsPaletteIndex => (Flags & 0x01) != 0;

    /// <summary>Gets the scheme index when <see cref="IsSchemeIndex"/> is true.</summary>
    public byte? SchemeIndex => IsSchemeIndex ? Red : null;

    /// <summary>
    /// Resolves the reference to RGB. Scheme colors are delegated to
    /// <paramref name="schemeColorResolver"/>; system and palette indexes remain unresolved.
    /// </summary>
    public bool TryResolve(Func<byte, OfficeColor?>? schemeColorResolver, out OfficeColor color) {
        color = default;
        if (IsIgnored) return false;
        if (IsSchemeIndex) {
            OfficeColor? resolved = schemeColorResolver?.Invoke(Red);
            if (!resolved.HasValue) return false;
            color = resolved.Value;
            return true;
        }
        if (IsSystemIndex || IsPaletteIndex) return false;
        color = OfficeColor.FromRgb(Red, Green, Blue);
        return true;
    }

    /// <inheritdoc />
    public bool Equals(OfficeArtColorReference other) => Value == other.Value;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeArtColorReference other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => Value.GetHashCode();

    /// <inheritdoc />
    public override string ToString() => $"0x{Value:X8}";

    /// <summary>Compares two color references.</summary>
    public static bool operator ==(OfficeArtColorReference left, OfficeArtColorReference right) => left.Equals(right);

    /// <summary>Compares two color references.</summary>
    public static bool operator !=(OfficeArtColorReference left, OfficeArtColorReference right) => !left.Equals(right);
}
