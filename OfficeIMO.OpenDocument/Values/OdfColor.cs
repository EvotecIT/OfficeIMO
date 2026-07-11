namespace OfficeIMO.OpenDocument;

/// <summary>ODF RGB color serialized as `#RRGGBB`.</summary>
public readonly struct OdfColor : IEquatable<OdfColor> {
    private readonly string? _hex;

    /// <summary>Creates a color from red, green, and blue components.</summary>
    public OdfColor(byte red, byte green, byte blue) {
        _hex = "#" + red.ToString("X2", CultureInfo.InvariantCulture) + green.ToString("X2", CultureInfo.InvariantCulture) + blue.ToString("X2", CultureInfo.InvariantCulture);
    }

    private OdfColor(string hex) {
        _hex = hex;
    }

    /// <summary>Parses `#RRGGBB` or `RRGGBB`.</summary>
    public static OdfColor Parse(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        string hex = value.Trim();
        if (!hex.StartsWith("#", StringComparison.Ordinal)) hex = "#" + hex;
        if (hex.Length != 7 || !int.TryParse(hex.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)) {
            throw new FormatException("ODF colors must use RRGGBB hexadecimal notation.");
        }
        return new OdfColor(hex.ToUpperInvariant());
    }

    /// <summary>Attempts to parse <c>#RRGGBB</c> or <c>RRGGBB</c> without throwing.</summary>
    public static bool TryParse(string? value, out OdfColor color) {
        color = default;
        if (string.IsNullOrWhiteSpace(value)) return false;
        string hex = value!.Trim();
        if (!hex.StartsWith("#", StringComparison.Ordinal)) hex = "#" + hex;
        if (hex.Length != 7 || !int.TryParse(hex.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out _)) return false;
        color = new OdfColor(hex.ToUpperInvariant());
        return true;
    }

    /// <inheritdoc />
    public override string ToString() => _hex ?? "#000000";
    /// <inheritdoc />
    public bool Equals(OdfColor other) => string.Equals(ToString(), other.ToString(), StringComparison.Ordinal);
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OdfColor other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(ToString());
}
