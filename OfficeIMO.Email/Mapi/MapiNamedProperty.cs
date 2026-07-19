namespace OfficeIMO.Email;

/// <summary>Canonical name associated with a mapped MAPI property ID.</summary>
public sealed class MapiNamedProperty : IEquatable<MapiNamedProperty> {
    /// <summary>Creates a numeric named property.</summary>
    public MapiNamedProperty(Guid propertySet, uint localId) {
        PropertySet = propertySet;
        LocalId = localId;
    }

    /// <summary>Creates a string named property.</summary>
    public MapiNamedProperty(Guid propertySet, string name) {
        PropertySet = propertySet;
        Name = name ?? throw new ArgumentNullException(nameof(name));
    }

    /// <summary>Property-set GUID.</summary>
    public Guid PropertySet { get; }

    /// <summary>Long ID for a numeric named property.</summary>
    public uint? LocalId { get; }

    /// <summary>Name for a string named property.</summary>
    public string? Name { get; }

    /// <summary>True for a string-named property.</summary>
    public bool IsStringNamed => Name != null;

    /// <inheritdoc />
    public bool Equals(MapiNamedProperty? other) {
        if (other == null || PropertySet != other.PropertySet || IsStringNamed != other.IsStringNamed) return false;
        return IsStringNamed
            ? string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase)
            : LocalId == other.LocalId;
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => Equals(obj as MapiNamedProperty);

    /// <inheritdoc />
    public override int GetHashCode() {
        int hash = PropertySet.GetHashCode();
        if (IsStringNamed) {
            return unchecked((hash * 397) ^ StringComparer.OrdinalIgnoreCase.GetHashCode(Name!));
        }
        return unchecked((hash * 397) ^ LocalId.GetValueOrDefault().GetHashCode());
    }

    /// <inheritdoc />
    public override string ToString() {
        return Name == null
            ? string.Concat(PropertySet.ToString("D"), ":0x", LocalId.GetValueOrDefault().ToString("X8", CultureInfo.InvariantCulture))
            : string.Concat(PropertySet.ToString("D"), ":", Name);
    }
}
