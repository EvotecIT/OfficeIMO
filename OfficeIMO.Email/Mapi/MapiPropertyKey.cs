namespace OfficeIMO.Email;

/// <summary>
/// Describes the canonical identity and supported wire types of a standard or named MAPI property.
/// </summary>
public class MapiPropertyKey {
    private readonly MapiPropertyType[] _acceptedTypes;

    /// <summary>Creates a standard-property key.</summary>
    protected MapiPropertyKey(string canonicalName, ushort propertyId, MapiPropertyType preferredType,
        params MapiPropertyType[] compatibleTypes) {
        CanonicalName = ValidateCanonicalName(canonicalName);
        PropertyId = propertyId;
        PreferredType = preferredType;
        _acceptedTypes = BuildAcceptedTypes(preferredType, compatibleTypes);
    }

    /// <summary>Creates a numeric named-property key.</summary>
    protected MapiPropertyKey(string canonicalName, Guid propertySet, uint localId,
        MapiPropertyType preferredType, params MapiPropertyType[] compatibleTypes) {
        CanonicalName = ValidateCanonicalName(canonicalName);
        Name = new MapiNamedProperty(propertySet, localId);
        PreferredType = preferredType;
        _acceptedTypes = BuildAcceptedTypes(preferredType, compatibleTypes);
    }

    /// <summary>Creates a string named-property key.</summary>
    protected MapiPropertyKey(string canonicalName, Guid propertySet, string name,
        MapiPropertyType preferredType, params MapiPropertyType[] compatibleTypes) {
        CanonicalName = ValidateCanonicalName(canonicalName);
        Name = new MapiNamedProperty(propertySet, name);
        PreferredType = preferredType;
        _acceptedTypes = BuildAcceptedTypes(preferredType, compatibleTypes);
    }

    /// <summary>Microsoft canonical name or a stable application-defined name.</summary>
    public string CanonicalName { get; }

    /// <summary>Standard property identifier, or <see langword="null"/> for a named property.</summary>
    public ushort? PropertyId { get; }

    /// <summary>Named-property identity, or <see langword="null"/> for a standard property.</summary>
    public MapiNamedProperty? Name { get; }

    /// <summary>Preferred wire type used when a new value is written.</summary>
    public MapiPropertyType PreferredType { get; }

    /// <summary>Wire types accepted when reading the property.</summary>
    public IReadOnlyList<MapiPropertyType> AcceptedTypes => _acceptedTypes;

    /// <summary>True when this key identifies a named property.</summary>
    public bool IsNamed => Name != null;

    /// <summary>Managed value type exposed by a typed key.</summary>
    public virtual Type ValueType => typeof(object);

    /// <summary>Returns whether a property has this key's canonical identity, without checking its wire type.</summary>
    public bool MatchesIdentity(MapiProperty property) {
        if (property == null) throw new ArgumentNullException(nameof(property));
        if (Name == null) return property.Name == null && property.PropertyId == PropertyId;
        return property.Name != null && Name.Equals(property.Name);
    }

    /// <summary>Returns whether a standard property identifier has this key's canonical identity.</summary>
    public bool MatchesIdentity(ushort propertyId) => Name == null && PropertyId == propertyId;

    /// <summary>Returns the standard property identifier, or throws when this key identifies a named property.</summary>
    public ushort GetStandardPropertyId() {
        if (PropertyId.HasValue) return PropertyId.Value;
        throw new InvalidOperationException(string.Concat(CanonicalName, " is a named MAPI property."));
    }

    /// <summary>Returns whether a property has this key's canonical identity and an accepted wire type.</summary>
    public bool Matches(MapiProperty property) => MatchesIdentity(property) && Accepts(property.PropertyType);

    /// <summary>Returns whether a wire type is accepted by this key.</summary>
    public bool Accepts(MapiPropertyType propertyType) => Array.IndexOf(_acceptedTypes, propertyType) >= 0;

    /// <inheritdoc />
    public override string ToString() => CanonicalName;

    private static string ValidateCanonicalName(string canonicalName) {
        if (canonicalName == null) throw new ArgumentNullException(nameof(canonicalName));
        if (string.IsNullOrWhiteSpace(canonicalName)) {
            throw new ArgumentException("A MAPI property key requires a canonical name.", nameof(canonicalName));
        }
        return canonicalName;
    }

    private static MapiPropertyType[] BuildAcceptedTypes(MapiPropertyType preferredType,
        MapiPropertyType[]? compatibleTypes) {
        var result = new List<MapiPropertyType> { preferredType };
        if (compatibleTypes != null) {
            foreach (MapiPropertyType compatibleType in compatibleTypes) {
                if (!result.Contains(compatibleType)) result.Add(compatibleType);
            }
        }
        return result.ToArray();
    }
}

/// <summary>
/// Describes a standard or named MAPI property with a managed value type.
/// </summary>
/// <typeparam name="T">Managed value type returned by the property bag.</typeparam>
public sealed class MapiPropertyKey<T> : MapiPropertyKey {
    /// <summary>Creates a standard-property key.</summary>
    public MapiPropertyKey(string canonicalName, ushort propertyId, MapiPropertyType preferredType,
        params MapiPropertyType[] compatibleTypes)
        : base(canonicalName, propertyId, preferredType, compatibleTypes) { }

    /// <summary>Creates a numeric named-property key.</summary>
    public MapiPropertyKey(string canonicalName, Guid propertySet, uint localId,
        MapiPropertyType preferredType, params MapiPropertyType[] compatibleTypes)
        : base(canonicalName, propertySet, localId, preferredType, compatibleTypes) { }

    /// <summary>Creates a string named-property key.</summary>
    public MapiPropertyKey(string canonicalName, Guid propertySet, string name,
        MapiPropertyType preferredType, params MapiPropertyType[] compatibleTypes)
        : base(canonicalName, propertySet, name, preferredType, compatibleTypes) { }

    /// <inheritdoc />
    public override Type ValueType => typeof(T);
}
