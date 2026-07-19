namespace OfficeIMO.Email;

/// <summary>Convenience lookup APIs for retained standard and named MAPI properties.</summary>
public static class MapiPropertyExtensions {
    /// <summary>Finds the last property matching a typed key's identity and accepted wire types.</summary>
    public static MapiProperty? GetMapiProperty(this IEnumerable<MapiProperty> properties, MapiPropertyKey key) {
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        if (key == null) throw new ArgumentNullException(nameof(key));
        return properties.LastOrDefault(key.Matches);
    }

    /// <summary>Finds the last property matching a typed key's identity without checking its wire type.</summary>
    public static MapiProperty? GetRawMapiProperty(this IEnumerable<MapiProperty> properties, MapiPropertyKey key) {
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        if (key == null) throw new ArgumentNullException(nameof(key));
        return properties.LastOrDefault(key.MatchesIdentity);
    }

    /// <summary>Attempts to read a property using a typed key.</summary>
    public static bool TryGetMapiValue<T>(this IEnumerable<MapiProperty> properties, MapiPropertyKey<T> key,
        out T? value) {
        MapiProperty? property = GetMapiProperty(properties, key);
        if (property == null) {
            value = default;
            return false;
        }
        return MapiValueConverter.TryConvert(property.Value, out value);
    }

    /// <summary>Reads a typed property, or the managed default when it is absent or incompatible.</summary>
    public static T? GetMapiValueOrDefault<T>(this IEnumerable<MapiProperty> properties, MapiPropertyKey<T> key) {
        return TryGetMapiValue(properties, key, out T? value) ? value : default;
    }

    /// <summary>Reads a value-type property, or null when it is absent or incompatible.</summary>
    public static T? GetNullableMapiValue<T>(this IEnumerable<MapiProperty> properties, MapiPropertyKey<T> key)
        where T : struct {
        return TryGetMapiValue(properties, key, out T value) ? value : (T?)null;
    }

    /// <summary>Finds the last standard property with the specified property identifier.</summary>
    public static MapiProperty? GetMapiProperty(this IEnumerable<MapiProperty> properties, ushort propertyId) {
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        return properties.LastOrDefault(property => property.Name == null && property.PropertyId == propertyId);
    }

    /// <summary>Finds the last numeric named property with the specified canonical identity.</summary>
    public static MapiProperty? GetMapiProperty(this IEnumerable<MapiProperty> properties,
        Guid propertySet, uint localId) {
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        return properties.LastOrDefault(property => property.Name?.PropertySet == propertySet &&
            property.Name.LocalId == localId);
    }

    /// <summary>Finds the last string named property with the specified canonical identity.</summary>
    public static MapiProperty? GetMapiProperty(this IEnumerable<MapiProperty> properties,
        Guid propertySet, string name) {
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        if (name == null) throw new ArgumentNullException(nameof(name));
        return properties.LastOrDefault(property => property.Name?.PropertySet == propertySet &&
            string.Equals(property.Name.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Gets a typed standard property value, or the default value when it is absent or has another type.</summary>
    public static T? GetMapiValue<T>(this IEnumerable<MapiProperty> properties, ushort propertyId) {
        MapiProperty? property = GetMapiProperty(properties, propertyId);
        return property?.Value is T value ? value : default;
    }

    /// <summary>Gets a typed numeric named property value, or the default value when it is absent or has another type.</summary>
    public static T? GetMapiValue<T>(this IEnumerable<MapiProperty> properties, Guid propertySet, uint localId) {
        MapiProperty? property = GetMapiProperty(properties, propertySet, localId);
        return property?.Value is T value ? value : default;
    }

    /// <summary>Gets a typed string named property value, or the default value when it is absent or has another type.</summary>
    public static T? GetMapiValue<T>(this IEnumerable<MapiProperty> properties, Guid propertySet, string name) {
        MapiProperty? property = GetMapiProperty(properties, propertySet, name);
        return property?.Value is T value ? value : default;
    }
}
