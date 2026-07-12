namespace OfficeIMO.Email;

/// <summary>Convenience lookup APIs for retained standard and named MAPI properties.</summary>
public static class MapiPropertyExtensions {
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
