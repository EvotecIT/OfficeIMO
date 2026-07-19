namespace OfficeIMO.Email;

/// <summary>
/// Typed access to an ordered, duplicate-preserving MAPI property collection.
/// </summary>
/// <remarks>
/// Reads use the last property with a matching canonical identity. Raw properties remain untouched until
/// <see cref="Set{T}(MapiPropertyKey{T}, T, uint?)"/> or <see cref="Remove(MapiPropertyKey)"/> is called.
/// </remarks>
public sealed class MapiPropertyBag {
    private readonly IList<MapiProperty> _properties;

    /// <summary>Creates a typed view over an existing property collection.</summary>
    public MapiPropertyBag(IList<MapiProperty> properties) {
        _properties = properties ?? throw new ArgumentNullException(nameof(properties));
    }

    /// <summary>The exact mutable raw collection backing this bag.</summary>
    public IList<MapiProperty> Properties => _properties;

    /// <summary>Finds the winning property with the key's canonical identity and an accepted wire type.</summary>
    public MapiProperty? Find(MapiPropertyKey key) {
        return _properties.GetMapiProperty(key);
    }

    /// <summary>Finds the winning raw property by canonical identity without requiring an accepted wire type.</summary>
    public MapiProperty? FindRaw(MapiPropertyKey key) {
        return _properties.GetRawMapiProperty(key);
    }

    /// <summary>Returns every property with the specified canonical identity and an accepted wire type.</summary>
    public IReadOnlyList<MapiProperty> FindAll(MapiPropertyKey key) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        return _properties.Where(key.Matches).ToArray();
    }

    /// <summary>Returns every raw property with the specified canonical identity, in source order.</summary>
    public IReadOnlyList<MapiProperty> FindAllRaw(MapiPropertyKey key) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        return _properties.Where(key.MatchesIdentity).ToArray();
    }

    /// <summary>Returns whether at least one raw property has the specified canonical identity.</summary>
    public bool Contains(MapiPropertyKey key) => FindRaw(key) != null;

    /// <summary>Attempts to read the winning property using the key's wire and managed type contract.</summary>
    public bool TryGetValue<T>(MapiPropertyKey<T> key, out T? value) {
        return _properties.TryGetMapiValue(key, out value);
    }

    /// <summary>Reads the winning typed value, or the managed default when it is absent or incompatible.</summary>
    public T? GetValueOrDefault<T>(MapiPropertyKey<T> key) {
        return TryGetValue(key, out T? value) ? value : default;
    }

    /// <summary>Reads the winning value-type property, or <see langword="null"/> when it is absent or incompatible.</summary>
    public T? GetNullableValue<T>(MapiPropertyKey<T> key) where T : struct {
        return TryGetValue(key, out T value) ? value : (T?)null;
    }

    /// <summary>
    /// Replaces every property with the same canonical identity and appends one property using the key's preferred
    /// wire type. Existing mapped property ID and flags are retained unless flags are supplied explicitly.
    /// </summary>
    public MapiProperty Set<T>(MapiPropertyKey<T> key, T value, uint? flags = null) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (value == null) throw new ArgumentNullException(nameof(value));

        return SetValue(key, value, key.PreferredType, flags);
    }

    /// <summary>Writes a value using one of the key's accepted wire types.</summary>
    public MapiProperty Set<T>(MapiPropertyKey<T> key, T value, MapiPropertyType wireType, uint? flags = null) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (value == null) throw new ArgumentNullException(nameof(value));
        return SetValue(key, value, wireType, flags);
    }

    /// <summary>Writes a null placeholder for an object or null-typed MAPI property.</summary>
    public MapiProperty SetNull(MapiPropertyKey key, MapiPropertyType wireType, uint? flags = null) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (wireType != MapiPropertyType.Object && wireType != MapiPropertyType.Null) {
            throw new ArgumentException("Only Object and Null MAPI wire types can be written without a value.",
                nameof(wireType));
        }
        ValidateWireType(key, wireType);
        return Replace(key, null, wireType, flags);
    }

    internal MapiProperty SetValue(MapiPropertyKey key, object value, uint? flags = null) =>
        SetValue(key, value, key.PreferredType, flags);

    internal MapiProperty SetValue(MapiPropertyKey key, object value, MapiPropertyType wireType,
        uint? flags = null) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (value == null) throw new ArgumentNullException(nameof(value));
        ValidateWireType(key, wireType);
        if (!key.ValueType.IsInstanceOfType(value)) {
            throw new ArgumentException(string.Concat("Property ", key.CanonicalName, " requires managed value type ",
                key.ValueType.FullName, "."), nameof(value));
        }

        return Replace(key, value, wireType, flags);
    }

    private MapiProperty Replace(MapiPropertyKey key, object? value, MapiPropertyType wireType, uint? flags) {
        MapiProperty? previous = FindRaw(key);
        ushort propertyId = key.PropertyId ?? previous?.PropertyId ?? 0x8000;
        uint propertyFlags = flags ?? previous?.Flags ?? 0x00000006;
        Remove(key);

        var property = new MapiProperty(propertyId, wireType, value, propertyFlags, key.Name);
        _properties.Add(property);
        return property;
    }

    private static void ValidateWireType(MapiPropertyKey key, MapiPropertyType wireType) {
        if (!key.Accepts(wireType)) {
            throw new ArgumentException(string.Concat("Wire type ", wireType.ToString(), " is not accepted by ",
                key.CanonicalName, "."), nameof(wireType));
        }
    }

    /// <summary>Removes all properties with the specified canonical identity.</summary>
    /// <returns>The number of removed raw properties.</returns>
    public int Remove(MapiPropertyKey key) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        int removed = 0;
        for (int index = _properties.Count - 1; index >= 0; index--) {
            if (!key.MatchesIdentity(_properties[index])) continue;
            _properties.RemoveAt(index);
            removed++;
        }
        return removed;
    }
}

internal static class MapiValueConverter {
    internal static bool TryConvert<T>(object? source, out T? value) {
        if (source is T exact) {
            value = exact;
            return true;
        }
        if (source == null) {
            value = default;
            return false;
        }

        object? converted = null;
        Type target = typeof(T);
        if (target == typeof(int)) {
            if (source is short shortValue) converted = (int)shortValue;
            else if (source is uint uintValue && uintValue <= int.MaxValue) converted = (int)uintValue;
            else if (source is long longValue && longValue >= int.MinValue && longValue <= int.MaxValue) converted = (int)longValue;
        } else if (target == typeof(long)) {
            if (source is short shortValue) converted = (long)shortValue;
            else if (source is int intValue) converted = (long)intValue;
            else if (source is uint uintValue) converted = (long)uintValue;
        } else if (target == typeof(double) && source is float floatValue) {
            converted = (double)floatValue;
        } else if (target == typeof(bool)) {
            if (source is int intValue) converted = intValue != 0;
            else if (source is short shortValue) converted = shortValue != 0;
            else if (source is uint uintValue) converted = uintValue != 0;
        } else if (target == typeof(DateTimeOffset) && source is DateTime dateValue) {
            converted = new DateTimeOffset(dateValue);
        }

        if (converted is T typed) {
            value = typed;
            return true;
        }
        value = default;
        return false;
    }
}
