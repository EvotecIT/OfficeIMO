namespace OfficeIMO.Email;

internal sealed class MsgPropertyBuilder {
    private readonly List<MapiProperty> _properties;
    private readonly MapiPropertyBag _bag;

    internal MsgPropertyBuilder(IEnumerable<MapiProperty> source) {
        _properties = source.Select(Clone).ToList();
        _bag = new MapiPropertyBag(_properties);
    }

    internal IReadOnlyList<MapiProperty> Properties => _properties;

    internal void Set(MapiPropertyKey key, object? value) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (value == null) {
            _bag.Remove(key);
            if (!key.IsNamed && (key.PreferredType == MapiPropertyType.Object ||
                key.PreferredType == MapiPropertyType.Null)) {
                _properties.Add(new MapiProperty(key.PropertyId!.Value, key.PreferredType));
            }
            return;
        }
        _bag.SetValue(key, value);
    }

    internal void Set(MapiPropertyKey key, MapiPropertyType wireType, object? value) {
        Set(key, wireType, value, null);
    }

    internal void Set(MapiPropertyKey key, MapiPropertyType wireType, object? value, uint? flags) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (value == null) {
            if (wireType == MapiPropertyType.Object || wireType == MapiPropertyType.Null) {
                _bag.SetNull(key, wireType, flags);
            } else {
                _bag.Remove(key);
            }
            return;
        }
        _bag.SetValue(key, value, wireType, flags);
    }

    internal void SetDefault(MapiPropertyKey key, object value) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (_bag.FindRaw(key) != null) return;
        _bag.SetValue(key, value);
    }

    internal void Remove(MapiPropertyKey key) => _bag.Remove(key);

    private static MapiProperty Clone(MapiProperty property) {
        return new MapiProperty(property.PropertyId, property.PropertyType, property.Value, property.Flags, property.Name) {
            RawData = property.RawData == null ? null : (byte[])property.RawData.Clone()
        };
    }
}
