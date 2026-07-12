namespace OfficeIMO.Email;

internal sealed class MsgPropertyBuilder {
    private readonly List<MapiProperty> _properties;

    internal MsgPropertyBuilder(IEnumerable<MapiProperty> source) {
        _properties = source.Select(Clone).ToList();
    }

    internal IReadOnlyList<MapiProperty> Properties => _properties;

    internal void Set(ushort id, MapiPropertyType type, object? value) {
        _properties.RemoveAll(property => property.Name == null && property.PropertyId == id);
        if (value == null && type != MapiPropertyType.Object && type != MapiPropertyType.Null) return;
        _properties.Add(new MapiProperty(id, type, value));
    }

    internal void SetNamed(Guid propertySet, uint localId, MapiPropertyType type, object? value) {
        _properties.RemoveAll(property => property.Name?.PropertySet == propertySet && property.Name.LocalId == localId);
        if (value == null) return;
        _properties.Add(new MapiProperty(0x8000, type, value, name: new MapiNamedProperty(propertySet, localId)));
    }

    internal void SetNamedDefault(Guid propertySet, uint localId, MapiPropertyType type, object value) {
        if (_properties.Any(property => property.Name?.PropertySet == propertySet && property.Name.LocalId == localId)) {
            return;
        }
        _properties.Add(new MapiProperty(0x8000, type, value, name: new MapiNamedProperty(propertySet, localId)));
    }

    internal void SetNamedDefault(Guid propertySet, string name, MapiPropertyType type, object value) {
        if (_properties.Any(property => property.Name?.PropertySet == propertySet &&
            string.Equals(property.Name.Name, name, StringComparison.OrdinalIgnoreCase))) {
            return;
        }
        _properties.Add(new MapiProperty(0x8000, type, value, name: new MapiNamedProperty(propertySet, name)));
    }

    internal void SetNamed(Guid propertySet, string name, MapiPropertyType type, object? value) {
        _properties.RemoveAll(property => property.Name?.PropertySet == propertySet &&
            string.Equals(property.Name.Name, name, StringComparison.OrdinalIgnoreCase));
        if (value == null) return;
        _properties.Add(new MapiProperty(0x8000, type, value, name: new MapiNamedProperty(propertySet, name)));
    }

    internal void Remove(ushort id) {
        _properties.RemoveAll(property => property.Name == null && property.PropertyId == id);
    }

    private static MapiProperty Clone(MapiProperty property) {
        return new MapiProperty(property.PropertyId, property.PropertyType, property.Value, property.Flags, property.Name) {
            RawData = property.RawData == null ? null : (byte[])property.RawData.Clone()
        };
    }
}
