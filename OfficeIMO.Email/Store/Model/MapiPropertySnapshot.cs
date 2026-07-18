using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal static class MapiPropertySnapshot {
    internal static MapiProperty Clone(MapiProperty source) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return new MapiProperty(source.PropertyId, source.PropertyType, CloneValue(source.Value),
            source.Flags, source.Name) {
            RawData = source.RawData == null ? null : (byte[])source.RawData.Clone()
        };
    }

    private static object? CloneValue(object? value) {
        if (value is byte[] bytes) return (byte[])bytes.Clone();
        if (value is object[] objects) return objects.Select(CloneValue).ToArray();
        if (value is string[] strings) return (string[])strings.Clone();
        if (value is Array array) return array.Clone();
        return value;
    }
}
