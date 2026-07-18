using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

internal static class OabPropertyValues {
    internal static MapiProperty? Find(IReadOnlyList<MapiProperty> properties, MapiPropertyKey key) =>
        properties.GetMapiProperty(key);

    internal static string? String(IReadOnlyList<MapiProperty> properties, MapiPropertyKey<string> key) =>
        Find(properties, key)?.Value as string;

    internal static IReadOnlyList<string> Strings(IReadOnlyList<MapiProperty> properties, MapiPropertyKey<string[]> key) =>
        Find(properties, key)?.Value as string[] ?? Array.Empty<string>();

    internal static IReadOnlyList<uint> UInt32s(IReadOnlyList<MapiProperty> properties, MapiPropertyKey key) =>
        Find(properties, key)?.Value as uint[] ?? Array.Empty<uint>();

    internal static uint? UInt32(IReadOnlyList<MapiProperty> properties, MapiPropertyKey key) {
        object? value = Find(properties, key)?.Value;
        if (value is uint unsigned) return unsigned;
        if (value is int signed) return unchecked((uint)signed);
        return null;
    }

    internal static bool? Boolean(IReadOnlyList<MapiProperty> properties, MapiPropertyKey<bool> key) =>
        Find(properties, key)?.Value as bool?;
}
