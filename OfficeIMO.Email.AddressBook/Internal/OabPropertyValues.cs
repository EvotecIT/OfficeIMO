using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

internal static class OabPropertyValues {
    internal static MapiProperty? Find(IReadOnlyList<MapiProperty> properties, ushort propertyId) =>
        properties.LastOrDefault(property => property.PropertyId == propertyId);

    internal static string? String(IReadOnlyList<MapiProperty> properties, ushort propertyId) =>
        Find(properties, propertyId)?.Value as string;

    internal static IReadOnlyList<string> Strings(IReadOnlyList<MapiProperty> properties, ushort propertyId) =>
        Find(properties, propertyId)?.Value as string[] ?? Array.Empty<string>();

    internal static IReadOnlyList<uint> UInt32s(IReadOnlyList<MapiProperty> properties, ushort propertyId) =>
        Find(properties, propertyId)?.Value as uint[] ?? Array.Empty<uint>();

    internal static uint? UInt32(IReadOnlyList<MapiProperty> properties, ushort propertyId) {
        object? value = Find(properties, propertyId)?.Value;
        if (value is uint unsigned) return unsigned;
        if (value is int signed) return unchecked((uint)signed);
        return null;
    }

    internal static bool? Boolean(IReadOnlyList<MapiProperty> properties, ushort propertyId) =>
        Find(properties, propertyId)?.Value as bool?;
}
