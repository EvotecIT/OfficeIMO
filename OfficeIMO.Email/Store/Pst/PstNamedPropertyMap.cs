using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstNamedPropertyMap {
    private static readonly Guid PsMapi = new Guid("00020328-0000-0000-C000-000000000046");
    private static readonly Guid PsPublicStrings = new Guid("00020329-0000-0000-C000-000000000046");
    private readonly Dictionary<ushort, MapiNamedProperty> _properties =
        new Dictionary<ushort, MapiNamedProperty>();

    internal static PstNamedPropertyMap Empty { get; } = new PstNamedPropertyMap();

    internal void Apply(IEnumerable<MapiProperty> properties) {
        foreach (MapiProperty property in properties) {
            if (property.PropertyId >= 0x8000 &&
                _properties.TryGetValue(property.PropertyId, out MapiNamedProperty? name)) {
                property.Name = name;
            }
        }
    }

    internal bool TryGetPropertyId(Guid propertySet, uint localId, out ushort propertyId) {
        foreach (KeyValuePair<ushort, MapiNamedProperty> pair in _properties) {
            if (pair.Value.PropertySet == propertySet && pair.Value.LocalId == localId) {
                propertyId = pair.Key;
                return true;
            }
        }
        propertyId = 0;
        return false;
    }

    internal bool TryGetPropertyId(MapiPropertyKey key, out ushort propertyId) {
        if (key == null) throw new ArgumentNullException(nameof(key));
        if (key.Name == null) {
            propertyId = 0;
            return false;
        }
        foreach (KeyValuePair<ushort, MapiNamedProperty> pair in _properties) {
            if (key.Name.Equals(pair.Value)) {
                propertyId = pair.Key;
                return true;
            }
        }
        propertyId = 0;
        return false;
    }

    internal static PstNamedPropertyMap Read(IEnumerable<MapiProperty> properties,
        IList<EmailStoreDiagnostic> diagnostics, string location) {
        var result = new PstNamedPropertyMap();
        byte[] entries = GetBytes(properties, MapiKnownProperties.PidTag.NameidStreamEntry);
        byte[] guids = GetBytes(properties, MapiKnownProperties.PidTag.NameidStreamGuid);
        byte[] strings = GetBytes(properties, MapiKnownProperties.PidTag.NameidStreamString);
        if (entries.Length == 0) return result;

        if (entries.Length % 8 != 0 || guids.Length % 16 != 0) {
            diagnostics.Add(new EmailStoreDiagnostic(
                "EMAIL_STORE_PST_NAMEID_MALFORMED",
                "The PST named-property mapping streams have invalid lengths.",
                EmailStoreDiagnosticSeverity.Error,
                location));
        }

        int count = entries.Length / 8;
        for (int index = 0; index < count; index++) {
            int offset = index * 8;
            uint identifier = PstBinary.UInt32(entries, offset);
            ushort guidAndKind = PstBinary.UInt16(entries, offset + 4);
            ushort propertyIndex = PstBinary.UInt16(entries, offset + 6);
            if (propertyIndex > 0x7FFF) {
                diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_NAMEID_PROPERTY_INDEX_INVALID",
                    string.Concat("Named-property index ",
                        propertyIndex.ToString(CultureInfo.InvariantCulture),
                        " exceeds the 0x7FFF mapping range and was ignored."),
                    EmailStoreDiagnosticSeverity.Warning,
                    location));
                continue;
            }

            int guidIndex = guidAndKind >> 1;
            bool stringNamed = (guidAndKind & 0x0001) != 0;
            Guid? propertySet = ResolveGuid(guidIndex, guids, stringNamed);
            if (!propertySet.HasValue) {
                diagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_NAMEID_GUID_INVALID",
                    string.Concat("Named property 0x",
                        (0x8000 + propertyIndex).ToString("X4", CultureInfo.InvariantCulture),
                        " references an unavailable property-set GUID."),
                    EmailStoreDiagnosticSeverity.Warning,
                    location));
                continue;
            }

            MapiNamedProperty name;
            if (stringNamed) {
                string? text = ReadName(strings, identifier);
                if (text == null) {
                    diagnostics.Add(new EmailStoreDiagnostic(
                        "EMAIL_STORE_PST_NAMEID_STRING_INVALID",
                        "A string-named property references an invalid string-stream offset.",
                        EmailStoreDiagnosticSeverity.Warning,
                        location));
                    continue;
                }
                name = new MapiNamedProperty(propertySet.Value, text);
            } else {
                name = new MapiNamedProperty(propertySet.Value, identifier);
            }
            result._properties[checked((ushort)(0x8000 + propertyIndex))] = name;
        }
        return result;
    }

    private static byte[] GetBytes(IEnumerable<MapiProperty> properties, MapiPropertyKey<byte[]> key) {
        MapiProperty? property = properties.GetMapiProperty(key);
        return property?.Value as byte[] ?? property?.RawData ?? Array.Empty<byte>();
    }

    private static Guid? ResolveGuid(int guidIndex, byte[] bytes, bool stringNamed) {
        if (guidIndex == 0) return stringNamed ? Guid.Empty : (Guid?)null;
        if (guidIndex == 1) return PsMapi;
        if (guidIndex == 2) return PsPublicStrings;
        int offset = checked((guidIndex - 3) * 16);
        if (guidIndex < 3 || offset < 0 || offset + 16 > bytes.Length) return null;
        var value = new byte[16];
        Buffer.BlockCopy(bytes, offset, value, 0, value.Length);
        return new Guid(value);
    }

    private static string? ReadName(byte[] bytes, uint rawOffset) {
        if (rawOffset > int.MaxValue) return null;
        int offset = (int)rawOffset;
        if (offset < 0 || offset + 4 > bytes.Length) return null;
        uint rawLength = PstBinary.UInt32(bytes, offset);
        if (rawLength > int.MaxValue || (rawLength & 1) != 0 ||
            offset + 4 > bytes.Length - (int)rawLength) return null;
        return Encoding.Unicode.GetString(bytes, offset + 4, (int)rawLength);
    }
}
