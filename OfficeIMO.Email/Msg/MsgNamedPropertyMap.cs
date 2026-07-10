using OfficeIMO.Shared;

namespace OfficeIMO.Email;

internal sealed class MsgNamedPropertyMap {
    internal static readonly Guid PsMapi = new Guid("00020328-0000-0000-C000-000000000046");
    internal static readonly Guid PsPublicStrings = new Guid("00020329-0000-0000-C000-000000000046");
    private readonly Dictionary<ushort, MapiNamedProperty> _properties = new Dictionary<ushort, MapiNamedProperty>();

    internal MapiNamedProperty? Get(ushort propertyId) {
        MapiNamedProperty value;
        return _properties.TryGetValue(propertyId, out value!) ? value : null;
    }

    internal static MsgNamedPropertyMap Read(OfficeCompoundFile compound, IList<EmailDiagnostic> diagnostics, MsgParserState state) {
        var result = new MsgNamedPropertyMap();
        const string prefix = "__nameid_version1.0";
        if (!compound.Streams.TryGetValue(string.Concat(prefix, "/__substg1.0_00030102"), out byte[]? entries)) return result;
        compound.Streams.TryGetValue(string.Concat(prefix, "/__substg1.0_00020102"), out byte[]? guidBytes);
        compound.Streams.TryGetValue(string.Concat(prefix, "/__substg1.0_00040102"), out byte[]? stringBytes);
        guidBytes = guidBytes ?? Array.Empty<byte>();
        stringBytes = stringBytes ?? Array.Empty<byte>();
        if (guidBytes.Length % 16 != 0 || entries.Length % 8 != 0) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_NAMEID_MALFORMED",
                "The named-property mapping streams have invalid lengths.", EmailDiagnosticSeverity.Error, prefix));
        }

        int count = entries.Length / 8;
        for (int index = 0; index < count; index++) {
            state.ThrowIfCancellationRequested();
            int offset = index * 8;
            uint identifier = MsgBinary.ReadUInt32(entries, offset);
            uint indexAndKind = MsgBinary.ReadUInt32(entries, offset + 4);
            ushort propertyIndex = unchecked((ushort)indexAndKind);
            int guidIndex = (int)((indexAndKind >> 16) & 0x7fff);
            bool stringNamed = (indexAndKind & 0x80000000U) != 0;
            Guid? propertySet = ResolveGuid(guidIndex, guidBytes);
            if (!propertySet.HasValue) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_NAMEID_GUID_INVALID",
                    string.Concat("Named property 0x", (0x8000 + propertyIndex).ToString("X4", CultureInfo.InvariantCulture),
                        " references an unavailable property-set GUID."),
                    EmailDiagnosticSeverity.Warning, prefix));
                continue;
            }

            MapiNamedProperty name;
            if (stringNamed) {
                string? text = ReadName(stringBytes, identifier);
                if (text == null) {
                    diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_NAMEID_STRING_INVALID",
                        "A string-named property references an invalid string-stream offset.",
                        EmailDiagnosticSeverity.Warning, prefix));
                    continue;
                }
                name = new MapiNamedProperty(propertySet.Value, text);
            } else {
                name = new MapiNamedProperty(propertySet.Value, identifier);
            }
            result._properties[unchecked((ushort)(0x8000 + propertyIndex))] = name;
        }
        return result;
    }

    private static Guid? ResolveGuid(int guidIndex, byte[] bytes) {
        if (guidIndex == 1) return PsMapi;
        if (guidIndex == 2) return PsPublicStrings;
        int offset = checked((guidIndex - 3) * 16);
        if (guidIndex < 3 || offset < 0 || offset + 16 > bytes.Length) return null;
        return new Guid(MsgBinary.Slice(bytes, offset, 16));
    }

    private static string? ReadName(byte[] bytes, uint rawOffset) {
        if (rawOffset > int.MaxValue) return null;
        int offset = (int)rawOffset;
        if (offset < 0 || offset + 4 > bytes.Length) return null;
        uint rawLength = MsgBinary.ReadUInt32(bytes, offset);
        if (rawLength > int.MaxValue || offset + 4 > bytes.Length - (int)rawLength) return null;
        return Encoding.Unicode.GetString(bytes, offset + 4, (int)rawLength);
    }
}
