using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstNamedPropertyMapTests {
    [Fact]
    public void AppliesNumericStringAndCustomPropertySetMappings() {
        var customPropertySet = new Guid("4F3A6B38-21C4-4A2D-83BE-A04B8B428DE8");
        byte[] strings = CreateStringStream("OfficeIMO-Label");
        byte[] entries = new byte[24];
        WriteEntry(entries, 0, 0x820D, guidIndex: 1, stringNamed: false, propertyIndex: 0);
        WriteEntry(entries, 8, 0, guidIndex: 2, stringNamed: true, propertyIndex: 1);
        WriteEntry(entries, 16, 0x1234, guidIndex: 3, stringNamed: false, propertyIndex: 2);
        var mappingProperties = new[] {
            BinaryProperty(0x0002, customPropertySet.ToByteArray()),
            BinaryProperty(0x0003, entries),
            BinaryProperty(0x0004, strings)
        };
        var diagnostics = new List<EmailStoreDiagnostic>();

        PstNamedPropertyMap map = PstNamedPropertyMap.Read(mappingProperties, diagnostics, "nameid");
        var values = new[] {
            new MapiProperty(0x8000, MapiPropertyType.Time),
            new MapiProperty(0x8001, MapiPropertyType.Unicode),
            new MapiProperty(0x8002, MapiPropertyType.Integer32)
        };
        map.Apply(values);

        Assert.Empty(diagnostics);
        Assert.Equal(new Guid("00020328-0000-0000-C000-000000000046"), values[0].Name!.PropertySet);
        Assert.Equal((uint)0x820D, values[0].Name!.LocalId);
        Assert.Equal(new Guid("00020329-0000-0000-C000-000000000046"), values[1].Name!.PropertySet);
        Assert.Equal("OfficeIMO-Label", values[1].Name!.Name);
        Assert.Equal(customPropertySet, values[2].Name!.PropertySet);
        Assert.Equal((uint)0x1234, values[2].Name!.LocalId);
    }

    [Fact]
    public void ReportsInvalidMappingReferencesWithoutAssigningNames() {
        byte[] entries = new byte[16];
        WriteEntry(entries, 0, 7, guidIndex: 8, stringNamed: false, propertyIndex: 0);
        WriteEntry(entries, 8, 500, guidIndex: 2, stringNamed: true, propertyIndex: 1);
        var diagnostics = new List<EmailStoreDiagnostic>();

        PstNamedPropertyMap map = PstNamedPropertyMap.Read(
            new[] { BinaryProperty(0x0003, entries) }, diagnostics, "nameid");
        var values = new[] {
            new MapiProperty(0x8000, MapiPropertyType.Integer32),
            new MapiProperty(0x8001, MapiPropertyType.Unicode)
        };
        map.Apply(values);

        Assert.Equal(2, diagnostics.Count);
        Assert.All(values, property => Assert.Null(property.Name));
    }

    private static MapiProperty BinaryProperty(ushort id, byte[] value) =>
        new MapiProperty(id, MapiPropertyType.Binary, value) { RawData = value };

    private static byte[] CreateStringStream(string value) {
        byte[] text = Encoding.Unicode.GetBytes(value);
        var stream = new byte[4 + text.Length];
        WriteUInt32(stream, 0, (uint)text.Length);
        Buffer.BlockCopy(text, 0, stream, 4, text.Length);
        return stream;
    }

    private static void WriteEntry(byte[] bytes, int offset, uint identifier, int guidIndex,
        bool stringNamed, ushort propertyIndex) {
        WriteUInt32(bytes, offset, identifier);
        WriteUInt16(bytes, offset + 4, (guidIndex << 1) | (stringNamed ? 1 : 0));
        WriteUInt16(bytes, offset + 6, propertyIndex);
    }

    private static void WriteUInt16(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }
}
