using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook.Tests;

internal sealed class OabV4Fixture {
    private readonly IReadOnlyList<PropertySpec> _headerSchema;
    private readonly IReadOnlyList<PropertySpec> _entrySchema;
    private readonly Dictionary<uint, object> _headerValues;
    private readonly List<Dictionary<uint, object>> _entries = new List<Dictionary<uint, object>>();

    static OabV4Fixture() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    internal OabV4Fixture() {
        _headerSchema = new[] {
            Property(OabPropertyTags.AddressBookName, MapiPropertyType.Unicode),
            Property(OabPropertyTags.AddressBookSequence, MapiPropertyType.Integer32),
            Property(OabPropertyTags.AddressBookContainerGuid, MapiPropertyType.Unicode),
            Property(OabPropertyTags.AddressBookDistinguishedName, MapiPropertyType.Unicode)
        };
        _entrySchema = new[] {
            Property(OabPropertyTags.EmailAddress, MapiPropertyType.Unicode, flags: 3),
            Property(OabPropertyTags.SmtpAddress, MapiPropertyType.Unicode, flags: 3),
            Property(OabPropertyTags.DisplayName, MapiPropertyType.Unicode, flags: 1),
            Property(OabPropertyTags.ObjectType, MapiPropertyType.Integer32),
            Property(OabPropertyTags.Account, MapiPropertyType.String8),
            Property(OabPropertyTags.GivenName, MapiPropertyType.Unicode),
            Property(OabPropertyTags.Surname, MapiPropertyType.Unicode),
            Property(OabPropertyTags.CompanyName, MapiPropertyType.Unicode),
            Property(OabPropertyTags.Department, MapiPropertyType.Unicode),
            Property(OabPropertyTags.BusinessTelephone, MapiPropertyType.Unicode),
            Property(OabPropertyTags.ProxyAddresses, MapiPropertyType.MultipleUnicode),
            Property(OabPropertyTags.Members, MapiPropertyType.MultipleUnicode),
            Property(OabPropertyTags.MemberOf, MapiPropertyType.MultipleUnicode),
            Property(OabPropertyTags.SendRichInfo, MapiPropertyType.Boolean),
            Property(OabPropertyTags.TruncatedProperties, MapiPropertyType.MultipleInteger32),
            Property(0x8FFF, MapiPropertyType.Binary),
            Property(0x8FFE, MapiPropertyType.MultipleBinary)
        };
        _headerValues = new Dictionary<uint, object> {
            [Tag(OabPropertyTags.AddressBookName, MapiPropertyType.Unicode)] = "Synthetic Global Address List",
            [Tag(OabPropertyTags.AddressBookSequence, MapiPropertyType.Integer32)] = 42U,
            [Tag(OabPropertyTags.AddressBookContainerGuid, MapiPropertyType.Unicode)] = "55d15c71-12b6-4618-a386-7c21700a616f",
            [Tag(OabPropertyTags.AddressBookDistinguishedName, MapiPropertyType.Unicode)] = "/o=Example/ou=Exchange Administrative Group/cn=Recipients"
        };
        AddPerson("Ada Lovelace", "ada@example.test", "ada", "Ada", "Lovelace", "Research");
        AddPerson("Grace Hopper", "grace@example.test", "grace", "Grace", "Hopper", "Engineering");
        AddDistributionList();
    }

    internal OabV4Fixture AddPerson(string name, string smtp, string account,
        string givenName, string surname, string department) {
        _entries.Add(new Dictionary<uint, object> {
            [Tag(OabPropertyTags.EmailAddress, MapiPropertyType.Unicode)] = "/o=Example/ou=Recipients/cn=" + account,
            [Tag(OabPropertyTags.SmtpAddress, MapiPropertyType.Unicode)] = smtp,
            [Tag(OabPropertyTags.DisplayName, MapiPropertyType.Unicode)] = name,
            [Tag(OabPropertyTags.ObjectType, MapiPropertyType.Integer32)] = 6U,
            [Tag(OabPropertyTags.Account, MapiPropertyType.String8)] = account,
            [Tag(OabPropertyTags.GivenName, MapiPropertyType.Unicode)] = givenName,
            [Tag(OabPropertyTags.Surname, MapiPropertyType.Unicode)] = surname,
            [Tag(OabPropertyTags.CompanyName, MapiPropertyType.Unicode)] = "Example Ltd",
            [Tag(OabPropertyTags.Department, MapiPropertyType.Unicode)] = department,
            [Tag(OabPropertyTags.BusinessTelephone, MapiPropertyType.Unicode)] = "+1 555 0100",
            [Tag(OabPropertyTags.ProxyAddresses, MapiPropertyType.MultipleUnicode)] = new[] { "SMTP:" + smtp, "smtp:alias-" + smtp },
            [Tag(OabPropertyTags.MemberOf, MapiPropertyType.MultipleUnicode)] = new[] { "/o=Example/ou=Recipients/cn=all" },
            [Tag(OabPropertyTags.SendRichInfo, MapiPropertyType.Boolean)] = true,
            [Tag(OabPropertyTags.TruncatedProperties, MapiPropertyType.MultipleInteger32)] = new[] { 0x8009101FU },
            [Tag(0x8FFF, MapiPropertyType.Binary)] = new byte[] { 1, 2, 3, 4 },
            [Tag(0x8FFE, MapiPropertyType.MultipleBinary)] = new[] { new byte[] { 5 }, new byte[] { 6, 7 } }
        });
        return this;
    }

    internal OabV4Fixture RemoveEntryProperty(int entryIndex, ushort propertyId, MapiPropertyType type) {
        _entries[entryIndex].Remove(Tag(propertyId, type));
        return this;
    }

    internal byte[] Build() {
        using (var stream = new MemoryStream()) {
            WriteUInt32(stream, 0x00000020U);
            WriteUInt32(stream, 0U);
            WriteUInt32(stream, checked((uint)_entries.Count));

            using (var metadata = new MemoryStream()) {
                WritePropertyTable(metadata, _headerSchema);
                WritePropertyTable(metadata, _entrySchema);
                WriteUInt32(stream, checked((uint)(metadata.Length + 4)));
                metadata.Position = 0;
                metadata.CopyTo(stream);
            }

            WriteRecord(stream, _headerSchema, _headerValues);
            foreach (Dictionary<uint, object> entry in _entries) WriteRecord(stream, _entrySchema, entry);
            byte[] result = stream.ToArray();
            WriteUInt32(result, 4, ComputeCrc(result, 12));
            return result;
        }
    }

    internal static uint ComputeCrc(byte[] data, int offset) {
        uint crc = 0xFFFFFFFFU;
        for (int index = offset; index < data.Length; index++) {
            crc ^= data[index];
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1U) != 0 ? (crc >> 1) ^ 0xEDB88320U : crc >> 1;
        }
        return crc;
    }

    private void AddDistributionList() {
        _entries.Add(new Dictionary<uint, object> {
            [Tag(OabPropertyTags.EmailAddress, MapiPropertyType.Unicode)] = "/o=Example/ou=Recipients/cn=all",
            [Tag(OabPropertyTags.SmtpAddress, MapiPropertyType.Unicode)] = "all@example.test",
            [Tag(OabPropertyTags.DisplayName, MapiPropertyType.Unicode)] = "All Example",
            [Tag(OabPropertyTags.ObjectType, MapiPropertyType.Integer32)] = 8U,
            [Tag(OabPropertyTags.Account, MapiPropertyType.String8)] = "all",
            [Tag(OabPropertyTags.Members, MapiPropertyType.MultipleUnicode)] = new[] {
                "/o=Example/ou=Recipients/cn=ada",
                "/o=Example/ou=Recipients/cn=grace"
            },
            [Tag(OabPropertyTags.ProxyAddresses, MapiPropertyType.MultipleUnicode)] = new[] { "SMTP:all@example.test" }
        });
    }

    private static PropertySpec Property(ushort id, MapiPropertyType type, uint flags = 0) =>
        new PropertySpec(Tag(id, type), type, flags);

    private static uint Tag(ushort id, MapiPropertyType type) => ((uint)id << 16) | (ushort)type;

    private static void WritePropertyTable(Stream stream, IReadOnlyList<PropertySpec> schema) {
        WriteUInt32(stream, checked((uint)schema.Count));
        foreach (PropertySpec property in schema) {
            WriteUInt32(stream, property.Tag);
            WriteUInt32(stream, property.Flags);
        }
    }

    private static void WriteRecord(Stream destination, IReadOnlyList<PropertySpec> schema,
        IReadOnlyDictionary<uint, object> values) {
        using (var body = new MemoryStream()) {
            int presenceBytes = (schema.Count + 7) / 8;
            var presence = new byte[presenceBytes];
            for (int index = 0; index < schema.Count; index++) {
                if (values.ContainsKey(schema[index].Tag)) presence[index / 8] |= (byte)(0x80 >> (index % 8));
            }
            body.Write(presence, 0, presence.Length);
            foreach (PropertySpec property in schema) {
                if (!values.TryGetValue(property.Tag, out object? value)) continue;
                WriteValue(body, property.Type, value);
            }
            WriteUInt32(destination, checked((uint)(body.Length + 4)));
            body.Position = 0;
            body.CopyTo(destination);
        }
    }

    private static void WriteValue(Stream stream, MapiPropertyType type, object value) {
        switch (type) {
            case MapiPropertyType.Integer32:
                WriteCompactUInt32(stream, (uint)value);
                return;
            case MapiPropertyType.Boolean:
                stream.WriteByte((bool)value ? (byte)1 : (byte)0);
                return;
            case MapiPropertyType.String8:
                WriteString(stream, Encoding.GetEncoding(1252), (string)value);
                return;
            case MapiPropertyType.Unicode:
                WriteString(stream, Encoding.UTF8, (string)value);
                return;
            case MapiPropertyType.Binary:
                WriteBinary(stream, (byte[])value);
                return;
            case MapiPropertyType.MultipleInteger32:
                uint[] integers = (uint[])value;
                WriteCompactUInt32(stream, checked((uint)integers.Length));
                foreach (uint integer in integers) WriteCompactUInt32(stream, integer);
                return;
            case MapiPropertyType.MultipleString8:
                WriteStrings(stream, Encoding.GetEncoding(1252), (string[])value);
                return;
            case MapiPropertyType.MultipleUnicode:
                WriteStrings(stream, Encoding.UTF8, (string[])value);
                return;
            case MapiPropertyType.MultipleBinary:
                byte[][] binaries = (byte[][])value;
                WriteCompactUInt32(stream, checked((uint)binaries.Length));
                foreach (byte[] binary in binaries) WriteBinary(stream, binary);
                return;
            default:
                throw new InvalidOperationException("Unsupported synthetic OAB property type.");
        }
    }

    private static void WriteStrings(Stream stream, Encoding encoding, string[] values) {
        WriteCompactUInt32(stream, checked((uint)values.Length));
        foreach (string value in values) WriteString(stream, encoding, value);
    }

    private static void WriteString(Stream stream, Encoding encoding, string value) {
        byte[] bytes = encoding.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
        stream.WriteByte(0);
    }

    private static void WriteBinary(Stream stream, byte[] value) {
        WriteCompactUInt32(stream, checked((uint)value.Length));
        stream.Write(value, 0, value.Length);
    }

    private static void WriteCompactUInt32(Stream stream, uint value) {
        if (value <= 0x7F) {
            stream.WriteByte((byte)value);
            return;
        }
        int count = value <= 0xFF ? 1 : value <= 0xFFFF ? 2 : value <= 0xFFFFFF ? 3 : 4;
        stream.WriteByte((byte)(0x80 | count));
        for (int index = 0; index < count; index++) stream.WriteByte((byte)(value >> (index * 8)));
    }

    private static void WriteUInt32(Stream stream, uint value) {
        stream.WriteByte((byte)value);
        stream.WriteByte((byte)(value >> 8));
        stream.WriteByte((byte)(value >> 16));
        stream.WriteByte((byte)(value >> 24));
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)(value >> 16);
        data[offset + 3] = (byte)(value >> 24);
    }

    private sealed class PropertySpec {
        internal PropertySpec(uint tag, MapiPropertyType type, uint flags) {
            Tag = tag;
            Type = type;
            Flags = flags;
        }
        internal uint Tag { get; }
        internal MapiPropertyType Type { get; }
        internal uint Flags { get; }
    }
}
