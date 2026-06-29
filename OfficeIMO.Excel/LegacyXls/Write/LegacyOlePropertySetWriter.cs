using System.Globalization;
using System.Text;
using OfficeIMO.Shared;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyOlePropertySetWriter {
        private const string SummaryInformationStreamName = "\u0005SummaryInformation";
        private const string DocumentSummaryInformationStreamName = "\u0005DocumentSummaryInformation";
        private const uint CodePagePropertyId = 1;
        private static readonly Guid SummaryInformationFormatId = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
        private static readonly Guid DocumentSummaryInformationFormatId = new Guid("D5CDD502-2E9C-101B-9397-08002B2CF9AE");
        private static readonly Guid UserDefinedPropertiesFormatId = new Guid("D5CDD505-2E9C-101B-9397-08002B2CF9AE");

        internal static IReadOnlyList<OfficeCompoundStream> CreateDocumentPropertyStreams(ExcelDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var streams = new List<OfficeCompoundStream>(2);
            byte[]? summaryInformation = CreateSummaryInformation(document);
            if (summaryInformation != null) {
                streams.Add(new OfficeCompoundStream(SummaryInformationStreamName, summaryInformation));
            }

            byte[]? documentSummaryInformation = CreateDocumentSummaryInformation(document);
            if (documentSummaryInformation != null) {
                streams.Add(new OfficeCompoundStream(DocumentSummaryInformationStreamName, documentSummaryInformation));
            }

            return streams;
        }

        private static byte[]? CreateSummaryInformation(ExcelDocument document) {
            BuiltinDocumentProperties properties = document.BuiltinDocumentProperties;
            var oleProperties = new List<OleProperty> {
                OleProperty.Int16(CodePagePropertyId, 1200)
            };
            AddString(oleProperties, 2, properties.Title);
            AddString(oleProperties, 3, properties.Subject);
            AddString(oleProperties, 4, properties.Creator);
            AddString(oleProperties, 5, properties.Keywords);
            AddString(oleProperties, 6, properties.Description);
            AddString(oleProperties, 8, properties.LastModifiedBy);
            AddString(oleProperties, 9, properties.Revision);
            AddFileTime(oleProperties, 11, properties.LastPrinted);
            AddFileTime(oleProperties, 12, properties.Created);
            AddFileTime(oleProperties, 13, properties.Modified);

            return oleProperties.Count == 1
                ? null
                : CreatePropertySet((SummaryInformationFormatId, CreateSection(oleProperties)));
        }

        private static byte[]? CreateDocumentSummaryInformation(ExcelDocument document) {
            var sections = new List<(Guid FormatId, byte[] Section)>();
            var documentSummaryProperties = new List<OleProperty> {
                OleProperty.Int16(CodePagePropertyId, 1200)
            };
            AddString(documentSummaryProperties, 2, document.BuiltinDocumentProperties.Category);
            AddString(documentSummaryProperties, 14, document.ApplicationProperties.Manager);
            AddString(documentSummaryProperties, 15, document.ApplicationProperties.Company);
            if (documentSummaryProperties.Count > 1) {
                sections.Add((DocumentSummaryInformationFormatId, CreateSection(documentSummaryProperties)));
            }

            if (document.CustomDocumentProperties.Count > 0) {
                var customProperties = new List<OleProperty> {
                    OleProperty.Int16(CodePagePropertyId, 1200)
                };
                var dictionary = new Dictionary<uint, string>();
                uint propertyId = 2;
                foreach (KeyValuePair<string, ExcelCustomProperty> pair in document.CustomDocumentProperties.OrderBy(property => property.Key, StringComparer.OrdinalIgnoreCase)) {
                    if (TryCreateCustomProperty(propertyId, pair.Value, out OleProperty property)) {
                        dictionary[propertyId] = pair.Key;
                        customProperties.Add(property);
                        propertyId++;
                    }
                }

                if (dictionary.Count > 0) {
                    customProperties.Insert(1, OleProperty.Dictionary(0, dictionary));
                    sections.Add((UserDefinedPropertiesFormatId, CreateSection(customProperties)));
                }
            }

            return sections.Count == 0 ? null : CreatePropertySet(sections.ToArray());
        }

        private static bool TryCreateCustomProperty(uint propertyId, ExcelCustomProperty customProperty, out OleProperty property) {
            object? value = customProperty.Value;
            switch (customProperty.PropertyType) {
                case ExcelCustomPropertyType.DateTime:
                    property = OleProperty.FileTime(propertyId, Convert.ToDateTime(value, CultureInfo.InvariantCulture));
                    return true;
                case ExcelCustomPropertyType.NumberInteger:
                    property = CreateIntegerProperty(propertyId, value);
                    return true;
                case ExcelCustomPropertyType.NumberDouble:
                    property = OleProperty.Double(propertyId, Convert.ToDouble(value, CultureInfo.InvariantCulture));
                    return true;
                case ExcelCustomPropertyType.YesNo:
                    property = OleProperty.Boolean(propertyId, Convert.ToBoolean(value, CultureInfo.InvariantCulture));
                    return true;
                case ExcelCustomPropertyType.Binary:
                    property = OleProperty.Blob(propertyId, GetBinaryValue(value));
                    return true;
                default:
                    property = OleProperty.String(propertyId, Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                    return true;
            }
        }

        private static byte[] GetBinaryValue(object? value) {
            return value switch {
                null => Array.Empty<byte>(),
                byte[] bytes => (byte[])bytes.Clone(),
                _ => Convert.FromBase64String(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
            };
        }

        private static OleProperty CreateIntegerProperty(uint propertyId, object? value) {
            switch (value) {
                case sbyte sbyteValue:
                    return OleProperty.Int8(propertyId, sbyteValue);
                case byte byteValue:
                    return OleProperty.UInt8(propertyId, byteValue);
                case short shortValue:
                    return OleProperty.Int16(propertyId, shortValue);
                case ushort ushortValue:
                    return OleProperty.UInt16(propertyId, ushortValue);
                case int intValue:
                    return OleProperty.Int32(propertyId, intValue);
                case uint uintValue:
                    return OleProperty.UInt32(propertyId, uintValue);
                case long longValue:
                    return longValue >= int.MinValue && longValue <= int.MaxValue
                        ? OleProperty.Int32(propertyId, unchecked((int)longValue))
                        : OleProperty.Int64(propertyId, longValue);
                case ulong ulongValue:
                    return OleProperty.UInt64(propertyId, ulongValue);
                default:
                    long integer = Convert.ToInt64(value, CultureInfo.InvariantCulture);
                    return integer >= int.MinValue && integer <= int.MaxValue
                        ? OleProperty.Int32(propertyId, unchecked((int)integer))
                        : OleProperty.Int64(propertyId, integer);
            }
        }

        private static void AddString(List<OleProperty> properties, uint propertyId, string? value) {
            if (!string.IsNullOrEmpty(value)) {
                properties.Add(OleProperty.String(propertyId, value!));
            }
        }

        private static void AddFileTime(List<OleProperty> properties, uint propertyId, DateTime? value) {
            if (value.HasValue) {
                properties.Add(OleProperty.FileTime(propertyId, value.Value));
            }
        }

        private static byte[] CreatePropertySet(params (Guid FormatId, byte[] Section)[] sections) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0xfffe);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            stream.Write(new byte[16], 0, 16);
            WriteUInt32(stream, checked((uint)sections.Length));

            int sectionOffset = 28 + sections.Length * 20;
            foreach ((Guid formatId, byte[] section) in sections) {
                byte[] formatIdBytes = formatId.ToByteArray();
                stream.Write(formatIdBytes, 0, formatIdBytes.Length);
                WriteUInt32(stream, checked((uint)sectionOffset));
                sectionOffset += section.Length;
            }

            foreach ((Guid _, byte[] section) in sections) {
                stream.Write(section, 0, section.Length);
            }

            return stream.ToArray();
        }

        private static byte[] CreateSection(IReadOnlyList<OleProperty> properties) {
            using var values = new MemoryStream();
            var offsets = new List<uint>(properties.Count);
            foreach (OleProperty property in properties) {
                offsets.Add(checked((uint)(8 + properties.Count * 8 + values.Length)));
                values.Write(property.ValueBytes, 0, property.ValueBytes.Length);
                PadToInt32(values);
            }

            using var stream = new MemoryStream();
            WriteUInt32(stream, checked((uint)(8 + properties.Count * 8 + values.Length)));
            WriteUInt32(stream, checked((uint)properties.Count));
            for (int i = 0; i < properties.Count; i++) {
                WriteUInt32(stream, properties[i].PropertyId);
                WriteUInt32(stream, offsets[i]);
            }

            byte[] valueBytes = values.ToArray();
            stream.Write(valueBytes, 0, valueBytes.Length);
            return stream.ToArray();
        }

        private static void PadToInt32(Stream stream) {
            while (stream.Position % 4 != 0) {
                stream.WriteByte(0);
            }
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 24) & 0xff));
        }

        private static void WriteUInt64(Stream stream, ulong value) {
            WriteUInt32(stream, unchecked((uint)(value & 0xffffffffUL)));
            WriteUInt32(stream, unchecked((uint)(value >> 32)));
        }

        private readonly struct OleProperty {
            private OleProperty(uint propertyId, byte[] valueBytes) {
                PropertyId = propertyId;
                ValueBytes = valueBytes;
            }

            internal uint PropertyId { get; }

            internal byte[] ValueBytes { get; }

            internal static OleProperty Int16(uint id, short value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0002);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, unchecked((ushort)value));
                WriteUInt16(stream, 0);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty Int32(uint id, int value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0003);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, unchecked((uint)value));
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty Int8(uint id, sbyte value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0010);
                WriteUInt16(stream, 0);
                stream.WriteByte(unchecked((byte)value));
                PadToInt32(stream);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty UInt8(uint id, byte value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0011);
                WriteUInt16(stream, 0);
                stream.WriteByte(value);
                PadToInt32(stream);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty UInt16(uint id, ushort value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0012);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, value);
                WriteUInt16(stream, 0);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty UInt32(uint id, uint value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0013);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, value);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty Int64(uint id, long value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0014);
                WriteUInt16(stream, 0);
                WriteUInt64(stream, unchecked((ulong)value));
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty UInt64(uint id, ulong value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0015);
                WriteUInt16(stream, 0);
                WriteUInt64(stream, value);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty Double(uint id, double value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0005);
                WriteUInt16(stream, 0);
                byte[] bytes = BitConverter.GetBytes(value);
                stream.Write(bytes, 0, bytes.Length);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty Boolean(uint id, bool value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x000b);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, value ? (ushort)0xffff : (ushort)0);
                WriteUInt16(stream, 0);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty FileTime(uint id, DateTime value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0040);
                WriteUInt16(stream, 0);
                WriteUInt64(stream, unchecked((ulong)value.ToUniversalTime().ToFileTimeUtc()));
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty Blob(uint id, byte[] value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0041);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, checked((uint)value.Length));
                stream.Write(value, 0, value.Length);
                PadToInt32(stream);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty String(uint id, string value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x001f);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, checked((uint)(value.Length + 1)));
                byte[] bytes = Encoding.Unicode.GetBytes(value + '\0');
                stream.Write(bytes, 0, bytes.Length);
                PadToInt32(stream);
                return new OleProperty(id, stream.ToArray());
            }

            internal static OleProperty Dictionary(uint id, IReadOnlyDictionary<uint, string> names) {
                using var stream = new MemoryStream();
                WriteUInt32(stream, checked((uint)names.Count));
                foreach (KeyValuePair<uint, string> name in names.OrderBy(entry => entry.Key)) {
                    WriteUInt32(stream, name.Key);
                    WriteUInt32(stream, checked((uint)(name.Value.Length + 1)));
                    byte[] bytes = Encoding.Unicode.GetBytes(name.Value + '\0');
                    stream.Write(bytes, 0, bytes.Length);
                    PadToInt32(stream);
                }

                return new OleProperty(id, stream.ToArray());
            }
        }
    }
}
