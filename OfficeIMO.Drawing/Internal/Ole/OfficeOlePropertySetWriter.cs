using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Writes OLE property-set streams used by legacy Office binary formats.
    /// </summary>
    internal static class OfficeOlePropertySetWriter {
        internal const string SummaryInformationStreamName = "\u0005SummaryInformation";
        internal const string DocumentSummaryInformationStreamName = "\u0005DocumentSummaryInformation";
        internal const uint CodePagePropertyId = 1;
        internal static readonly Guid SummaryInformationFormatId = new Guid("F29F85E0-4FF9-1068-AB91-08002B27B3D9");
        internal static readonly Guid DocumentSummaryInformationFormatId = new Guid("D5CDD502-2E9C-101B-9397-08002B2CF9AE");
        internal static readonly Guid UserDefinedPropertiesFormatId = new Guid("D5CDD505-2E9C-101B-9397-08002B2CF9AE");

        internal static byte[] CreatePropertySet(params (Guid FormatId, byte[] Section)[] sections) {
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

        internal static byte[] CreateSection(IReadOnlyList<OfficeOleProperty> properties) {
            using var values = new MemoryStream();
            var offsets = new List<uint>(properties.Count);
            foreach (OfficeOleProperty property in properties) {
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

        internal static void AddString(List<OfficeOleProperty> properties, uint propertyId, string? value) {
            if (!string.IsNullOrEmpty(value)) {
                properties.Add(OfficeOleProperty.String(propertyId, value!));
            }
        }

        internal static void AddFileTime(List<OfficeOleProperty> properties, uint propertyId, DateTime? value) {
            if (value.HasValue) {
                properties.Add(OfficeOleProperty.FileTime(propertyId, value.Value));
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

        internal static void WriteUInt64(Stream stream, ulong value) {
            WriteUInt32(stream, unchecked((uint)(value & 0xffffffffUL)));
            WriteUInt32(stream, unchecked((uint)(value >> 32)));
        }

        internal static void PadToInt32(Stream stream) {
            while (stream.Position % 4 != 0) {
                stream.WriteByte(0);
            }
        }
    }

    /// <summary>
    /// Typed OLE property value ready to serialize into a property-set section.
    /// </summary>
    internal readonly struct OfficeOleProperty {
        private OfficeOleProperty(uint propertyId, byte[] valueBytes) {
            PropertyId = propertyId;
            ValueBytes = valueBytes;
        }

        internal uint PropertyId { get; }

        internal byte[] ValueBytes { get; }

        internal static OfficeOleProperty Integer(uint id, object? value) {
            switch (value) {
                case sbyte sbyteValue:
                    return Int8(id, sbyteValue);
                case byte byteValue:
                    return UInt8(id, byteValue);
                case short shortValue:
                    return Int16(id, shortValue);
                case ushort ushortValue:
                    return UInt16(id, ushortValue);
                case int intValue:
                    return Int32(id, intValue);
                case uint uintValue:
                    return UInt32(id, uintValue);
                case long longValue:
                    return longValue >= int.MinValue && longValue <= int.MaxValue
                        ? Int32(id, unchecked((int)longValue))
                        : Int64(id, longValue);
                case ulong ulongValue:
                    return UInt64(id, ulongValue);
                default:
                    long integer = Convert.ToInt64(value, CultureInfo.InvariantCulture);
                    return integer >= int.MinValue && integer <= int.MaxValue
                        ? Int32(id, unchecked((int)integer))
                        : Int64(id, integer);
            }
        }

        internal static OfficeOleProperty Int16(uint id, short value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0002);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, unchecked((ushort)value));
            WriteUInt16(stream, 0);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty Int32(uint id, int value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0003);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, unchecked((uint)value));
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty Int8(uint id, sbyte value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0010);
            WriteUInt16(stream, 0);
            stream.WriteByte(unchecked((byte)value));
            OfficeOlePropertySetWriter.PadToInt32(stream);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty UInt8(uint id, byte value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0011);
            WriteUInt16(stream, 0);
            stream.WriteByte(value);
            OfficeOlePropertySetWriter.PadToInt32(stream);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty UInt16(uint id, ushort value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0012);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, value);
            WriteUInt16(stream, 0);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty UInt32(uint id, uint value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0013);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, value);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty Int64(uint id, long value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0014);
            WriteUInt16(stream, 0);
            OfficeOlePropertySetWriter.WriteUInt64(stream, unchecked((ulong)value));
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty UInt64(uint id, ulong value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0015);
            WriteUInt16(stream, 0);
            OfficeOlePropertySetWriter.WriteUInt64(stream, value);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty Double(uint id, double value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0005);
            WriteUInt16(stream, 0);
            byte[] bytes = BitConverter.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty Float(uint id, float value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0004);
            WriteUInt16(stream, 0);
            byte[] bytes = BitConverter.GetBytes(value);
            stream.Write(bytes, 0, bytes.Length);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty Boolean(uint id, bool value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x000b);
            WriteUInt16(stream, 0);
            WriteUInt16(stream, value ? (ushort)0xffff : (ushort)0);
            WriteUInt16(stream, 0);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty FileTime(uint id, DateTime value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0040);
            WriteUInt16(stream, 0);
            OfficeOlePropertySetWriter.WriteUInt64(stream, unchecked((ulong)value.ToUniversalTime().ToFileTimeUtc()));
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty FileTimeDuration(uint id,
            TimeSpan value) {
            if (value < TimeSpan.Zero) {
                throw new ArgumentOutOfRangeException(nameof(value));
            }
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0040);
            WriteUInt16(stream, 0);
            OfficeOlePropertySetWriter.WriteUInt64(stream,
                checked((ulong)value.Ticks));
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty Blob(uint id, byte[] value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x0041);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, checked((uint)value.Length));
            stream.Write(value, 0, value.Length);
            OfficeOlePropertySetWriter.PadToInt32(stream);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty String(uint id, string value) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x001f);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, checked((uint)(value.Length + 1)));
            byte[] bytes = Encoding.Unicode.GetBytes(value + '\0');
            stream.Write(bytes, 0, bytes.Length);
            OfficeOlePropertySetWriter.PadToInt32(stream);
            return new OfficeOleProperty(id, stream.ToArray());
        }

        internal static OfficeOleProperty Dictionary(uint id, IReadOnlyDictionary<uint, string> names) {
            using var stream = new MemoryStream();
            WriteUInt32(stream, checked((uint)names.Count));
            foreach (KeyValuePair<uint, string> name in names.OrderBy(entry => entry.Key)) {
                WriteUInt32(stream, name.Key);
                WriteUInt32(stream, checked((uint)(name.Value.Length + 1)));
                byte[] bytes = Encoding.Unicode.GetBytes(name.Value + '\0');
                stream.Write(bytes, 0, bytes.Length);
                OfficeOlePropertySetWriter.PadToInt32(stream);
            }

            return new OfficeOleProperty(id, stream.ToArray());
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
    }
}
