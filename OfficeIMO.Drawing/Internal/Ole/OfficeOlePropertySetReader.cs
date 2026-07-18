using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing.Internal {
    /// <summary>
    /// Reads OLE property set streams such as SummaryInformation and DocumentSummaryInformation.
    /// </summary>
    internal static class OfficeOlePropertySetReader {
        private const uint PropertyDictionaryId = 0;
        private const uint CodePagePropertyId = 1;

        internal static IReadOnlyList<OfficeOlePropertySection> ReadSections(byte[] bytes) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (bytes.Length < 28) {
                throw new InvalidDataException("The OLE property set stream is too short.");
            }

            ushort byteOrder = ReadUInt16(bytes, 0);
            if (byteOrder != 0xfffe) {
                throw new InvalidDataException("The OLE property set stream uses an unsupported byte order.");
            }

            uint sectionCount = ReadUInt32(bytes, 24);
            if (sectionCount == 0 || sectionCount > 8) {
                throw new InvalidDataException("The OLE property set stream declares an invalid section count.");
            }

            int sectionListOffset = 28;
            if (sectionListOffset + checked((int)sectionCount) * 20 > bytes.Length) {
                throw new InvalidDataException("The OLE property set stream section table is truncated.");
            }

            var sections = new List<OfficeOlePropertySection>(checked((int)sectionCount));
            for (int i = 0; i < sectionCount; i++) {
                int entryOffset = sectionListOffset + i * 20;
                var formatIdBytes = new byte[16];
                Buffer.BlockCopy(bytes, entryOffset, formatIdBytes, 0,
                    formatIdBytes.Length);
                var formatId = new Guid(formatIdBytes);
                uint sectionOffset = ReadUInt32(bytes, entryOffset + 16);
                sections.Add(ReadSection(bytes, checked((int)sectionOffset),
                    formatId));
            }

            return sections;
        }

        private static OfficeOlePropertySection ReadSection(byte[] bytes,
            int sectionOffset, Guid formatId) {
            if (sectionOffset < 0 || sectionOffset + 8 > bytes.Length) {
                throw new InvalidDataException("The OLE property section is outside the stream.");
            }

            uint sectionSize = ReadUInt32(bytes, sectionOffset);
            uint propertyCount = ReadUInt32(bytes, sectionOffset + 4);
            if (sectionSize < 8 || sectionOffset + sectionSize > bytes.Length || propertyCount > 1024) {
                throw new InvalidDataException("The OLE property section header is invalid.");
            }

            int propertyListOffset = sectionOffset + 8;
            if (propertyListOffset + checked((int)propertyCount) * 8 > sectionOffset + sectionSize) {
                throw new InvalidDataException("The OLE property section table is truncated.");
            }

            int codePage = 1252;
            var offsets = new Dictionary<uint, int>();
            for (int i = 0; i < propertyCount; i++) {
                uint propertyId = ReadUInt32(bytes, propertyListOffset + i * 8);
                uint propertyOffset = ReadUInt32(bytes, propertyListOffset + i * 8 + 4);
                offsets[propertyId] = checked(sectionOffset + (int)propertyOffset);
            }

            if (offsets.TryGetValue(CodePagePropertyId, out int codePageOffset)) {
                OfficeOlePropertyValue codePageValue = ReadPropertyValue(bytes, codePageOffset, codePage);
                if (codePageValue.Value is short shortCodePage) {
                    codePage = shortCodePage;
                } else if (codePageValue.Value is int intCodePage) {
                    codePage = intCodePage;
                }
            }

            var dictionary = new Dictionary<uint, string>();
            if (offsets.TryGetValue(PropertyDictionaryId, out int dictionaryOffset)) {
                dictionary = ReadDictionary(bytes, dictionaryOffset, codePage);
            }

            var properties = new Dictionary<uint, OfficeOlePropertyValue>();
            foreach (KeyValuePair<uint, int> offset in offsets) {
                if (offset.Key == PropertyDictionaryId) {
                    continue;
                }

                properties[offset.Key] = ReadPropertyValue(bytes, offset.Value, codePage);
            }

            return new OfficeOlePropertySection(formatId, properties,
                dictionary);
        }

        private static Dictionary<uint, string> ReadDictionary(byte[] bytes, int offset, int codePage) {
            uint count = ReadUInt32(bytes, offset);
            if (count > 1024) {
                throw new InvalidDataException("The OLE property dictionary declares too many entries.");
            }

            int cursor = offset + 4;
            var result = new Dictionary<uint, string>();
            for (int i = 0; i < count; i++) {
                uint propertyId = ReadUInt32(bytes, cursor);
                uint length = ReadUInt32(bytes, cursor + 4);
                cursor += 8;

                string name;
                if (codePage == 1200) {
                    int byteLength = checked((int)length * 2);
                    if (cursor + byteLength > bytes.Length) {
                        throw new InvalidDataException("The OLE property dictionary string is truncated.");
                    }

                    name = Encoding.Unicode.GetString(bytes, cursor, byteLength).TrimEnd('\0');
                    cursor += byteLength;
                } else {
                    int byteLength = checked((int)length);
                    name = DecodeAnsiString(bytes, cursor, byteLength, codePage).TrimEnd('\0');
                    cursor += byteLength;
                }

                result[propertyId] = name;
                cursor = AlignToInt32(cursor);
            }

            return result;
        }

        private static OfficeOlePropertyValue ReadPropertyValue(byte[] bytes, int offset, int codePage) {
            if (offset < 0 || offset + 4 > bytes.Length) {
                throw new InvalidDataException("The OLE property value is outside the stream.");
            }

            ushort type = ReadUInt16(bytes, offset);
            int valueOffset = offset + 4;
            switch (type) {
                case 0x0006:
                    return new OfficeOlePropertyValue(type, ReadCurrency(bytes, valueOffset));
                case 0x0002:
                    return new OfficeOlePropertyValue(type, unchecked((short)ReadUInt16(bytes, valueOffset)));
                case 0x0003:
                    return new OfficeOlePropertyValue(type, unchecked((int)ReadUInt32(bytes, valueOffset)));
                case 0x0004:
                    EnsureAvailable(bytes, valueOffset, 4);
                    return new OfficeOlePropertyValue(type, BitConverter.ToSingle(bytes, valueOffset));
                case 0x0005:
                    EnsureAvailable(bytes, valueOffset, 8);
                    return new OfficeOlePropertyValue(type, BitConverter.ToDouble(bytes, valueOffset));
                case 0x0007:
                    EnsureAvailable(bytes, valueOffset, 8);
                    return new OfficeOlePropertyValue(type, DateTime.FromOADate(BitConverter.ToDouble(bytes, valueOffset)).ToUniversalTime());
                case 0x000b:
                    return new OfficeOlePropertyValue(type, ReadInt16(bytes, valueOffset) != 0);
                case 0x0010:
                    return new OfficeOlePropertyValue(type, unchecked((sbyte)ReadByte(bytes, valueOffset)));
                case 0x0011:
                    return new OfficeOlePropertyValue(type, ReadByte(bytes, valueOffset));
                case 0x0012:
                    return new OfficeOlePropertyValue(type, ReadUInt16(bytes, valueOffset));
                case 0x0013:
                    return new OfficeOlePropertyValue(type, ReadUInt32(bytes, valueOffset));
                case 0x0014:
                    return new OfficeOlePropertyValue(type, unchecked((long)ReadUInt64(bytes, valueOffset)));
                case 0x0015:
                    return new OfficeOlePropertyValue(type, ReadUInt64(bytes, valueOffset));
                case 0x0016:
                    return new OfficeOlePropertyValue(type, unchecked((int)ReadUInt32(bytes, valueOffset)));
                case 0x0017:
                    return new OfficeOlePropertyValue(type, ReadUInt32(bytes, valueOffset));
                case 0x001e:
                    return new OfficeOlePropertyValue(type, ReadLengthPrefixedAnsiString(bytes, valueOffset, codePage));
                case 0x001f:
                    return new OfficeOlePropertyValue(type, ReadLengthPrefixedUnicodeString(bytes, valueOffset));
                case 0x0040:
                    return new OfficeOlePropertyValue(type, DateTime.FromFileTimeUtc(unchecked((long)ReadUInt64(bytes, valueOffset))));
                case 0x0041:
                    return new OfficeOlePropertyValue(type, ReadBlob(bytes, valueOffset));
                default:
                    return new OfficeOlePropertyValue(type, null);
            }
        }

        private static double ReadCurrency(byte[] bytes, int offset) {
            long scaledValue = unchecked((long)ReadUInt64(bytes, offset));
            return scaledValue / 10000D;
        }

        private static byte[] ReadBlob(byte[] bytes, int offset) {
            uint length = ReadUInt32(bytes, offset);
            if (length == 0) {
                return Array.Empty<byte>();
            }

            int byteCount = checked((int)length);
            if (offset + 4 + byteCount > bytes.Length) {
                throw new InvalidDataException("The OLE binary blob value is truncated.");
            }

            var value = new byte[byteCount];
            Buffer.BlockCopy(bytes, offset + 4, value, 0, byteCount);
            return value;
        }

        private static string ReadLengthPrefixedAnsiString(byte[] bytes, int offset, int codePage) {
            uint length = ReadUInt32(bytes, offset);
            if (length == 0) {
                return string.Empty;
            }

            return DecodeAnsiString(bytes, offset + 4, checked((int)length), codePage).TrimEnd('\0');
        }

        private static string ReadLengthPrefixedUnicodeString(byte[] bytes, int offset) {
            uint charCount = ReadUInt32(bytes, offset);
            if (charCount == 0) {
                return string.Empty;
            }

            int byteCount = checked((int)charCount * 2);
            if (offset + 4 + byteCount > bytes.Length) {
                throw new InvalidDataException("The OLE Unicode string value is truncated.");
            }

            return Encoding.Unicode.GetString(bytes, offset + 4, byteCount).TrimEnd('\0');
        }

        private static string DecodeAnsiString(byte[] bytes, int offset, int length, int codePage) {
            if (offset < 0 || length < 0 || offset + length > bytes.Length) {
                throw new InvalidDataException("The OLE ANSI string value is truncated.");
            }

            try {
                return Encoding.GetEncoding(codePage).GetString(bytes, offset, length);
            } catch (ArgumentException) {
                return DecodeWindows1252Fallback(bytes, offset, length);
            } catch (NotSupportedException) {
                return DecodeWindows1252Fallback(bytes, offset, length);
            }
        }

        private static string DecodeWindows1252Fallback(byte[] bytes, int offset, int length) {
            var chars = new char[length];
            for (int i = 0; i < length; i++) {
                byte value = bytes[offset + i];
                chars[i] = value switch {
                    0x80 => '\u20ac',
                    0x82 => '\u201a',
                    0x83 => '\u0192',
                    0x84 => '\u201e',
                    0x85 => '\u2026',
                    0x86 => '\u2020',
                    0x87 => '\u2021',
                    0x88 => '\u02c6',
                    0x89 => '\u2030',
                    0x8a => '\u0160',
                    0x8b => '\u2039',
                    0x8c => '\u0152',
                    0x8e => '\u017d',
                    0x91 => '\u2018',
                    0x92 => '\u2019',
                    0x93 => '\u201c',
                    0x94 => '\u201d',
                    0x95 => '\u2022',
                    0x96 => '\u2013',
                    0x97 => '\u2014',
                    0x98 => '\u02dc',
                    0x99 => '\u2122',
                    0x9a => '\u0161',
                    0x9b => '\u203a',
                    0x9c => '\u0153',
                    0x9e => '\u017e',
                    0x9f => '\u0178',
                    _ => (char)value
                };
            }

            return new string(chars);
        }

        private static int AlignToInt32(int value) {
            int remainder = value % 4;
            return remainder == 0 ? value : checked(value + (4 - remainder));
        }

        private static short ReadInt16(byte[] bytes, int offset) {
            return unchecked((short)ReadUInt16(bytes, offset));
        }

        private static byte ReadByte(byte[] bytes, int offset) {
            if (offset < 0 || offset >= bytes.Length) throw new InvalidDataException("Unexpected end of OLE property stream.");
            return bytes[offset];
        }

        private static ushort ReadUInt16(byte[] bytes, int offset) {
            EnsureAvailable(bytes, offset, 2);
            return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
        }

        private static uint ReadUInt32(byte[] bytes, int offset) {
            EnsureAvailable(bytes, offset, 4);
            return (uint)(bytes[offset]
                | (bytes[offset + 1] << 8)
                | (bytes[offset + 2] << 16)
                | (bytes[offset + 3] << 24));
        }

        private static ulong ReadUInt64(byte[] bytes, int offset) {
            uint low = ReadUInt32(bytes, offset);
            uint high = ReadUInt32(bytes, offset + 4);
            return low | ((ulong)high << 32);
        }

        private static void EnsureAvailable(byte[] bytes, int offset, int count) {
            if (offset < 0 || count < 0 || offset + count > bytes.Length) {
                throw new InvalidDataException("Unexpected end of OLE property stream.");
            }
        }
    }
}
