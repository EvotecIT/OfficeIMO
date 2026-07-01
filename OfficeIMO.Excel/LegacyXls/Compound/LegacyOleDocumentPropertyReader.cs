using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Compound {
    internal static class LegacyOleDocumentPropertyReader {
        private const string SummaryInformationStreamName = "\u0005SummaryInformation";
        private const string DocumentSummaryInformationStreamName = "\u0005DocumentSummaryInformation";
        private const uint PropertyDictionaryId = 0;
        private const uint CodePagePropertyId = 1;
        private const string UnsupportedCustomPropertyCode = "XLS-OLE-CUSTOM-DOCUMENT-PROPERTY-UNSUPPORTED";

        internal static void AddDocumentProperties(LegacyCompoundFile compoundFile, LegacyXlsWorkbook workbook, LegacyXlsImportOptions options) {
            if (compoundFile == null) throw new ArgumentNullException(nameof(compoundFile));
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var properties = new LegacyXlsDocumentProperties();
            bool parsedAny = false;

            if (TryGetStream(compoundFile.Streams, SummaryInformationStreamName, out byte[]? summaryStream)) {
                parsedAny |= TryReadSummaryInformation(summaryStream!, properties, workbook);
            }

            if (TryGetStream(compoundFile.Streams, DocumentSummaryInformationStreamName, out byte[]? documentSummaryStream)) {
                parsedAny |= TryReadDocumentSummaryInformation(documentSummaryStream!, properties, workbook, options);
            }

            if (parsedAny && properties.HasAnyProperties) {
                workbook.SetDocumentProperties(properties);
            }
        }

        private static bool TryGetStream(IReadOnlyDictionary<string, byte[]> streams, string name, out byte[]? bytes) {
            return streams.TryGetValue(name, out bytes) && bytes != null && bytes.Length > 0;
        }

        private static bool TryReadSummaryInformation(byte[] bytes, LegacyXlsDocumentProperties target, LegacyXlsWorkbook workbook) {
            try {
                IReadOnlyList<OlePropertySection> sections = ReadPropertySetSections(bytes);
                foreach (OlePropertySection section in sections) {
                    foreach (KeyValuePair<uint, OlePropertyValue> property in section.Properties) {
                        switch (property.Key) {
                            case 2:
                                target.Title = property.Value.AsString();
                                break;
                            case 3:
                                target.Subject = property.Value.AsString();
                                break;
                            case 4:
                                target.Creator = property.Value.AsString();
                                break;
                            case 5:
                                target.Keywords = property.Value.AsString();
                                break;
                            case 6:
                                target.Description = property.Value.AsString();
                                break;
                            case 8:
                                target.LastModifiedBy = property.Value.AsString();
                                break;
                            case 9:
                                target.Revision = property.Value.AsString();
                                break;
                            case 11:
                                target.LastPrinted = property.Value.AsDateTime();
                                break;
                            case 12:
                                target.Created = property.Value.AsDateTime();
                                break;
                            case 13:
                                target.Modified = property.Value.AsDateTime();
                                break;
                        }
                    }
                }

                return true;
            } catch (Exception ex) when (ex is IOException || ex is ArgumentException || ex is InvalidDataException || ex is OverflowException) {
                AddPropertyWarning(workbook, SummaryInformationStreamName, ex);
                return false;
            }
        }

        private static bool TryReadDocumentSummaryInformation(byte[] bytes, LegacyXlsDocumentProperties target, LegacyXlsWorkbook workbook, LegacyXlsImportOptions options) {
            try {
                IReadOnlyList<OlePropertySection> sections = ReadPropertySetSections(bytes);
                bool parsed = false;
                foreach (OlePropertySection section in sections) {
                    if (section.Dictionary.Count == 0) {
                        if (section.Properties.TryGetValue(2, out OlePropertyValue? category)) {
                            target.Category = category.AsString();
                            parsed = true;
                        }

                        if (section.Properties.TryGetValue(14, out OlePropertyValue? manager)) {
                            target.Manager = manager.AsString();
                            parsed = true;
                        }

                        if (section.Properties.TryGetValue(15, out OlePropertyValue? company)) {
                            target.Company = company.AsString();
                            parsed = true;
                        }
                    }

                    foreach (KeyValuePair<uint, string> name in section.Dictionary) {
                        if (name.Key == PropertyDictionaryId || name.Key == CodePagePropertyId) {
                            continue;
                        }

                        if (!section.Properties.TryGetValue(name.Key, out OlePropertyValue? value)) {
                            continue;
                        }

                        if (TryCreateCustomPropertyValue(value, out LegacyXlsDocumentPropertyValue? customValue)) {
                            target.SetCustomProperty(name.Value, customValue!);
                            parsed = true;
                        } else {
                            AddUnsupportedCustomProperty(workbook, options, name.Key, name.Value, value.Type);
                        }
                    }
                }

                return parsed;
            } catch (Exception ex) when (ex is IOException || ex is ArgumentException || ex is InvalidDataException || ex is OverflowException) {
                AddPropertyWarning(workbook, DocumentSummaryInformationStreamName, ex);
                return false;
            }
        }

        private static void AddUnsupportedCustomProperty(
            LegacyXlsWorkbook workbook,
            LegacyXlsImportOptions options,
            uint propertyId,
            string propertyName,
            ushort propertyType) {
            string detailCode = $"DocumentProperty:Custom:PropertyId:0x{propertyId:X4}:Type:0x{propertyType:X4}";
            string description = $"The OLE custom document property '{propertyName}' uses unsupported VARTYPE 0x{propertyType:X4}; the property was not projected.";
            var feature = new LegacyXlsUnsupportedFeature(
                LegacyXlsUnsupportedFeatureKind.DocumentProperty,
                UnsupportedCustomPropertyCode,
                description,
                detailCode: detailCode);
            workbook.MutableUnsupportedFeatures.Add(feature);
            if (options.ReportUnsupportedRecords) {
                workbook.MutableDiagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Info,
                    feature.Code,
                    feature.Description,
                    detailCode: feature.DetailCode));
            }
        }

        private static bool TryCreateCustomPropertyValue(OlePropertyValue value, out LegacyXlsDocumentPropertyValue? customValue) {
            customValue = null;
            object? rawValue = value.Value;
            if (rawValue == null) {
                return false;
            }

            switch (rawValue) {
                case sbyte signedByte:
                    customValue = new LegacyXlsDocumentPropertyValue((int)signedByte, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case byte unsignedByte:
                    customValue = new LegacyXlsDocumentPropertyValue(unsignedByte, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case short signedShort:
                    customValue = new LegacyXlsDocumentPropertyValue((int)signedShort, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case ushort unsignedShort:
                    customValue = new LegacyXlsDocumentPropertyValue(unsignedShort, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case string text:
                    customValue = new LegacyXlsDocumentPropertyValue(text, LegacyXlsDocumentPropertyValueKind.Text);
                    return true;
                case bool boolean:
                    customValue = new LegacyXlsDocumentPropertyValue(boolean, LegacyXlsDocumentPropertyValueKind.Boolean);
                    return true;
                case DateTime dateTime:
                    customValue = new LegacyXlsDocumentPropertyValue(dateTime, LegacyXlsDocumentPropertyValueKind.DateTime);
                    return true;
                case int integer:
                    customValue = new LegacyXlsDocumentPropertyValue(integer, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case long integer64 when integer64 >= int.MinValue && integer64 <= int.MaxValue:
                    customValue = new LegacyXlsDocumentPropertyValue((int)integer64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case long integer64:
                    customValue = new LegacyXlsDocumentPropertyValue(integer64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case uint unsignedInteger:
                    customValue = new LegacyXlsDocumentPropertyValue(unsignedInteger, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case ulong unsignedInteger64 when unsignedInteger64 <= int.MaxValue:
                    customValue = new LegacyXlsDocumentPropertyValue((int)unsignedInteger64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case ulong unsignedInteger64 when unsignedInteger64 <= long.MaxValue:
                    customValue = new LegacyXlsDocumentPropertyValue((long)unsignedInteger64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case ulong unsignedInteger64:
                    customValue = new LegacyXlsDocumentPropertyValue(unsignedInteger64, LegacyXlsDocumentPropertyValueKind.Integer);
                    return true;
                case double number:
                    customValue = new LegacyXlsDocumentPropertyValue(number, LegacyXlsDocumentPropertyValueKind.Number);
                    return true;
                case float number:
                    customValue = new LegacyXlsDocumentPropertyValue((double)number, LegacyXlsDocumentPropertyValueKind.Number);
                    return true;
                case byte[] bytes:
                    customValue = new LegacyXlsDocumentPropertyValue((byte[])bytes.Clone(), LegacyXlsDocumentPropertyValueKind.Binary);
                    return true;
            }

            return false;
        }

        private static IReadOnlyList<OlePropertySection> ReadPropertySetSections(byte[] bytes) {
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

            var sections = new List<OlePropertySection>(checked((int)sectionCount));
            for (int i = 0; i < sectionCount; i++) {
                int entryOffset = sectionListOffset + i * 20;
                uint sectionOffset = ReadUInt32(bytes, entryOffset + 16);
                sections.Add(ReadSection(bytes, checked((int)sectionOffset)));
            }

            return sections;
        }

        private static OlePropertySection ReadSection(byte[] bytes, int sectionOffset) {
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
                OlePropertyValue codePageValue = ReadPropertyValue(bytes, codePageOffset, codePage);
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

            var properties = new Dictionary<uint, OlePropertyValue>();
            foreach (KeyValuePair<uint, int> offset in offsets) {
                if (offset.Key == PropertyDictionaryId) {
                    continue;
                }

                properties[offset.Key] = ReadPropertyValue(bytes, offset.Value, codePage);
            }

            return new OlePropertySection(properties, dictionary);
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

        private static OlePropertyValue ReadPropertyValue(byte[] bytes, int offset, int codePage) {
            if (offset < 0 || offset + 4 > bytes.Length) {
                throw new InvalidDataException("The OLE property value is outside the stream.");
            }

            ushort type = ReadUInt16(bytes, offset);
            int valueOffset = offset + 4;
            switch (type) {
                case 0x0006:
                    return new OlePropertyValue(type, ReadCurrency(bytes, valueOffset));
                case 0x0002:
                    return new OlePropertyValue(type, unchecked((short)ReadUInt16(bytes, valueOffset)));
                case 0x0003:
                    return new OlePropertyValue(type, unchecked((int)ReadUInt32(bytes, valueOffset)));
                case 0x0004:
                    return new OlePropertyValue(type, BitConverter.ToSingle(bytes, valueOffset));
                case 0x0005:
                    return new OlePropertyValue(type, BitConverter.ToDouble(bytes, valueOffset));
                case 0x0007:
                    return new OlePropertyValue(type, DateTime.FromOADate(BitConverter.ToDouble(bytes, valueOffset)).ToUniversalTime());
                case 0x000b:
                    return new OlePropertyValue(type, ReadInt16(bytes, valueOffset) != 0);
                case 0x0010:
                    return new OlePropertyValue(type, unchecked((sbyte)ReadByte(bytes, valueOffset)));
                case 0x0011:
                    return new OlePropertyValue(type, ReadByte(bytes, valueOffset));
                case 0x0012:
                    return new OlePropertyValue(type, ReadUInt16(bytes, valueOffset));
                case 0x0013:
                    return new OlePropertyValue(type, ReadUInt32(bytes, valueOffset));
                case 0x0014:
                    return new OlePropertyValue(type, unchecked((long)ReadUInt64(bytes, valueOffset)));
                case 0x0015:
                    return new OlePropertyValue(type, ReadUInt64(bytes, valueOffset));
                case 0x0016:
                    return new OlePropertyValue(type, unchecked((int)ReadUInt32(bytes, valueOffset)));
                case 0x0017:
                    return new OlePropertyValue(type, ReadUInt32(bytes, valueOffset));
                case 0x001e:
                    return new OlePropertyValue(type, ReadLengthPrefixedAnsiString(bytes, valueOffset, codePage));
                case 0x001f:
                    return new OlePropertyValue(type, ReadLengthPrefixedUnicodeString(bytes, valueOffset));
                case 0x0040:
                    return new OlePropertyValue(type, DateTime.FromFileTimeUtc(unchecked((long)ReadUInt64(bytes, valueOffset))));
                case 0x0041:
                    return new OlePropertyValue(type, ReadBlob(bytes, valueOffset));
                default:
                    return new OlePropertyValue(type, null);
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

        private static void AddPropertyWarning(LegacyXlsWorkbook workbook, string streamName, Exception exception) {
            workbook.MutableDiagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Warning,
                "XLS-OLE-PROPERTIES-UNREADABLE",
                string.Format(CultureInfo.InvariantCulture, "The OLE document property stream '{0}' could not be read. {1}", streamName, exception.Message)));
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
            if (offset < 0 || offset + 2 > bytes.Length) throw new InvalidDataException("Unexpected end of OLE property stream.");
            return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
        }

        private static uint ReadUInt32(byte[] bytes, int offset) {
            if (offset < 0 || offset + 4 > bytes.Length) throw new InvalidDataException("Unexpected end of OLE property stream.");
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

        private sealed class OlePropertySection {
            internal OlePropertySection(
                IReadOnlyDictionary<uint, OlePropertyValue> properties,
                IReadOnlyDictionary<uint, string> dictionary) {
                Properties = properties;
                Dictionary = dictionary;
            }

            internal IReadOnlyDictionary<uint, OlePropertyValue> Properties { get; }

            internal IReadOnlyDictionary<uint, string> Dictionary { get; }
        }

        private sealed class OlePropertyValue {
            internal OlePropertyValue(ushort type, object? value) {
                Type = type;
                Value = value;
            }

            internal ushort Type { get; }

            internal object? Value { get; }

            internal string? AsString() {
                return Value as string;
            }

            internal DateTime? AsDateTime() {
                return Value as DateTime?;
            }
        }
    }
}
