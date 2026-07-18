using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

internal static class OabV4RecordReader {
    private static readonly Encoding StrictUtf8 = new UTF8Encoding(false, true);

    static OabV4RecordReader() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    internal static OabRecordEnvelope ReadEnvelope(OabSource source, Stream stream,
        OfflineAddressBookReaderOptions options, string location) {
        int size = ReadRecordSize(source, stream, options, location);
        return new OabRecordEnvelope(size, OabBinary.ReadExactly(stream, size - 4, location));
    }

    internal static int ReadRecordSize(OabSource source, Stream stream,
        OfflineAddressBookReaderOptions options, string location) {
        long relative = stream.Position - source.BaseOffset;
        if (relative < 0 || source.Length - relative < 4) {
            throw new InvalidDataException(string.Concat("OAB record header is truncated at ", location, "."));
        }
        uint encodedSize = OabBinary.ReadUInt32(stream, location);
        if (encodedSize > int.MaxValue) {
            throw new OfflineAddressBookLimitExceededException(
                nameof(options.MaxRecordBytes), encodedSize, options.MaxRecordBytes, location);
        }
        int size = checked((int)encodedSize);
        if (size < 5) throw new InvalidDataException(string.Concat("OAB record is smaller than its framing at ", location, "."));
        if (size > options.MaxRecordBytes) {
            throw new OfflineAddressBookLimitExceededException(
                nameof(options.MaxRecordBytes), size, options.MaxRecordBytes, location);
        }
        if (size > source.Length - relative) {
            throw new InvalidDataException(string.Concat("OAB record extends beyond the source at ", location, "."));
        }
        return size;
    }

    internal static OabParsedRecord Parse(OabRecordEnvelope envelope,
        IReadOnlyList<OfflineAddressBookPropertyDefinition> definitions,
        OfflineAddressBookReaderOptions options,
        string location) {
        if (envelope == null) throw new ArgumentNullException(nameof(envelope));
        if (definitions == null) throw new ArgumentNullException(nameof(definitions));
        int presenceBytes = checked((definitions.Count + 7) / 8);
        if (envelope.Body.Length < presenceBytes) {
            throw new InvalidDataException(string.Concat("OAB record presence array is truncated at ", location, "."));
        }

        var diagnostics = new List<EmailDiagnostic>();
        if (definitions.Count % 8 != 0 && presenceBytes > 0) {
            int unused = 8 - (definitions.Count % 8);
            int mask = (1 << unused) - 1;
            if ((envelope.Body[presenceBytes - 1] & mask) != 0) {
                diagnostics.Add(new EmailDiagnostic(
                    "OAB_RECORD_UNUSED_PRESENCE_BITS",
                    "Unused bits in the final OAB presence byte are not zero.",
                    EmailDiagnosticSeverity.Warning,
                    location));
            }
        }

        Encoding string8;
        try {
            string8 = Encoding.GetEncoding(options.String8CodePage);
        } catch (ArgumentException exception) {
            throw new NotSupportedException(string.Concat(
                "The configured OAB String8 code page is unavailable: ",
                options.String8CodePage.ToString(CultureInfo.InvariantCulture), "."), exception);
        }

        int cursor = presenceBytes;
        var properties = new List<MapiProperty>(definitions.Count);
        for (int index = 0; index < definitions.Count; index++) {
            OfflineAddressBookPropertyDefinition definition = definitions[index];
            int bit = 0x80 >> (index % 8);
            if ((envelope.Body[index / 8] & bit) == 0) {
                if (definition.IsPrimaryKey) {
                    throw new InvalidDataException(string.Concat(
                        "Required OAB primary-key property 0x",
                        definition.PropertyTag.ToString("X8", CultureInfo.InvariantCulture),
                        " is absent at ", location, "."));
                }
                continue;
            }
            if (definition.PropertyType == MapiPropertyType.Object) {
                throw new InvalidDataException(string.Concat(
                    "OAB PtypObject property is marked present at ", location, "."));
            }
            int valueStart = cursor;
            object value = ReadValue(envelope.Body, ref cursor, definition.PropertyType,
                string8, options, location);
            byte[]? raw = null;
            if (options.RetainRawPropertyBytes) {
                raw = new byte[cursor - valueStart];
                Buffer.BlockCopy(envelope.Body, valueStart, raw, 0, raw.Length);
            }
            properties.Add(new MapiProperty(definition.PropertyId, definition.PropertyType, value, flags: 0) {
                RawData = raw
            });
        }

        if (cursor != envelope.Body.Length) {
            diagnostics.Add(new EmailDiagnostic(
                "OAB_RECORD_TRAILING_DATA",
                string.Concat((envelope.Body.Length - cursor).ToString(CultureInfo.InvariantCulture),
                    " unconsumed byte(s) remain in the OAB record."),
                EmailDiagnosticSeverity.Warning,
                location));
        }
        return new OabParsedRecord(properties, diagnostics);
    }

    private static object ReadValue(byte[] data, ref int cursor, MapiPropertyType type,
        Encoding string8, OfflineAddressBookReaderOptions options, string location) {
        switch (type) {
            case MapiPropertyType.Integer32:
                return ReadCompactUInt32(data, ref cursor, location);
            case MapiPropertyType.Boolean:
                EnsureAvailable(data, cursor, 1, location);
                byte boolean = data[cursor++];
                if (boolean > 1) throw new InvalidDataException(string.Concat("Invalid OAB Boolean value at ", location, "."));
                return boolean == 1;
            case MapiPropertyType.String8:
                return ReadNonEmptyString(data, ref cursor, string8, options.MaxStringBytes, location);
            case MapiPropertyType.Unicode:
                return ReadNonEmptyString(data, ref cursor, StrictUtf8, options.MaxStringBytes, location);
            case MapiPropertyType.Binary:
                return ReadBinary(data, ref cursor, options, location);
            case MapiPropertyType.MultipleInteger32:
                return ReadUInt32Array(data, ref cursor, options, location);
            case MapiPropertyType.MultipleString8:
                return ReadStringArray(data, ref cursor, string8, options, location);
            case MapiPropertyType.MultipleUnicode:
                return ReadStringArray(data, ref cursor, StrictUtf8, options, location);
            case MapiPropertyType.MultipleBinary:
                return ReadBinaryArray(data, ref cursor, options, location);
            default:
                throw new NotSupportedException(string.Concat(
                    "Unsupported OAB property type 0x", ((ushort)type).ToString("X4", CultureInfo.InvariantCulture),
                    " at ", location, "."));
        }
    }

    private static uint ReadCompactUInt32(byte[] data, ref int cursor, string location) {
        EnsureAvailable(data, cursor, 1, location);
        byte first = data[cursor++];
        if (first <= 0x7F) return first;
        if (first < 0x81 || first > 0x84) {
            throw new InvalidDataException(string.Concat("Invalid compact OAB integer prefix at ", location, "."));
        }
        int count = first & 0x7F;
        EnsureAvailable(data, cursor, count, location);
        uint value = 0;
        for (int index = 0; index < count; index++) value |= (uint)data[cursor++] << (index * 8);
        uint minimum = count == 1 ? 0x80U : 1U << ((count - 1) * 8);
        if (value < minimum) {
            throw new InvalidDataException(string.Concat("Non-canonical compact OAB integer at ", location, "."));
        }
        return value;
    }

    private static string ReadString(byte[] data, ref int cursor, Encoding encoding,
        int maxBytes, string location) {
        int start = cursor;
        int limit = Math.Min(data.Length, checked(start + maxBytes + 1));
        while (cursor < limit && data[cursor] != 0) cursor++;
        if (cursor >= data.Length || data[cursor] != 0) {
            if (cursor >= start + maxBytes) {
                throw new OfflineAddressBookLimitExceededException(
                    nameof(OfflineAddressBookReaderOptions.MaxStringBytes), cursor - start, maxBytes, location);
            }
            throw new InvalidDataException(string.Concat("Unterminated OAB string at ", location, "."));
        }
        int count = cursor - start;
        cursor++;
        try {
            return encoding.GetString(data, start, count);
        } catch (DecoderFallbackException exception) {
            throw new InvalidDataException(string.Concat("Invalid OAB string encoding at ", location, "."), exception);
        }
    }

    private static string ReadNonEmptyString(byte[] data, ref int cursor, Encoding encoding,
        int maxBytes, string location) {
        string value = ReadString(data, ref cursor, encoding, maxBytes, location);
        if (value.Length == 0) {
            throw new InvalidDataException(string.Concat(
                "Empty OAB string value is marked present at ", location, "."));
        }
        return value;
    }

    private static byte[] ReadBinary(byte[] data, ref int cursor,
        OfflineAddressBookReaderOptions options, string location) {
        uint encodedLength = ReadCompactUInt32(data, ref cursor, location);
        if (encodedLength > int.MaxValue || encodedLength > options.MaxBinaryBytes) {
            throw new OfflineAddressBookLimitExceededException(
                nameof(options.MaxBinaryBytes), encodedLength, options.MaxBinaryBytes, location);
        }
        int length = checked((int)encodedLength);
        if (length == 0) throw new InvalidDataException(string.Concat("Empty OAB binary value at ", location, "."));
        EnsureAvailable(data, cursor, length, location);
        var value = new byte[length];
        Buffer.BlockCopy(data, cursor, value, 0, length);
        cursor += length;
        return value;
    }

    private static uint[] ReadUInt32Array(byte[] data, ref int cursor,
        OfflineAddressBookReaderOptions options, string location) {
        int count = ReadCount(data, ref cursor, options, location);
        var values = new uint[count];
        for (int index = 0; index < count; index++) values[index] = ReadCompactUInt32(data, ref cursor, location);
        return values;
    }

    private static string[] ReadStringArray(byte[] data, ref int cursor, Encoding encoding,
        OfflineAddressBookReaderOptions options, string location) {
        int count = ReadCount(data, ref cursor, options, location);
        var values = new string[count];
        for (int index = 0; index < count; index++) {
            values[index] = ReadString(data, ref cursor, encoding, options.MaxStringBytes, location);
            if (values[index].Length == 0) throw new InvalidDataException(string.Concat("Empty OAB string-array value at ", location, "."));
        }
        return values;
    }

    private static byte[][] ReadBinaryArray(byte[] data, ref int cursor,
        OfflineAddressBookReaderOptions options, string location) {
        int count = ReadCount(data, ref cursor, options, location);
        var values = new byte[count][];
        for (int index = 0; index < count; index++) values[index] = ReadBinary(data, ref cursor, options, location);
        return values;
    }

    private static int ReadCount(byte[] data, ref int cursor,
        OfflineAddressBookReaderOptions options, string location) {
        uint encoded = ReadCompactUInt32(data, ref cursor, location);
        if (encoded > int.MaxValue || encoded > options.MaxValuesPerProperty) {
            throw new OfflineAddressBookLimitExceededException(
                nameof(options.MaxValuesPerProperty), encoded, options.MaxValuesPerProperty, location);
        }
        return checked((int)encoded);
    }

    private static void EnsureAvailable(byte[] data, int offset, int count, string location) {
        if (offset < 0 || count < 0 || offset > data.Length - count) {
            throw new InvalidDataException(string.Concat("OAB property value is truncated at ", location, "."));
        }
    }
}
