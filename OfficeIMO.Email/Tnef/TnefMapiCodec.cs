namespace OfficeIMO.Email;

internal static class TnefMapiCodec {
    internal static List<MapiProperty> ReadProperties(byte[] data, int codePage, MsgParserState state, string location) {
        var cursor = new Cursor(data);
        return ReadPropertyArray(cursor, codePage, state, location);
    }

    internal static List<List<MapiProperty>> ReadRecipientTable(byte[] data, int codePage, MsgParserState state, string location) {
        var cursor = new Cursor(data);
        if (cursor.Remaining < 4) {
            ReportTruncated(state, location);
            return new List<List<MapiProperty>>();
        }
        uint rawCount = cursor.ReadUInt32();
        if (rawCount > state.Options.MaxPartCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxPartCount), rawCount, state.Options.MaxPartCount);
        }
        var rows = new List<List<MapiProperty>>((int)rawCount);
        for (int index = 0; index < rawCount; index++) {
            state.ThrowIfCancellationRequested();
            if (cursor.Remaining < 4) {
                ReportTruncated(state, string.Concat(location, "/recipient[",
                    index.ToString(CultureInfo.InvariantCulture), "]"));
                break;
            }
            rows.Add(ReadPropertyArray(cursor, codePage, state,
                string.Concat(location, "/recipient[", index.ToString(CultureInfo.InvariantCulture), "]")));
        }
        return rows;
    }

    internal static byte[] WriteProperties(IEnumerable<MapiProperty> properties, int codePage,
        IList<EmailDiagnostic> diagnostics, string location) {
        MapiProperty[] values = PrepareProperties(properties, diagnostics, location);
        using (MemoryStream output = new MemoryStream()) {
            WriteUInt32(output, unchecked((uint)values.Length));
            foreach (MapiProperty property in values) WriteProperty(output, property, codePage);
            return output.ToArray();
        }
    }

    internal static byte[] WriteRecipientTable(IEnumerable<IReadOnlyList<MapiProperty>> rows, int codePage,
        IList<EmailDiagnostic> diagnostics, string location) {
        MapiProperty[][] values = rows.Select((row, index) => PrepareProperties(row, diagnostics,
            string.Concat(location, "/row[", index.ToString(CultureInfo.InvariantCulture), "]"))).ToArray();
        using (MemoryStream output = new MemoryStream()) {
            WriteUInt32(output, unchecked((uint)values.Length));
            foreach (IReadOnlyList<MapiProperty> row in values) {
                WriteUInt32(output, unchecked((uint)row.Count));
                foreach (MapiProperty property in row) WriteProperty(output, property, codePage);
            }
            return output.ToArray();
        }
    }

    private static MapiProperty[] PrepareProperties(IEnumerable<MapiProperty> properties,
        IList<EmailDiagnostic> diagnostics, string location) {
        var values = new List<MapiProperty>();
        foreach (MapiProperty property in properties) {
            if (property.PropertyId >= 0x8000 && property.Name == null) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_NAMED_PROPERTY_DESCRIPTOR_MISSING",
                    string.Concat("TNEF named property 0x",
                        property.PropertyId.ToString("X4", CultureInfo.InvariantCulture),
                        " has no property-set descriptor and was not written."),
                    EmailDiagnosticSeverity.Error, location));
                continue;
            }
            values.Add(property);
        }
        return values.ToArray();
    }

    private static List<MapiProperty> ReadPropertyArray(Cursor cursor, int codePage, MsgParserState state, string location) {
        if (cursor.Remaining < 4) {
            ReportTruncated(state, location);
            return new List<MapiProperty>();
        }
        uint rawCount = cursor.ReadUInt32();
        if (rawCount > state.Options.MaxMapiPropertyCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxMapiPropertyCount),
                rawCount, state.Options.MaxMapiPropertyCount);
        }
        var properties = new List<MapiProperty>((int)rawCount);
        for (int index = 0; index < rawCount; index++) {
            state.ThrowIfCancellationRequested();
            try {
                uint tag = cursor.ReadUInt32();
                ushort propertyId = unchecked((ushort)(tag >> 16));
                MapiPropertyType type = (MapiPropertyType)unchecked((ushort)tag);
                MapiNamedProperty? name = propertyId >= 0x8000 ? ReadNamedProperty(cursor) : null;
                bool multiple = MsgValueWriter.IsMultiple(type);
                bool variable = IsVariableValue(type) || multiple;
                object? value;
                byte[]? raw;
                if (variable) {
                    uint valueCount = cursor.ReadUInt32();
                    if (valueCount > state.Options.MaxMapiPropertyCount) {
                        throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxMapiPropertyCount),
                            valueCount, state.Options.MaxMapiPropertyCount);
                    }
                    var decoded = new object[valueCount];
                    using (MemoryStream rawValues = new MemoryStream()) {
                        MapiPropertyType itemType = multiple ? MsgValueWriter.GetMultipleItemType(type) : type;
                        for (int valueIndex = 0; valueIndex < valueCount; valueIndex++) {
                            state.ThrowIfCancellationRequested();
                            byte[] itemBytes;
                            if (IsVariableValue(itemType)) {
                                uint rawLength = cursor.ReadUInt32();
                                if (rawLength > int.MaxValue) throw new InvalidDataException("TNEF MAPI value is too large.");
                                itemBytes = cursor.ReadBytes((int)rawLength);
                                cursor.Align4();
                            } else {
                                int size = GetFixedSize(itemType);
                                itemBytes = cursor.ReadBytes(size);
                                cursor.Align4();
                            }
                            rawValues.Write(itemBytes, 0, itemBytes.Length);
                            decoded[valueIndex] = DecodeValue(itemType, itemBytes, codePage, state.Diagnostics, location) ?? Array.Empty<byte>();
                        }
                        raw = rawValues.ToArray();
                    }
                    value = multiple ? decoded : decoded.FirstOrDefault();
                } else {
                    int size = GetFixedSize(type);
                    raw = cursor.ReadBytes(size);
                    cursor.Align4();
                    value = DecodeValue(type, raw, codePage, state.Diagnostics, location);
                }
                state.CountProperty(raw?.Length ?? 0);
                properties.Add(new MapiProperty(propertyId, type, value, name: name) { RawData = raw });
            } catch (Exception ex) when (ex is InvalidDataException || ex is ArgumentOutOfRangeException ||
                ex is OverflowException) {
                ReportTruncated(state, string.Concat(location, "/property[",
                    index.ToString(CultureInfo.InvariantCulture), "]"));
                break;
            }
        }
        return properties;
    }

    private static void ReportTruncated(MsgParserState state, string location) {
        state.Diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_MAPI_TRUNCATED",
            "A TNEF MAPI property array ended before its declared rows or values were complete.",
            EmailDiagnosticSeverity.Error, location));
    }

    private static MapiNamedProperty ReadNamedProperty(Cursor cursor) {
        Guid propertySet = new Guid(cursor.ReadBytes(16));
        uint kind = cursor.ReadUInt32();
        if (kind == 0) return new MapiNamedProperty(propertySet, cursor.ReadUInt32());
        if (kind != 1) throw new InvalidDataException("Unknown TNEF named-property kind.");
        uint characters = cursor.ReadUInt32();
        if (characters > int.MaxValue / 2) throw new InvalidDataException("TNEF named-property string is too large.");
        byte[] bytes = cursor.ReadBytes(checked((int)characters * 2));
        cursor.Align4();
        return new MapiNamedProperty(propertySet, Encoding.Unicode.GetString(bytes).TrimEnd('\0'));
    }

    private static object? DecodeValue(MapiPropertyType type, byte[] bytes, int codePage,
        IList<EmailDiagnostic> diagnostics, string location) {
        switch (type) {
            case MapiPropertyType.Null: return null;
            case MapiPropertyType.Integer16: return MsgBinary.ReadInt16(bytes, 0);
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode: return MsgBinary.ReadInt32(bytes, 0);
            case MapiPropertyType.Floating32: return MsgBinary.ReadSingle(bytes, 0);
            case MapiPropertyType.Floating64: return MsgBinary.ReadDouble(bytes, 0);
            case MapiPropertyType.Currency: return MsgBinary.ReadInt64(bytes, 0) / 10000m;
            case MapiPropertyType.FloatingTime: return DateTime.FromOADate(MsgBinary.ReadDouble(bytes, 0));
            case MapiPropertyType.Boolean: return MsgBinary.ReadUInt16(bytes, 0) != 0;
            case MapiPropertyType.Integer64: return MsgBinary.ReadInt64(bytes, 0);
            case MapiPropertyType.String8:
                return MimeTextCodec.DecodeText(bytes, codePage, diagnostics, location).TrimEnd('\0');
            case MapiPropertyType.Unicode:
                return Encoding.Unicode.GetString(bytes, 0, bytes.Length - bytes.Length % 2).TrimEnd('\0');
            case MapiPropertyType.Time:
                return new DateTimeOffset(DateTime.FromFileTimeUtc(MsgBinary.ReadInt64(bytes, 0)), TimeSpan.Zero);
            case MapiPropertyType.Guid: return new Guid(MsgBinary.Slice(bytes, 0, 16));
            case MapiPropertyType.Binary:
            case MapiPropertyType.Object: return (byte[])bytes.Clone();
            default: return (byte[])bytes.Clone();
        }
    }

    private static void WriteProperty(Stream output, MapiProperty property, int codePage) {
        ushort propertyId = property.Name != null && property.PropertyId < 0x8000 ? (ushort)0x8000 : property.PropertyId;
        uint tag = ((uint)propertyId << 16) | (ushort)property.PropertyType;
        WriteUInt32(output, tag);
        if (property.Name != null) WriteNamedProperty(output, property.Name);
        bool multiple = MsgValueWriter.IsMultiple(property.PropertyType);
        bool variable = IsVariableValue(property.PropertyType) || multiple;
        if (variable) {
            object[] values = multiple ? MsgValueWriter.GetMultipleValues(property) : new[] { property.Value ?? property.RawData ?? Array.Empty<byte>() };
            WriteUInt32(output, unchecked((uint)values.Length));
            MapiPropertyType itemType = multiple ? MsgValueWriter.GetMultipleItemType(property.PropertyType) : property.PropertyType;
            foreach (object value in values) {
                var item = new MapiProperty(propertyId, itemType, value) { RawData = property.RawData };
                byte[] bytes = EncodeValue(item, codePage);
                if (IsVariableValue(itemType)) WriteUInt32(output, unchecked((uint)bytes.Length));
                output.Write(bytes, 0, bytes.Length);
                Pad4(output);
            }
        } else {
            byte[] bytes = EncodeValue(property, codePage);
            output.Write(bytes, 0, bytes.Length);
            Pad4(output);
        }
    }

    private static void WriteNamedProperty(Stream output, MapiNamedProperty name) {
        byte[] guid = name.PropertySet.ToByteArray();
        output.Write(guid, 0, guid.Length);
        WriteUInt32(output, name.Name == null ? 0U : 1U);
        if (name.Name == null) {
            WriteUInt32(output, name.LocalId.GetValueOrDefault());
        } else {
            string text = string.Concat(name.Name, "\0");
            WriteUInt32(output, unchecked((uint)text.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(text);
            output.Write(bytes, 0, bytes.Length);
            Pad4(output);
        }
    }

    private static byte[] EncodeValue(MapiProperty property, int codePage) {
        if (property.PropertyType == MapiPropertyType.String8) {
            string text = string.Concat(Convert.ToString(property.Value, CultureInfo.InvariantCulture) ?? string.Empty, "\0");
            return MsgValueWriter.EncodeString8(text, codePage);
        }
        if (property.PropertyType == MapiPropertyType.Unicode) {
            return Encoding.Unicode.GetBytes(string.Concat(Convert.ToString(property.Value, CultureInfo.InvariantCulture) ?? string.Empty, "\0"));
        }
        if ((property.PropertyType == MapiPropertyType.Object || property.PropertyType == MapiPropertyType.Binary) && property.Value is byte[] bytes) {
            return (byte[])bytes.Clone();
        }
        return MsgValueWriter.EncodeScalar(property);
    }

    private static int GetFixedSize(MapiPropertyType type) {
        switch (type) {
            case MapiPropertyType.Integer16:
            case MapiPropertyType.Boolean: return 2;
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
            case MapiPropertyType.Floating32: return 4;
            case MapiPropertyType.Guid: return 16;
            default: return 8;
        }
    }

    private static bool IsVariableValue(MapiPropertyType type) {
        return type != MapiPropertyType.Guid && MsgValueWriter.IsVariable(type);
    }

    private static void WriteUInt32(Stream stream, uint value) {
        byte[] bytes = new byte[4];
        MsgBinary.WriteUInt32(bytes, 0, value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void Pad4(Stream stream) {
        while (stream.Position % 4 != 0) stream.WriteByte(0);
    }

    private sealed class Cursor {
        private readonly byte[] _data;

        internal Cursor(byte[] data) { _data = data; }

        internal int Position { get; private set; }

        internal int Remaining => _data.Length - Position;

        internal uint ReadUInt32() {
            uint value = MsgBinary.ReadUInt32(_data, Position);
            Position += 4;
            return value;
        }

        internal byte[] ReadBytes(int count) {
            byte[] value = MsgBinary.Slice(_data, Position, count);
            Position += count;
            return value;
        }

        internal void Align4() {
            Position = checked((Position + 3) & ~3);
            if (Position > _data.Length) throw new InvalidDataException("Unexpected end of TNEF MAPI data.");
        }
    }
}
