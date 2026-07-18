namespace OfficeIMO.Email;

internal static class TnefMapiCodec {
    internal static bool TryPreflightProperties(byte[] data, int offset, int count, MsgParserState state,
        bool recipientTable, out long decodedPropertyBytes, out long attachmentPayloadLength) {
        decodedPropertyBytes = 0;
        attachmentPayloadLength = 0;
        try {
            var cursor = new PreflightCursor(data, offset, count);
            if (recipientTable) {
                uint rowCount = cursor.ReadUInt32();
                if (rowCount > state.Options.MaxPartCount) {
                    throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxPartCount),
                        rowCount, state.Options.MaxPartCount);
                }
                for (uint rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    PreflightPropertyArray(cursor, state, ref decodedPropertyBytes, ref attachmentPayloadLength);
                }
            } else {
                PreflightPropertyArray(cursor, state, ref decodedPropertyBytes, ref attachmentPayloadLength);
            }
            if (cursor.Remaining != 0) {
                throw new InvalidDataException("TNEF MAPI property data contains trailing bytes.");
            }
            return true;
        } catch (Exception exception) when (exception is InvalidDataException || exception is ArgumentOutOfRangeException ||
            exception is OverflowException || exception is IndexOutOfRangeException) {
            decodedPropertyBytes = 0;
            attachmentPayloadLength = 0;
            return false;
        }
    }

    private static void PreflightPropertyArray(PreflightCursor cursor, MsgParserState state,
        ref long decodedPropertyBytes, ref long attachmentPayloadLength) {
        uint propertyCount = cursor.ReadUInt32();
        if (propertyCount > state.Options.MaxMapiPropertyCount) {
            throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxMapiPropertyCount),
                propertyCount, state.Options.MaxMapiPropertyCount);
        }

        for (uint propertyIndex = 0; propertyIndex < propertyCount; propertyIndex++) {
            uint tag = cursor.ReadUInt32();
            ushort propertyId = unchecked((ushort)(tag >> 16));
            MapiPropertyType type = (MapiPropertyType)unchecked((ushort)tag);
            if (propertyId >= 0x8000) SkipNamedProperty(cursor);

            bool multiple = MsgValueWriter.IsMultiple(type);
            bool variable = IsVariableValue(type) || multiple;
            if (!variable) {
                int fixedSize = GetFixedSize(type);
                if (!IsAttachmentPayload(propertyId)) decodedPropertyBytes = checked(decodedPropertyBytes + fixedSize);
                cursor.Skip(fixedSize);
                cursor.Align4();
                continue;
            }

            uint valueCount = cursor.ReadUInt32();
            if (valueCount > state.Options.MaxMapiPropertyCount) {
                throw new EmailLimitExceededException(nameof(EmailReaderOptions.MaxMapiPropertyCount),
                    valueCount, state.Options.MaxMapiPropertyCount);
            }
            MapiPropertyType itemType = multiple ? MsgValueWriter.GetMultipleItemType(type) : type;
            for (uint valueIndex = 0; valueIndex < valueCount; valueIndex++) {
                long itemLength = IsVariableValue(itemType) ? cursor.ReadUInt32() : GetFixedSize(itemType);
                if (IsAttachmentPayload(propertyId)) {
                    attachmentPayloadLength = checked(attachmentPayloadLength + itemLength);
                }
                if (itemLength > int.MaxValue) throw new InvalidDataException("TNEF MAPI value is too large.");
                if (!IsAttachmentPayload(propertyId)) decodedPropertyBytes = checked(decodedPropertyBytes + itemLength);
                cursor.Skip((int)itemLength);
                cursor.Align4();
            }
        }
    }

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
            foreach (MapiProperty property in values) WriteProperty(output, property, codePage, diagnostics, location);
            return output.ToArray();
        }
    }

    internal static byte[] WriteRecipientTable(IEnumerable<IReadOnlyList<MapiProperty>> rows, int codePage,
        IList<EmailDiagnostic> diagnostics, string location) {
        MapiProperty[][] values = rows.Select((row, index) => PrepareProperties(row, diagnostics,
            string.Concat(location, "/row[", index.ToString(CultureInfo.InvariantCulture), "]"))).ToArray();
        using (MemoryStream output = new MemoryStream()) {
            WriteUInt32(output, unchecked((uint)values.Length));
            for (int rowIndex = 0; rowIndex < values.Length; rowIndex++) {
                IReadOnlyList<MapiProperty> row = values[rowIndex];
                WriteUInt32(output, unchecked((uint)row.Count));
                foreach (MapiProperty property in row) {
                    WriteProperty(output, property, codePage, diagnostics,
                        string.Concat(location, "/row[", rowIndex.ToString(CultureInfo.InvariantCulture), "]"));
                }
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
                                if (!IsAttachmentPayload(propertyId)) {
                                    state.EnsureDecodedPropertyBytesWithinLimits(
                                        checked(rawValues.Length + rawLength));
                                }
                                itemBytes = cursor.ReadBytes((int)rawLength);
                                cursor.Align4();
                            } else {
                                int size = GetFixedSize(itemType);
                                if (!IsAttachmentPayload(propertyId)) {
                                    state.EnsureDecodedPropertyBytesWithinLimits(
                                        checked(rawValues.Length + size));
                                }
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
                    if (!IsAttachmentPayload(propertyId)) state.EnsureDecodedPropertyBytesWithinLimits(size);
                    raw = cursor.ReadBytes(size);
                    cursor.Align4();
                    value = DecodeValue(type, raw, codePage, state.Diagnostics, location);
                }
                state.CountProperty(IsAttachmentPayload(propertyId) ? 0 : raw?.Length ?? 0);
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

    private static bool IsAttachmentPayload(ushort propertyId) =>
        MapiKnownProperties.PidTag.AttachData.MatchesIdentity(propertyId);

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
        uint byteCount = cursor.ReadUInt32();
        if (byteCount > int.MaxValue) throw new InvalidDataException("TNEF named-property string is too large.");
        byte[] bytes = cursor.ReadBytes((int)byteCount);
        cursor.Align4();
        return new MapiNamedProperty(propertySet,
            Encoding.Unicode.GetString(bytes, 0, bytes.Length - bytes.Length % 2).TrimEnd('\0'));
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

    private static void WriteProperty(Stream output, MapiProperty property, int codePage,
        IList<EmailDiagnostic> diagnostics, string location) {
        using (var encoded = new MemoryStream()) {
            try {
                WriteEncodedProperty(encoded, property, codePage, useReplacementEncoding: false);
            } catch (EncoderFallbackException) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_TNEF_MAPI_STRING8_CHARACTER_UNENCODABLE",
                    string.Concat("String8 property 0x", property.PropertyId.ToString("X4", CultureInfo.InvariantCulture),
                        " contains characters that code page ", codePage.ToString(CultureInfo.InvariantCulture),
                        " cannot represent; replacement encoding was used."),
                    EmailDiagnosticSeverity.Warning, location));
                encoded.SetLength(0);
                WriteEncodedProperty(encoded, property, codePage, useReplacementEncoding: true);
            }
            encoded.WriteTo(output);
        }
    }

    private static void WriteEncodedProperty(Stream output, MapiProperty property, int codePage,
        bool useReplacementEncoding) {
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
                var item = new MapiProperty(propertyId, itemType, value) { RawData = multiple ? null : property.RawData };
                byte[] bytes = EncodeValue(item, codePage, useReplacementEncoding);
                if (IsVariableValue(itemType)) WriteUInt32(output, unchecked((uint)bytes.Length));
                output.Write(bytes, 0, bytes.Length);
                Pad4(output);
            }
        } else {
            byte[] bytes = EncodeFixedValue(property, codePage, useReplacementEncoding);
            output.Write(bytes, 0, bytes.Length);
            Pad4(output);
        }
    }

    private static byte[] EncodeFixedValue(MapiProperty property, int codePage, bool useReplacementEncoding) {
        byte[] encoded = EncodeValue(property, codePage, useReplacementEncoding);
        int size = GetFixedSize(property.PropertyType);
        if (encoded.Length == size) return encoded;
        var fixedBytes = new byte[size];
        Buffer.BlockCopy(encoded, 0, fixedBytes, 0, Math.Min(encoded.Length, size));
        return fixedBytes;
    }

    private static void WriteNamedProperty(Stream output, MapiNamedProperty name) {
        byte[] guid = name.PropertySet.ToByteArray();
        output.Write(guid, 0, guid.Length);
        WriteUInt32(output, name.Name == null ? 0U : 1U);
        if (name.Name == null) {
            WriteUInt32(output, name.LocalId.GetValueOrDefault());
        } else {
            string text = string.Concat(name.Name, "\0");
            byte[] bytes = Encoding.Unicode.GetBytes(text);
            WriteUInt32(output, unchecked((uint)bytes.Length));
            output.Write(bytes, 0, bytes.Length);
            Pad4(output);
        }
    }

    private static byte[] EncodeValue(MapiProperty property, int codePage, bool useReplacementEncoding) {
        if (property.PropertyType == MapiPropertyType.String8) {
            if (property.RawData != null) return (byte[])property.RawData.Clone();
            string text = string.Concat(Convert.ToString(property.Value, CultureInfo.InvariantCulture) ?? string.Empty, "\0");
            return useReplacementEncoding
                ? MsgValueWriter.EncodeString8WithReplacement(text, codePage)
                : MsgValueWriter.EncodeString8(text, codePage);
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
            case MapiPropertyType.Unspecified:
            case MapiPropertyType.Null: return 0;
            case MapiPropertyType.Integer16:
            case MapiPropertyType.Boolean: return 4;
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

    private static void SkipNamedProperty(PreflightCursor cursor) {
        cursor.Skip(16);
        uint kind = cursor.ReadUInt32();
        if (kind == 0) {
            cursor.Skip(4);
            return;
        }
        if (kind != 1) throw new InvalidDataException("Unknown TNEF named-property kind.");
        uint byteCount = cursor.ReadUInt32();
        if (byteCount > int.MaxValue) throw new InvalidDataException("TNEF named-property string is too large.");
        cursor.Skip((int)byteCount);
        cursor.Align4();
    }

    private sealed class PreflightCursor {
        private readonly byte[] _data;
        private readonly int _origin;
        private readonly int _end;

        internal PreflightCursor(byte[] data, int offset, int count) {
            if (offset < 0 || count < 0 || offset > data.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
            _data = data;
            _origin = offset;
            _end = offset + count;
            Position = offset;
        }

        internal int Position { get; private set; }

        internal int Remaining => _end - Position;

        internal uint ReadUInt32() {
            if (Position > _end - 4) throw new InvalidDataException("TNEF MAPI property data is truncated.");
            uint value = MsgBinary.ReadUInt32(_data, Position);
            Position += 4;
            return value;
        }

        internal void Skip(int count) {
            if (count < 0 || Position > _end - count) throw new InvalidDataException("TNEF MAPI property data is truncated.");
            Position += count;
        }

        internal void Align4() {
            int remainder = (Position - _origin) % 4;
            if (remainder != 0) Skip(4 - remainder);
        }
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
