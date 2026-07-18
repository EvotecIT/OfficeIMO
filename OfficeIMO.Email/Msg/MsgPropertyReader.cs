using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Email;

internal enum MsgPropertyStreamKind {
    TopLevel,
    EmbeddedMessage,
    ChildObject
}

internal static class MsgPropertyReader {
    internal static List<MapiProperty> Read(OfficeCompoundFile compound, string prefix, MsgPropertyStreamKind kind,
        MsgNamedPropertyMap names, MsgParserState state, MapiStringEncodingContext? inheritedEncoding,
        out MapiStringEncodingContext encoding) {
        string propertyPath = MsgBinary.CombinePath(prefix, "__properties_version1.0");
        var result = new List<MapiProperty>();
        if (!compound.Streams.TryGetValue(propertyPath, out byte[]? propertyStream)) {
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_PROPERTIES_MISSING",
                "The required MSG property stream is missing.", EmailDiagnosticSeverity.Error, prefix));
            encoding = inheritedEncoding ?? MapiStringEncodingContext.Resolve(Array.Empty<byte>(), 0, null);
            return result;
        }

        int headerLength = kind == MsgPropertyStreamKind.TopLevel ? 32 :
            kind == MsgPropertyStreamKind.EmbeddedMessage ? 24 : 8;
        if (propertyStream.Length < headerLength) {
            state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_PROPERTIES_TRUNCATED",
                "The MSG property stream header is truncated.", EmailDiagnosticSeverity.Error, propertyPath));
            encoding = inheritedEncoding ?? MapiStringEncodingContext.Resolve(Array.Empty<byte>(), 0, null);
            return result;
        }
        int remainder = propertyStream.Length - headerLength;
        if (remainder % 16 != 0) {
            int completeLength = headerLength + remainder / 16 * 16;
            bool hasNonZeroTail = propertyStream.Skip(completeLength).Any(value => value != 0);
            if (hasNonZeroTail) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_PROPERTIES_MISALIGNED",
                    "The MSG property stream has a trailing nonzero partial entry.", EmailDiagnosticSeverity.Warning, propertyPath));
            }
        }

        encoding = MapiStringEncodingContext.Resolve(propertyStream, headerLength, inheritedEncoding);
        int count = remainder / 16;
        for (int index = 0; index < count; index++) {
            state.ThrowIfCancellationRequested();
            int offset = headerLength + index * 16;
            uint tag = MsgBinary.ReadUInt32(propertyStream, offset);
            uint flags = MsgBinary.ReadUInt32(propertyStream, offset + 4);
            ushort propertyId = unchecked((ushort)(tag >> 16));
            MapiPropertyType type = (MapiPropertyType)unchecked((ushort)tag);
            string valueName = string.Concat("__substg1.0_", tag.ToString("X8", CultureInfo.InvariantCulture));
            string valuePath = MsgBinary.CombinePath(prefix, valueName);
            object? value = null;
            byte[]? raw = null;

            try {
                if (IsVariable(type) || IsMultiple(type)) {
                    if (type == MapiPropertyType.Object) {
                        raw = null;
                    } else if (compound.Streams.TryGetValue(valuePath, out byte[]? streamBytes)) {
                        raw = IsMultiple(type) ? streamBytes : TrimStringTerminator(type, streamBytes);
                        value = IsMultiple(type)
                            ? DecodeMultiple(type, streamBytes, compound, prefix, valueName, encoding, state, valuePath)
                            : DecodeScalar(type, streamBytes, encoding, state, valuePath);
                    } else {
                        state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_PROPERTY_STREAM_MISSING",
                            string.Concat("Property stream ", valueName, " is missing."),
                            EmailDiagnosticSeverity.Warning, valuePath));
                    }
                } else {
                    raw = MsgBinary.Slice(propertyStream, offset + 8, 8);
                    value = DecodeScalar(type, raw, encoding, state, propertyPath);
                }
            } catch (Exception ex) when (ex is InvalidDataException || ex is ArgumentException || ex is OverflowException) {
                state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_PROPERTY_INVALID",
                    string.Concat("Property 0x", tag.ToString("X8", CultureInfo.InvariantCulture), " could not be decoded: ", ex.Message),
                    EmailDiagnosticSeverity.Warning, valuePath));
            }

            bool attachmentPayload = MapiKnownProperties.PidTag.AttachData.MatchesIdentity(propertyId) &&
                prefix.IndexOf("__attach_version1.0_#", StringComparison.OrdinalIgnoreCase) >= 0;
            state.CountProperty(attachmentPayload ? 0 : raw?.Length ?? 0);
            result.Add(new MapiProperty(propertyId, type, value, flags, names.Get(propertyId)) { RawData = raw });
        }
        return result;
    }

    private static byte[] TrimStringTerminator(MapiPropertyType type, byte[] value) {
        int terminatorLength = type == MapiPropertyType.Unicode ? 2 :
            type == MapiPropertyType.String8 ? 1 : 0;
        if (terminatorLength == 0 || value.Length < terminatorLength) return value;
        for (int index = value.Length - terminatorLength; index < value.Length; index++) {
            if (value[index] != 0) return value;
        }
        return MsgBinary.Slice(value, 0, value.Length - terminatorLength);
    }

    private static object? DecodeScalar(MapiPropertyType type, byte[] bytes, MapiStringEncodingContext encoding,
        MsgParserState state, string location) {
        switch (type) {
            case MapiPropertyType.Unspecified:
            case MapiPropertyType.Null:
                return null;
            case MapiPropertyType.Integer16:
                return MsgBinary.ReadInt16(bytes, 0);
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
                return MsgBinary.ReadInt32(bytes, 0);
            case MapiPropertyType.Floating32:
                return MsgBinary.ReadSingle(bytes, 0);
            case MapiPropertyType.Floating64:
                return MsgBinary.ReadDouble(bytes, 0);
            case MapiPropertyType.Currency:
                return MsgBinary.ReadInt64(bytes, 0) / 10000m;
            case MapiPropertyType.FloatingTime:
                return DateTime.FromOADate(MsgBinary.ReadDouble(bytes, 0));
            case MapiPropertyType.Boolean:
                return MsgBinary.ReadUInt16(bytes, 0) != 0;
            case MapiPropertyType.Integer64:
                return MsgBinary.ReadInt64(bytes, 0);
            case MapiPropertyType.String8:
                return encoding.Decode(bytes, state.Diagnostics, location).TrimEnd('\0');
            case MapiPropertyType.Unicode:
                return Encoding.Unicode.GetString(bytes, 0, bytes.Length - bytes.Length % 2).TrimEnd('\0');
            case MapiPropertyType.Time:
                return DecodeFileTime(MsgBinary.ReadInt64(bytes, 0));
            case MapiPropertyType.Guid:
                return bytes.Length < 16 ? throw new InvalidDataException("A GUID property is shorter than 16 bytes.") : new Guid(MsgBinary.Slice(bytes, 0, 16));
            case MapiPropertyType.Binary:
                return (byte[])bytes.Clone();
            default:
                return (byte[])bytes.Clone();
        }
    }

    private static object[] DecodeMultiple(MapiPropertyType type, byte[] lengthOrValueStream,
        OfficeCompoundFile compound, string prefix, string valueName, MapiStringEncodingContext encoding,
        MsgParserState state, string location) {
        if (type == MapiPropertyType.MultipleString8 || type == MapiPropertyType.MultipleUnicode ||
            type == MapiPropertyType.MultipleBinary) {
            int entrySize = type == MapiPropertyType.MultipleBinary ? 8 : 4;
            int count = lengthOrValueStream.Length / entrySize;
            var values = new object[count];
            MapiPropertyType scalarType = type == MapiPropertyType.MultipleBinary ? MapiPropertyType.Binary :
                type == MapiPropertyType.MultipleUnicode ? MapiPropertyType.Unicode : MapiPropertyType.String8;
            for (int index = 0; index < count; index++) {
                state.ThrowIfCancellationRequested();
                uint declaredLength = MsgBinary.ReadUInt32(lengthOrValueStream, index * entrySize);
                string itemName = string.Concat(valueName, "-", index.ToString("X8", CultureInfo.InvariantCulture));
                string itemPath = MsgBinary.CombinePath(prefix, itemName);
                if (!compound.Streams.TryGetValue(itemPath, out byte[]? itemBytes)) {
                    state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_MULTIVALUE_STREAM_MISSING",
                        string.Concat("Multiple-value stream ", itemName, " is missing."),
                        EmailDiagnosticSeverity.Warning, itemPath));
                    values[index] = Array.Empty<byte>();
                    continue;
                }
                if (declaredLength != itemBytes.Length) {
                    state.Diagnostics.Add(new EmailDiagnostic("EMAIL_MSG_MULTIVALUE_LENGTH_MISMATCH",
                        string.Concat("Declared length ", declaredLength.ToString(CultureInfo.InvariantCulture),
                            " differs from stream length ", itemBytes.Length.ToString(CultureInfo.InvariantCulture), "."),
                        EmailDiagnosticSeverity.Warning, itemPath));
                }
                state.CountDecodedBytes(itemBytes.Length);
                values[index] = DecodeScalar(scalarType, itemBytes, encoding, state, itemPath) ?? Array.Empty<byte>();
            }
            return values;
        }

        MapiPropertyType itemType = GetMultipleItemType(type);
        int itemSize = GetFixedSize(itemType);
        if (itemSize == 0 || lengthOrValueStream.Length % itemSize != 0) {
            throw new InvalidDataException("A fixed multiple-value property has an invalid stream length.");
        }
        var fixedValues = new object[lengthOrValueStream.Length / itemSize];
        for (int index = 0; index < fixedValues.Length; index++) {
            state.ThrowIfCancellationRequested();
            fixedValues[index] = DecodeScalar(itemType, MsgBinary.Slice(lengthOrValueStream, index * itemSize, itemSize),
                encoding, state, location) ?? 0;
        }
        return fixedValues;
    }

    private static MapiPropertyType GetMultipleItemType(MapiPropertyType type) {
        return (MapiPropertyType)((ushort)type & 0x0fff);
    }

    private static int GetFixedSize(MapiPropertyType type) {
        switch (type) {
            case MapiPropertyType.Integer16:
            case MapiPropertyType.Boolean:
                return 2;
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
            case MapiPropertyType.Floating32:
                return 4;
            case MapiPropertyType.Guid:
                return 16;
            default:
                return 8;
        }
    }

    private static bool IsVariable(MapiPropertyType type) {
        return type == MapiPropertyType.String8 || type == MapiPropertyType.Unicode ||
            type == MapiPropertyType.Binary || type == MapiPropertyType.Guid || type == MapiPropertyType.Object;
    }

    private static bool IsMultiple(MapiPropertyType type) => (((ushort)type) & 0x1000) != 0;

    private static DateTimeOffset DecodeFileTime(long fileTime) {
        return new DateTimeOffset(DateTime.FromFileTimeUtc(fileTime), TimeSpan.Zero);
    }
}
