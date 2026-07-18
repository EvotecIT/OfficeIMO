namespace OfficeIMO.Email.Store;

internal static class EmailStoreScalarCodec {
    private const byte NullMarker = 0;
    private const byte StringMarker = 1;
    private const byte BooleanMarker = 2;
    private const byte Int32Marker = 3;
    private const byte Int64Marker = 4;
    private const byte DateTimeOffsetMarker = 5;
    private const byte FolderIdMarker = 6;
    private const byte ItemIdMarker = 7;
    private const byte EnumMarker = 8;

    internal static string Signature(object? value) {
        if (value == null) return "n";
        if (value is string text) return string.Concat("s", EncodeText(text));
        if (value is bool boolean) return boolean ? "b1" : "b0";
        if (value is int int32) return string.Concat("i", int32.ToString(CultureInfo.InvariantCulture));
        if (value is long int64) return string.Concat("l", int64.ToString(CultureInfo.InvariantCulture));
        if (value is DateTimeOffset dateTime) return string.Concat("d", dateTime.ToString("O", CultureInfo.InvariantCulture));
        if (value is EmailStoreFolderId folderId) return string.Concat("f", EncodeText(folderId.Value));
        if (value is EmailStoreItemId itemId) return string.Concat("m", EncodeText(itemId.Value));
        Type type = value.GetType();
        if (type.IsEnum) {
            return string.Concat("e", EncodeText(type.FullName ?? type.Name), ":",
                Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture));
        }
        throw new NotSupportedException(string.Concat("Store query tokens do not support scalar type ", type.FullName, "."));
    }

    internal static byte[] Serialize(object? value) {
        using (var stream = new MemoryStream())
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
            if (value == null) {
                writer.Write(NullMarker);
            } else if (value is string text) {
                writer.Write(StringMarker);
                WriteString(writer, text);
            } else if (value is bool boolean) {
                writer.Write(BooleanMarker);
                writer.Write(boolean);
            } else if (value is int int32) {
                writer.Write(Int32Marker);
                writer.Write(int32);
            } else if (value is long int64) {
                writer.Write(Int64Marker);
                writer.Write(int64);
            } else if (value is DateTimeOffset dateTime) {
                writer.Write(DateTimeOffsetMarker);
                writer.Write(dateTime.Ticks);
                writer.Write((short)dateTime.Offset.TotalMinutes);
            } else if (value is EmailStoreFolderId folderId) {
                writer.Write(FolderIdMarker);
                WriteString(writer, folderId.Value);
            } else if (value is EmailStoreItemId itemId) {
                writer.Write(ItemIdMarker);
                WriteString(writer, itemId.Value);
            } else if (value.GetType().IsEnum) {
                writer.Write(EnumMarker);
                writer.Write(Convert.ToInt64(value, CultureInfo.InvariantCulture));
            } else {
                throw new NotSupportedException(string.Concat("Store continuation tokens do not support scalar type ", value.GetType().FullName, "."));
            }
            writer.Flush();
            return stream.ToArray();
        }
    }

    internal static object? Deserialize(byte[] payload, Type declaredType) {
        if (payload == null) throw new ArgumentNullException(nameof(payload));
        if (declaredType == null) throw new ArgumentNullException(nameof(declaredType));
        Type targetType = Nullable.GetUnderlyingType(declaredType) ?? declaredType;
        using (var stream = new MemoryStream(payload, writable: false))
        using (var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true)) {
            byte marker = reader.ReadByte();
            object? value;
            switch (marker) {
                case NullMarker:
                    value = null;
                    break;
                case StringMarker:
                    RequireType(targetType, typeof(string));
                    value = ReadString(reader, 65_536);
                    break;
                case BooleanMarker:
                    RequireType(targetType, typeof(bool));
                    value = reader.ReadBoolean();
                    break;
                case Int32Marker:
                    RequireType(targetType, typeof(int));
                    value = reader.ReadInt32();
                    break;
                case Int64Marker:
                    RequireType(targetType, typeof(long));
                    value = reader.ReadInt64();
                    break;
                case DateTimeOffsetMarker:
                    RequireType(targetType, typeof(DateTimeOffset));
                    long ticks = reader.ReadInt64();
                    short offsetMinutes = reader.ReadInt16();
                    if (offsetMinutes < -14 * 60 || offsetMinutes > 14 * 60) {
                        throw new InvalidDataException("A continuation token contains an invalid UTC offset.");
                    }
                    try {
                        value = new DateTimeOffset(ticks, TimeSpan.FromMinutes(offsetMinutes));
                    } catch (ArgumentOutOfRangeException exception) {
                        throw new InvalidDataException("A continuation token contains an invalid date-time value.", exception);
                    }
                    break;
                case FolderIdMarker:
                    RequireType(targetType, typeof(EmailStoreFolderId));
                    value = new EmailStoreFolderId(ReadString(reader, 65_536));
                    break;
                case ItemIdMarker:
                    RequireType(targetType, typeof(EmailStoreItemId));
                    value = new EmailStoreItemId(ReadString(reader, 65_536));
                    break;
                case EnumMarker:
                    if (!targetType.IsEnum) throw new InvalidDataException("A continuation token enum does not match the query field type.");
                    value = Enum.ToObject(targetType, reader.ReadInt64());
                    break;
                default:
                    throw new InvalidDataException("The continuation token contains an unsupported scalar marker.");
            }
            if (stream.Position != stream.Length) throw new InvalidDataException("The continuation token scalar contains trailing data.");
            if (value == null && targetType.IsValueType && Nullable.GetUnderlyingType(declaredType) == null) {
                throw new InvalidDataException("A continuation token contains null for a non-nullable query field.");
            }
            return value;
        }
    }

    internal static void WriteString(BinaryWriter writer, string value) {
        byte[] bytes = Encoding.UTF8.GetBytes(value);
        writer.Write(bytes.Length);
        writer.Write(bytes);
    }

    internal static string ReadString(BinaryReader reader, int maxBytes) {
        int length = reader.ReadInt32();
        if (length < 0 || length > maxBytes || length > reader.BaseStream.Length - reader.BaseStream.Position) {
            throw new InvalidDataException("A Store token string length is invalid.");
        }
        byte[] bytes = reader.ReadBytes(length);
        if (bytes.Length != length) throw new EndOfStreamException();
        return Encoding.UTF8.GetString(bytes);
    }

    private static string EncodeText(string value) => Convert.ToBase64String(Encoding.UTF8.GetBytes(value));

    private static void RequireType(Type actual, Type expected) {
        if (actual != expected) throw new InvalidDataException("A continuation token scalar does not match the query field type.");
    }
}
