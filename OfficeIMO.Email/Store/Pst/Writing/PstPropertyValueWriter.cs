using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal static class PstPropertyValueWriter {
    static PstPropertyValueWriter() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    internal static bool IsInline(MapiPropertyType type) =>
        type == MapiPropertyType.Null ||
        type == MapiPropertyType.Integer16 ||
        type == MapiPropertyType.Integer32 ||
        type == MapiPropertyType.ErrorCode ||
        type == MapiPropertyType.Floating32 ||
        type == MapiPropertyType.Boolean;

    internal static uint EncodeInline(MapiProperty property) {
        object? value = property.Value;
        switch (property.PropertyType) {
            case MapiPropertyType.Null:
                return 0;
            case MapiPropertyType.Integer16:
                return unchecked((ushort)Convert.ToInt16(value ?? 0, CultureInfo.InvariantCulture));
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
                return unchecked((uint)Convert.ToInt32(value ?? 0, CultureInfo.InvariantCulture));
            case MapiPropertyType.Floating32:
                return BitConverter.ToUInt32(BitConverter.GetBytes(
                    Convert.ToSingle(value ?? 0F, CultureInfo.InvariantCulture)), 0);
            case MapiPropertyType.Boolean:
                return Convert.ToBoolean(value ?? false, CultureInfo.InvariantCulture) ? 1U : 0U;
            default:
                throw new NotSupportedException("The MAPI property type is not a four-byte inline value.");
        }
    }

    internal static byte[] EncodeVariable(MapiProperty property, int codePage) {
        if (property.RawData != null) return (byte[])property.RawData.Clone();
        object? value = property.Value;
        switch (property.PropertyType) {
            case MapiPropertyType.Floating64:
            case MapiPropertyType.FloatingTime:
                return BitConverter.GetBytes(Convert.ToDouble(value ?? 0D, CultureInfo.InvariantCulture));
            case MapiPropertyType.Currency:
                return BitConverter.GetBytes(ToCurrency(value));
            case MapiPropertyType.Integer64:
                return BitConverter.GetBytes(Convert.ToInt64(value ?? 0L, CultureInfo.InvariantCulture));
            case MapiPropertyType.Time:
                return BitConverter.GetBytes(ToFileTime(value));
            case MapiPropertyType.Guid:
                return value is Guid guid ? guid.ToByteArray() : Guid.Empty.ToByteArray();
            case MapiPropertyType.Unicode:
                return Encoding.Unicode.GetBytes(string.Concat(Convert.ToString(value, CultureInfo.InvariantCulture), "\0"));
            case MapiPropertyType.String8:
                return GetEncoding(codePage).GetBytes(string.Concat(Convert.ToString(value, CultureInfo.InvariantCulture), "\0"));
            case MapiPropertyType.Binary:
            case MapiPropertyType.Object:
            case MapiPropertyType.Unspecified:
                return value as byte[] ?? Array.Empty<byte>();
            case MapiPropertyType.MultipleInteger16:
                return EncodeFixedValues(value, 2, item => BitConverter.GetBytes(Convert.ToInt16(item, CultureInfo.InvariantCulture)));
            case MapiPropertyType.MultipleInteger32:
                return EncodeFixedValues(value, 4, item => BitConverter.GetBytes(Convert.ToInt32(item, CultureInfo.InvariantCulture)));
            case MapiPropertyType.MultipleFloating32:
                return EncodeFixedValues(value, 4, item => BitConverter.GetBytes(Convert.ToSingle(item, CultureInfo.InvariantCulture)));
            case MapiPropertyType.MultipleFloating64:
            case MapiPropertyType.MultipleFloatingTime:
                return EncodeFixedValues(value, 8, item => BitConverter.GetBytes(Convert.ToDouble(item, CultureInfo.InvariantCulture)));
            case MapiPropertyType.MultipleCurrency:
                return EncodeFixedValues(value, 8, item => BitConverter.GetBytes(ToCurrency(item)));
            case MapiPropertyType.MultipleInteger64:
                return EncodeFixedValues(value, 8, item => BitConverter.GetBytes(Convert.ToInt64(item, CultureInfo.InvariantCulture)));
            case MapiPropertyType.MultipleTime:
                return EncodeFixedValues(value, 8, item => BitConverter.GetBytes(ToFileTime(item)));
            case MapiPropertyType.MultipleGuid:
                return EncodeFixedValues(value, 16, item => item is Guid itemGuid ? itemGuid.ToByteArray() : Guid.Empty.ToByteArray());
            case MapiPropertyType.MultipleUnicode:
                return EncodeVariableValues(value, item => Encoding.Unicode.GetBytes(
                    string.Concat(Convert.ToString(item, CultureInfo.InvariantCulture), "\0")));
            case MapiPropertyType.MultipleString8:
                Encoding encoding = GetEncoding(codePage);
                return EncodeVariableValues(value, item => encoding.GetBytes(
                    string.Concat(Convert.ToString(item, CultureInfo.InvariantCulture), "\0")));
            case MapiPropertyType.MultipleBinary:
                return EncodeVariableValues(value, item => item as byte[] ?? Array.Empty<byte>());
            default:
                throw new NotSupportedException(string.Concat("Unsupported MAPI property type 0x",
                    ((ushort)property.PropertyType).ToString("X4", CultureInfo.InvariantCulture), "."));
        }
    }

    private static byte[] EncodeFixedValues(object? source, int elementSize,
        Func<object?, byte[]> encode) {
        object?[] values = Enumerate(source);
        var result = new byte[checked(values.Length * elementSize)];
        for (int index = 0; index < values.Length; index++) {
            byte[] encoded = encode(values[index]);
            if (encoded.Length != elementSize) throw new InvalidDataException("A fixed MAPI value has an invalid size.");
            Buffer.BlockCopy(encoded, 0, result, index * elementSize, elementSize);
        }
        return result;
    }

    private static byte[] EncodeVariableValues(object? source, Func<object?, byte[]> encode) {
        object?[] values = Enumerate(source);
        byte[][] encoded = values.Select(encode).ToArray();
        int headerLength = checked(4 + encoded.Length * 4);
        int length = encoded.Aggregate(headerLength, (current, item) => checked(current + item.Length));
        var result = new byte[length];
        PstBinary.WriteUInt32(result, 0, checked((uint)encoded.Length));
        int cursor = headerLength;
        for (int index = 0; index < encoded.Length; index++) {
            PstBinary.WriteUInt32(result, 4 + index * 4, checked((uint)cursor));
            Buffer.BlockCopy(encoded[index], 0, result, cursor, encoded[index].Length);
            cursor += encoded[index].Length;
        }
        return result;
    }

    private static object?[] Enumerate(object? source) {
        if (source == null) return Array.Empty<object?>();
        if (source is string || source is byte[]) return new[] { source };
        if (source is System.Collections.IEnumerable enumerable) {
            var result = new List<object?>();
            foreach (object? item in enumerable) result.Add(item);
            return result.ToArray();
        }
        return new[] { source };
    }

    private static long ToFileTime(object? value) {
        if (value is DateTimeOffset offset) return offset.UtcDateTime.ToFileTimeUtc();
        if (value is DateTime dateTime) return dateTime.ToUniversalTime().ToFileTimeUtc();
        if (value == null) return 0;
        return Convert.ToInt64(value, CultureInfo.InvariantCulture);
    }

    private static long ToCurrency(object? value) => decimal.ToInt64(
        Convert.ToDecimal(value ?? 0m, CultureInfo.InvariantCulture) * 10000m);

    private static Encoding GetEncoding(int codePage) {
        try { return Encoding.GetEncoding(codePage > 0 ? codePage : 1252); }
        catch (ArgumentException) { return Encoding.GetEncoding(1252); }
    }
}
