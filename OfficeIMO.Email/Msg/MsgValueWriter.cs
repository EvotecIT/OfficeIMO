namespace OfficeIMO.Email;

internal static class MsgValueWriter {
    internal static bool IsVariable(MapiPropertyType type) {
        return type == MapiPropertyType.String8 || type == MapiPropertyType.Unicode ||
            type == MapiPropertyType.Binary || type == MapiPropertyType.Guid || type == MapiPropertyType.Object;
    }

    internal static bool IsMultiple(MapiPropertyType type) => (((ushort)type) & 0x1000) != 0;

    internal static byte[] EncodeScalar(MapiProperty property) {
        if (property.RawData != null && (property.Value == null || property.PropertyType == MapiPropertyType.String8)) {
            return (byte[])property.RawData.Clone();
        }
        object? value = property.Value;
        switch (property.PropertyType) {
            case MapiPropertyType.Unspecified:
            case MapiPropertyType.Null:
            case MapiPropertyType.Object:
                return Array.Empty<byte>();
            case MapiPropertyType.Integer16:
                return EncodeInt16(Convert.ToInt16(value, CultureInfo.InvariantCulture));
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
                return EncodeInt32(Convert.ToInt32(value, CultureInfo.InvariantCulture));
            case MapiPropertyType.Floating32:
                return EncodeSingle(Convert.ToSingle(value, CultureInfo.InvariantCulture));
            case MapiPropertyType.Floating64:
                return EncodeDouble(Convert.ToDouble(value, CultureInfo.InvariantCulture));
            case MapiPropertyType.Currency:
                return EncodeInt64(decimal.ToInt64(Convert.ToDecimal(value, CultureInfo.InvariantCulture) * 10000m));
            case MapiPropertyType.FloatingTime:
                return EncodeDouble(ConvertDateTime(value).ToOADate());
            case MapiPropertyType.Boolean:
                return EncodeInt16(Convert.ToBoolean(value, CultureInfo.InvariantCulture) ? (short)1 : (short)0);
            case MapiPropertyType.Integer64:
                return EncodeInt64(Convert.ToInt64(value, CultureInfo.InvariantCulture));
            case MapiPropertyType.String8:
                return EncodeWindows1252(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
            case MapiPropertyType.Unicode:
                return Encoding.Unicode.GetBytes(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
            case MapiPropertyType.Time:
                return EncodeInt64(ConvertDateTime(value).ToUniversalTime().ToFileTimeUtc());
            case MapiPropertyType.Guid:
                return value is Guid guid ? guid.ToByteArray() : Guid.Parse(Convert.ToString(value, CultureInfo.InvariantCulture)!).ToByteArray();
            case MapiPropertyType.Binary:
                return value is byte[] bytes ? (byte[])bytes.Clone() : Array.Empty<byte>();
            default:
                return property.RawData == null ? Array.Empty<byte>() : (byte[])property.RawData.Clone();
        }
    }

    internal static object[] GetMultipleValues(MapiProperty property) {
        if (property.Value is object[] values) return values;
        if (property.Value is System.Collections.IEnumerable enumerable && !(property.Value is string) && !(property.Value is byte[])) {
            var result = new List<object>();
            foreach (object? item in enumerable) result.Add(item ?? string.Empty);
            return result.ToArray();
        }
        return Array.Empty<object>();
    }

    internal static MapiPropertyType GetMultipleItemType(MapiPropertyType type) => (MapiPropertyType)((ushort)type & 0x0fff);

    internal static byte[] EncodeFixedValue(MapiProperty property) {
        byte[] value = EncodeScalar(property);
        byte[] result = new byte[8];
        Buffer.BlockCopy(value, 0, result, 0, Math.Min(result.Length, value.Length));
        return result;
    }

    private static DateTime ConvertDateTime(object? value) {
        if (value is DateTimeOffset offset) return offset.UtcDateTime;
        if (value is DateTime date) return date.Kind == DateTimeKind.Utc ? date : date.ToUniversalTime();
        throw new ArgumentException("A MAPI time property requires DateTime or DateTimeOffset.");
    }

    private static byte[] EncodeInt16(short value) {
        byte[] bytes = new byte[2];
        MsgBinary.WriteUInt16(bytes, 0, unchecked((ushort)value));
        return bytes;
    }

    private static byte[] EncodeInt32(int value) {
        byte[] bytes = new byte[4];
        MsgBinary.WriteUInt32(bytes, 0, unchecked((uint)value));
        return bytes;
    }

    private static byte[] EncodeInt64(long value) {
        byte[] bytes = new byte[8];
        MsgBinary.WriteUInt64(bytes, 0, unchecked((ulong)value));
        return bytes;
    }

    private static byte[] EncodeSingle(float value) {
        byte[] bytes = BitConverter.GetBytes(value);
        if (!BitConverter.IsLittleEndian) Array.Reverse(bytes);
        return bytes;
    }

    private static byte[] EncodeDouble(double value) {
        byte[] bytes = BitConverter.GetBytes(value);
        if (!BitConverter.IsLittleEndian) Array.Reverse(bytes);
        return bytes;
    }

    private static byte[] EncodeWindows1252(string value) {
        const string replacements = "\u20AC\u0081\u201A\u0192\u201E\u2026\u2020\u2021\u02C6\u2030\u0160\u2039\u0152\u008D\u017D\u008F" +
            "\u0090\u2018\u2019\u201C\u201D\u2022\u2013\u2014\u02DC\u2122\u0161\u203A\u0153\u009D\u017E\u0178";
        byte[] bytes = new byte[value.Length];
        for (int i = 0; i < value.Length; i++) {
            char character = value[i];
            int special = replacements.IndexOf(character);
            bytes[i] = character <= 0xff && !(character >= 0x80 && character <= 0x9f)
                ? (byte)character
                : special >= 0 ? (byte)(special + 0x80) : (byte)'?';
        }
        return bytes;
    }
}
