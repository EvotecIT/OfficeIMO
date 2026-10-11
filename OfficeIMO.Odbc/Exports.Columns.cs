using System.Runtime.InteropServices;

namespace OfficeIMO.Odbc;

internal static unsafe partial class Exports {
    [UnmanagedCallersOnly(EntryPoint = "SQLDescribeColW")]
    internal static short SQLDescribeColW(nint id, ushort ordinal, nint name, short capacity, short* nameLength, short* type,
        nuint* size, short* scale, short* nullable) => DescribeColumn(id, ordinal, name, capacity, nameLength, type, size, scale, nullable, true);
    [UnmanagedCallersOnly(EntryPoint = "SQLDescribeCol")]
    internal static short SQLDescribeCol(nint id, ushort ordinal, nint name, short capacity, short* nameLength, short* type,
        nuint* size, short* scale, short* nullable) => DescribeColumn(id, ordinal, name, capacity, nameLength, type, size, scale, nullable, false);

    private static short DescribeColumn(nint id, ushort ordinal, nint name, short capacity, short* nameLength, short* type,
        nuint* size, short* scale, short* nullable, bool wide) => DriverHandles.With<StatementHandle>(id, statement => {
        short result = EnsureMetadata(statement);
        if (result != 0) return result;
        var reader = statement.Reader!;
        if (ordinal == 0 || ordinal > reader.FieldCount) return statement.Fail("07009", "Invalid column ordinal.");
        short sqlType = NativeValue.SqlType(reader, ordinal - 1);
        if (type != null) *type = sqlType;
        if (size != null) *size = NativeValue.ColumnSize(sqlType);
        if (scale != null) *scale = sqlType == 93 ? (short)7 : (short)0;
        if (nullable != null) *nullable = 1;
        result = NativeText.Write(reader.GetName(ordinal - 1), name, capacity, nameLength, wide);
        if (result == 1) statement.Diagnostics.Add(new("01004", "Column name was truncated."));
        return result;
    });

    [UnmanagedCallersOnly(EntryPoint = "SQLColAttributeW")]
    internal static short SQLColAttributeW(nint id, ushort ordinal, ushort field, nint text, short capacity, short* textLength, nint* number) =>
        ColumnAttribute(id, ordinal, field, text, capacity, textLength, number, true);
    [UnmanagedCallersOnly(EntryPoint = "SQLColAttribute")]
    internal static short SQLColAttribute(nint id, ushort ordinal, ushort field, nint text, short capacity, short* textLength, nint* number) =>
        ColumnAttribute(id, ordinal, field, text, capacity, textLength, number, false);

    private static short ColumnAttribute(nint id, ushort ordinal, ushort field, nint text, short capacity, short* textLength, nint* number, bool wide) =>
        DriverHandles.With<StatementHandle>(id, statement => {
            short result = EnsureMetadata(statement);
            if (result != 0) return result;
            var reader = statement.Reader!;
            if (field is 0 or 1001) { if (number != null) *number = reader.FieldCount; return 0; }
            if (ordinal == 0 || ordinal > reader.FieldCount) return statement.Fail("07009", "Invalid column ordinal.");
            short sqlType = NativeValue.SqlType(reader, ordinal - 1);
            string? value = field switch {
                1 or 18 or 1011 => reader.GetName(ordinal - 1),
                14 => NativeValue.TypeName(sqlType),
                15 or 16 or 17 or 22 or 23 => "", _ => null
            };
            if (value is not null) {
                short units = 0;
                result = NativeText.Write(value, text, wide ? capacity / 2 : capacity, &units, wide);
                if (textLength != null) *textLength = checked((short)(units * (wide ? 2 : 1)));
                if (result == 1) statement.Diagnostics.Add(new("01004", "Column attribute was truncated."));
                return result;
            }
            if (number == null) return statement.Fail("HY009", "Numeric attribute requires an output buffer.");
            *number = field switch {
                2 => sqlType,
                1002 => sqlType == 93 ? 9 : sqlType,
                3 or 6 or 1003 => (nint)NativeValue.ColumnSize(sqlType),
                4 or 1005 => NativeValue.Precision(sqlType),
                5 or 1006 => sqlType == 93 ? 7 : 0,
                7 or 1008 => 1,
                8 or 9 or 10 or 11 or 35 => 0,
                12 => sqlType == -9 ? 1 : 0,
                13 => 3,
                1007 => sqlType == 93 ? 3 : 0,
                1013 => NativeValue.OctetLength(sqlType),
                _ => -999
            };
            return *number == -999 ? statement.Fail("HYC00", "Unsupported column attribute.") : (short)0;
        });
}
