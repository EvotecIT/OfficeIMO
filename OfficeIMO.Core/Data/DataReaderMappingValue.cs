#nullable enable
using System;
using System.Data.Common;

namespace OfficeIMO.Data;

/// <summary>Preserves a provider's original value and numeric getter result across conversion and row snapshots.</summary>
internal sealed class DataReaderMappingValue {
    private DataReaderMappingValue(object original, double numeric) {
        Original = original;
        Numeric = numeric;
    }

    internal object Original { get; }
    internal double Numeric { get; }

    internal static object? Read(DbDataReader reader, int ordinal, Type targetType) {
        object? value = reader.GetValue(ordinal);
        Type effective = Nullable.GetUnderlyingType(targetType) ?? targetType;
        if (value is not DateTime || !IsNumeric(effective)) return value;
        try {
            return new DataReaderMappingValue(value, reader.GetDouble(ordinal));
        } catch (InvalidCastException) {
        } catch (FormatException) {
        } catch (OverflowException) {
        } catch (NotSupportedException) {
        } catch (NotImplementedException) {
        }
        return value;
    }

    internal static bool IsNumeric(Type type) => type == typeof(double) || type == typeof(float)
        || type == typeof(decimal) || type == typeof(int) || type == typeof(long)
        || type == typeof(short) || type == typeof(byte) || type == typeof(sbyte)
        || type == typeof(ushort) || type == typeof(uint) || type == typeof(ulong);
}
