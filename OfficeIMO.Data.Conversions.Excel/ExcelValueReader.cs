using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Adapts row values to what a worksheet cell can hold, reporting each change: numbers beyond 15 significant digits
/// become text (a 16-digit identifier would otherwise show as 1.23E+15), date-times lose their offset, and fractional
/// seconds finer than a millisecond are rounded by Excel.
/// </summary>
internal sealed class ExcelValueReader : DbDataReader {
    private readonly StructuredDataReader _rows;

    public ExcelValueReader(StructuredDataReader rows) {
        _rows = rows;
    }

    public override object GetValue(int ordinal) {
        object value = _rows.GetValue(ordinal);
        switch (value) {
            case long integer when Math.Abs((double)integer) >= 1e15:
                return Precision(integer.ToString(CultureInfo.InvariantCulture));
            case decimal number when SignificantDigits(number) > 15:
                return Precision(number.ToString(CultureInfo.InvariantCulture));
            case DateTimeOffset offset:
                _rows.Context.Report(OfficeDiagnosticCodes.DatetimeOffsetDropped, OfficeConversionLossKind.Approximation,
                    "Excel dates have no time zone; date-times kept their clock time and dropped the offset.");
                return Sub(offset.DateTime);
            case DateTime date:
                return Sub(date);
            default:
                return value;
        }
    }

    private object Sub(DateTime date) {
        if (date.Ticks % TimeSpan.TicksPerMillisecond != 0) {
            _rows.Context.Report(OfficeDiagnosticCodes.DatetimePrecision, OfficeConversionLossKind.Approximation,
                "Excel stores time to the millisecond; finer fractional seconds were rounded.");
        }

        return date;
    }

    private string Precision(string text) {
        _rows.Context.Report(OfficeDiagnosticCodes.NumberPrecision, OfficeConversionLossKind.Approximation,
            "Numbers with more than 15 significant digits were written as text, because Excel numbers keep 15.");
        return text;
    }

    private static int SignificantDigits(decimal number) {
        string digits = decimal.Abs(number).ToString(CultureInfo.InvariantCulture).Replace(".", string.Empty).TrimStart('0');
        return digits.TrimEnd('0').Length == 0 ? 0 : digits.Length;
    }

    public override bool Read() => _rows.Read();
    public override bool NextResult() => _rows.NextResult();
    public override int FieldCount => _rows.FieldCount;
    public override string GetName(int ordinal) => _rows.GetName(ordinal);
    public override int GetOrdinal(string name) => _rows.GetOrdinal(name);

    [UnconditionalSuppressMessage("Trimming", "IL2063", Justification = "Column types are scalar value types and string, returned as data.")]
    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
    public override Type GetFieldType(int ordinal) => typeof(object);

    public override string GetDataTypeName(int ordinal) => "object";
    public override int GetValues(object[] values) {
        int count = Math.Min(values.Length, FieldCount);
        for (int i = 0; i < count; i++) values[i] = GetValue(i);
        return count;
    }

    public override bool IsDBNull(int ordinal) => _rows.IsDBNull(ordinal);
    public override object this[int ordinal] => GetValue(ordinal);
    public override object this[string name] => GetValue(GetOrdinal(name));
    public override bool GetBoolean(int ordinal) => _rows.GetBoolean(ordinal);
    public override byte GetByte(int ordinal) => _rows.GetByte(ordinal);
    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) => _rows.GetBytes(ordinal, dataOffset, buffer, bufferOffset, length);
    public override char GetChar(int ordinal) => _rows.GetChar(ordinal);
    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) => _rows.GetChars(ordinal, dataOffset, buffer, bufferOffset, length);
    public override DateTime GetDateTime(int ordinal) => _rows.GetDateTime(ordinal);
    public override decimal GetDecimal(int ordinal) => _rows.GetDecimal(ordinal);
    public override double GetDouble(int ordinal) => _rows.GetDouble(ordinal);
    public override float GetFloat(int ordinal) => _rows.GetFloat(ordinal);
    public override Guid GetGuid(int ordinal) => _rows.GetGuid(ordinal);
    public override short GetInt16(int ordinal) => _rows.GetInt16(ordinal);
    public override int GetInt32(int ordinal) => _rows.GetInt32(ordinal);
    public override long GetInt64(int ordinal) => _rows.GetInt64(ordinal);
    public override string GetString(int ordinal) => _rows.GetString(ordinal);
    public override IEnumerator GetEnumerator() => new DbEnumerator(this, closeReader: false);
    public override int Depth => 0;
    public override bool HasRows => _rows.HasRows;
    public override bool IsClosed => _rows.IsClosed;
    public override int RecordsAffected => -1;
}
