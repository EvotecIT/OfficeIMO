using System.Collections;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.CSV;

/// <summary>
/// Reads a <see cref="DataTable"/> for the CSV writers the way <see cref="DataTableReader"/> presents it: the table's
/// rows in order without the deleted ones, each value from the row's default version (expression columns computed),
/// and each field typed by its column. It walks the row collection with its enumerator and reads values through the
/// column objects, so it skips the reader's per-call validation and its lookup of the current row by index.
/// </summary>
internal sealed class CsvDataTableRowReader : DbDataReader
{
    private readonly DataTable _table;
    private readonly DataColumn[] _columns;
    private IEnumerator? _rows;
    private DataRow? _row;
    private bool _closed;

    internal CsvDataTableRowReader(DataTable table)
    {
        _table = table;
        _columns = new DataColumn[table.Columns.Count];
        table.Columns.CopyTo(_columns, 0);
    }

    private DataRow Row => _row ?? throw new InvalidOperationException("No current row.");

    public override bool Read()
    {
        if (_closed) throw new InvalidOperationException("The reader is closed.");
        _rows ??= _table.Rows.GetEnumerator();
        while (_rows.MoveNext())
        {
            var row = (DataRow)_rows.Current!;
            if (row.RowState != DataRowState.Deleted)
            {
                _row = row;
                return true;
            }
        }

        _row = null;
        return false;
    }

    public override object GetValue(int ordinal) => Row[_columns[ordinal]];

    public override bool IsDBNull(int ordinal) => Row.IsNull(_columns[ordinal]);

    public override int FieldCount => _columns.Length;

    public override string GetName(int ordinal) => _columns[ordinal].ColumnName;

#if NET5_0_OR_GREATER
    [return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
#endif
    public override Type GetFieldType(int ordinal) => _columns[ordinal].DataType;

    public override string GetDataTypeName(int ordinal) => _columns[ordinal].DataType.Name;

    public override int GetOrdinal(string name) => _table.Columns.IndexOf(name) is var index and >= 0
        ? index
        : throw new ArgumentException($"Column '{name}' does not belong to the table.", nameof(name));

    public override object this[int ordinal] => GetValue(ordinal);

    public override object this[string name] => GetValue(GetOrdinal(name));

    public override int GetValues(object[] values)
    {
        int count = Math.Min(values.Length, _columns.Length);
        for (int ordinal = 0; ordinal < count; ordinal++) values[ordinal] = GetValue(ordinal);
        return count;
    }

    public override bool GetBoolean(int ordinal) => (bool)GetValue(ordinal);

    public override byte GetByte(int ordinal) => (byte)GetValue(ordinal);

    public override char GetChar(int ordinal) => (char)GetValue(ordinal);

    public override DateTime GetDateTime(int ordinal) => (DateTime)GetValue(ordinal);

    public override decimal GetDecimal(int ordinal) => (decimal)GetValue(ordinal);

    public override double GetDouble(int ordinal) => (double)GetValue(ordinal);

    public override float GetFloat(int ordinal) => (float)GetValue(ordinal);

    public override Guid GetGuid(int ordinal) => (Guid)GetValue(ordinal);

    public override short GetInt16(int ordinal) => (short)GetValue(ordinal);

    public override int GetInt32(int ordinal) => (int)GetValue(ordinal);

    public override long GetInt64(int ordinal) => (long)GetValue(ordinal);

    public override string GetString(int ordinal) => (string)GetValue(ordinal);

    public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) =>
        throw new NotSupportedException();

    public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) =>
        throw new NotSupportedException();

    public override int Depth => 0;

    public override bool HasRows => _table.Rows.Count > 0;

    public override bool IsClosed => _closed;

    public override int RecordsAffected => 0;

    public override bool NextResult() => false;

    public override IEnumerator GetEnumerator() => new DbEnumerator(this);

    public override void Close() => _closed = true;
}
