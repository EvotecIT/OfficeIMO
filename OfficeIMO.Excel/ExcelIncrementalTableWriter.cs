using System.Threading;

namespace OfficeIMO.Excel;

/// <summary>Spools normalized scalar rows to disk while collecting columns in first-seen order.</summary>
/// <remarks>Complete freezes the schema and writes through the native workbook writer. Disposing an incomplete table aborts publication of its workbook.</remarks>
public sealed class ExcelIncrementalTableWriter : IDisposable {
    private readonly ExcelWorkbookWriter _owner;
    private readonly string _name;
    private readonly CancellationToken _cancellationToken;
    private readonly FileStream _spool;
    private readonly BinaryWriter _writer;
    private readonly List<string> _headers = new();
    private readonly Dictionary<string, int> _columns = new(StringComparer.OrdinalIgnoreCase);
    private readonly long? _maximumBytes;
    private int _rowCount;
    private bool _completed;
    private bool _failed;
    private bool _disposed;

    internal ExcelIncrementalTableWriter(ExcelWorkbookWriter owner, string name, long? maximumBytes, CancellationToken cancellationToken) {
        _owner = owner;
        _name = name;
        _maximumBytes = maximumBytes;
        _cancellationToken = cancellationToken;
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-Excel-Rows-" + Guid.NewGuid().ToString("N") + ".bin");
        _spool = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 65536,
            FileOptions.SequentialScan | FileOptions.DeleteOnClose);
        _writer = new BinaryWriter(_spool, System.Text.Encoding.UTF8, leaveOpen: true);
    }

    /// <summary>Copies one row immediately. Later rows may introduce columns; omitted columns become blank cells.</summary>
    public void WriteRow(IReadOnlyList<string> columnNames, IReadOnlyList<object?> values) {
        try {
            EnsureActive();
            ArgumentNullException.ThrowIfNull(columnNames);
            ArgumentNullException.ThrowIfNull(values);
            if (columnNames.Count != values.Count) throw new ArgumentException("Each column requires one value.", nameof(values));
            if (_rowCount >= ExcelA1.MaxRows) throw new InvalidOperationException("The worksheet row limit was exceeded.");
            DeclareColumns(columnNames);
            _writer.Write(columnNames.Count);
            for (int index = 0; index < columnNames.Count; index++) {
                _writer.Write(_columns[columnNames[index]]);
                WriteValue(values[index]);
            }
            _rowCount++;
            if (_maximumBytes.HasValue && _spool.Position > _maximumBytes.Value)
                throw new InvalidOperationException("The row spool exceeds MaxTemporaryPackageBytes.");
        } catch {
            _failed = true;
            _owner.AbortIncrementalTable(this);
            throw;
        }
    }

    /// <summary>Registers a known schema without writing a data row, including for empty readers and tables.</summary>
    public void DeclareColumns(IReadOnlyList<string> columnNames) {
        try {
            EnsureActive();
            ArgumentNullException.ThrowIfNull(columnNames);
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in columnNames) {
                if (string.IsNullOrWhiteSpace(name) || !used.Add(name))
                    throw new ArgumentException("Each row requires nonempty, unique column names.", nameof(columnNames));
                if (_columns.ContainsKey(name)) continue;
                if (_headers.Count >= ExcelA1.MaxColumns) throw new InvalidOperationException("The worksheet column limit was exceeded.");
                _columns.Add(name, _headers.Count);
                _headers.Add(name);
            }
        } catch {
            _failed = true;
            _owner.AbortIncrementalTable(this);
            throw;
        }
    }

    /// <summary>Writes the completed table with its final union schema and returns the occupied native range.</summary>
    public ExcelDataSetImportResult Complete() {
        try {
            EnsureActive();
            _writer.Flush();
            _spool.Position = 0;
            ExcelDataSetImportResult result = _owner.CompleteIncrementalTable(this, _name, _headers, ReadRows(), _cancellationToken);
            _completed = true;
            return result;
        } catch {
            _failed = true;
            _owner.AbortIncrementalTable(this);
            throw;
        } finally {
            Dispose();
        }
    }

    private IEnumerable<object?[]> ReadRows() {
        using var reader = new BinaryReader(_spool, System.Text.Encoding.UTF8, leaveOpen: true);
        var values = new object?[_headers.Count];
        for (int row = 0; row < _rowCount; row++) {
            _cancellationToken.ThrowIfCancellationRequested();
            Array.Clear(values, 0, values.Length);
            int count = reader.ReadInt32();
            for (int index = 0; index < count; index++) values[reader.ReadInt32()] = ReadValue(reader);
            yield return values;
        }
    }

    private void WriteValue(object? value) {
        if (value == null || value == DBNull.Value) { _writer.Write((byte)TypeCode.Empty); return; }
        if (value is DateTimeOffset offset) { _writer.Write((byte)32); _writer.Write(offset.Ticks); _writer.Write(offset.Offset.Ticks); return; }
        if (value is TimeSpan duration) { _writer.Write((byte)33); _writer.Write(duration.Ticks); return; }
#if NET6_0_OR_GREATER
        if (value is DateOnly date) { _writer.Write((byte)34); _writer.Write(date.DayNumber); return; }
        if (value is TimeOnly time) { _writer.Write((byte)35); _writer.Write(time.Ticks); return; }
#endif
        TypeCode code = value.GetType().IsEnum ? TypeCode.String : Type.GetTypeCode(value.GetType());
        if (code is TypeCode.Object or TypeCode.Char) code = TypeCode.String;
        _writer.Write((byte)code);
        switch (code) {
            case TypeCode.Boolean: _writer.Write((bool)value); break;
            case TypeCode.SByte: _writer.Write((sbyte)value); break;
            case TypeCode.Byte: _writer.Write((byte)value); break;
            case TypeCode.Int16: _writer.Write((short)value); break;
            case TypeCode.UInt16: _writer.Write((ushort)value); break;
            case TypeCode.Int32: _writer.Write((int)value); break;
            case TypeCode.UInt32: _writer.Write((uint)value); break;
            case TypeCode.Int64: _writer.Write((long)value); break;
            case TypeCode.UInt64: _writer.Write((ulong)value); break;
            case TypeCode.Single: _writer.Write((float)value); break;
            case TypeCode.Double: _writer.Write((double)value); break;
            case TypeCode.Decimal: _writer.Write((decimal)value); break;
            case TypeCode.DateTime: _writer.Write(((DateTime)value).ToBinary()); break;
            default: _writer.Write(value.ToString() ?? string.Empty); break;
        }
    }

    private static object? ReadValue(BinaryReader reader) => reader.ReadByte() switch {
        (byte)TypeCode.Empty => null,
        (byte)TypeCode.Boolean => reader.ReadBoolean(),
        (byte)TypeCode.SByte => reader.ReadSByte(),
        (byte)TypeCode.Byte => reader.ReadByte(),
        (byte)TypeCode.Int16 => reader.ReadInt16(),
        (byte)TypeCode.UInt16 => reader.ReadUInt16(),
        (byte)TypeCode.Int32 => reader.ReadInt32(),
        (byte)TypeCode.UInt32 => reader.ReadUInt32(),
        (byte)TypeCode.Int64 => reader.ReadInt64(),
        (byte)TypeCode.UInt64 => reader.ReadUInt64(),
        (byte)TypeCode.Single => reader.ReadSingle(),
        (byte)TypeCode.Double => reader.ReadDouble(),
        (byte)TypeCode.Decimal => reader.ReadDecimal(),
        (byte)TypeCode.DateTime => DateTime.FromBinary(reader.ReadInt64()),
        32 => new DateTimeOffset(reader.ReadInt64(), TimeSpan.FromTicks(reader.ReadInt64())),
        33 => TimeSpan.FromTicks(reader.ReadInt64()),
#if NET6_0_OR_GREATER
        34 => DateOnly.FromDayNumber(reader.ReadInt32()),
        35 => new TimeOnly(reader.ReadInt64()),
#endif
        _ => reader.ReadString()
    };

    private void EnsureActive() {
        if (_disposed) throw new ObjectDisposedException(nameof(ExcelIncrementalTableWriter));
        if (_failed) throw new InvalidOperationException("A failed incremental table cannot accept more rows.");
        _cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>Releases the native spool. An incomplete table marks the workbook failed, preserving an existing destination.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        if (!_completed) _owner.AbortIncrementalTable(this);
        try { _writer.Dispose(); }
        finally { _spool.Dispose(); }
    }
}
