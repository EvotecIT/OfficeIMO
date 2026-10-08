namespace OfficeIMO.Access;

/// <summary>One internal cursor boundary for modeled and native rows; the public reader owns the document lease.</summary>
internal interface IAccessRowCursor : IDisposable {
    bool Read(CancellationToken cancellation);
    bool HasRows { get; }
    bool IsSpecified(int ordinal);
    bool IsNull(int ordinal);
    Stream OpenBinary(int ordinal, CancellationToken cancellation);
    object? GetValue(int ordinal, CancellationToken cancellation);
}

internal sealed class AccessModeledRowCursor : IAccessRowCursor {
    private readonly AccessTable _table;
    private int _position = -1;
    internal AccessModeledRowCursor(AccessTable table) { _table = table; }
    public bool HasRows => _table.Rows.Count != 0;
    public bool Read(CancellationToken cancellation) { cancellation.ThrowIfCancellationRequested(); if (_position < _table.Rows.Count) _position++; return _position < _table.Rows.Count; }
    private Dictionary<string, object?> Current => _position >= 0 && _position < _table.Rows.Count ? _table.Rows[_position] : throw new InvalidOperationException("Call Read before accessing a current row.");
    public bool IsSpecified(int ordinal) => Current.ContainsKey(_table.Columns[ordinal].Name);
    public bool IsNull(int ordinal) => !Current.TryGetValue(_table.Columns[ordinal].Name, out var value) || value == null;
    public Stream OpenBinary(int ordinal, CancellationToken cancellation) { cancellation.ThrowIfCancellationRequested(); return new MemoryStream((byte[])AccessTable.CopyValue(GetValue(ordinal, cancellation))!, writable: false); }
    public object? GetValue(int ordinal, CancellationToken cancellation) { cancellation.ThrowIfCancellationRequested(); return Current.TryGetValue(_table.Columns[ordinal].Name, out object? value) ? value : null; }
    public void Dispose() { }
}
