namespace OfficeIMO.Access {
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
        public bool IsNull(int ordinal) => !Current.TryGetValue(_table.Columns[ordinal].Name, out object? value) || value == null;
        public Stream OpenBinary(int ordinal, CancellationToken cancellation) {
            cancellation.ThrowIfCancellationRequested();
            object? value = GetValue(ordinal, cancellation);
            if (value == null) throw new InvalidCastException("The current field is null.");
            if (value is not byte[] binary) throw new InvalidCastException("The current field is not binary.");
            return new OfficeIMO.Core.Internal.OfficeDocumentReadStream(new MemoryStream((byte[])AccessTable.CopyValue(binary)!, writable: false), _table.Document.EnsureNotDisposed, cancellation);
        }
        public object? GetValue(int ordinal, CancellationToken cancellation) { cancellation.ThrowIfCancellationRequested(); return Current.TryGetValue(_table.Columns[ordinal].Name, out object? value) ? value : null; }
        public void Dispose() { }
    }
}
