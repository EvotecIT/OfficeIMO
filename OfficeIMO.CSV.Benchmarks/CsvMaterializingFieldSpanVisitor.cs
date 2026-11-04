namespace OfficeIMO.CSV.Benchmarks;

internal struct CsvMaterializingFieldSpanVisitor : ICsvFieldSpanVisitor
{
    private readonly List<string> _currentRow;
    private int _currentRecordIndex;
    private readonly string[][]? _expectedRows;

    public CsvMaterializingFieldSpanVisitor() : this(null) { }

    public CsvMaterializingFieldSpanVisitor(string[][]? expectedRows)
    {
        _expectedRows = expectedRows;
        _currentRow = new List<string>(64);
        _currentRecordIndex = -1;
        FieldCount = 0;
        RowCount = 0;
        TextLength = 0;
    }

    public int FieldCount { get; private set; }

    public int RowCount { get; private set; }

    public int TextLength { get; private set; }

    public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
    {
        if (recordIndex != _currentRecordIndex)
        {
            if (_currentRecordIndex >= 0)
            {
                if (_expectedRows is not null && _currentRow.Count != _expectedRows[RowCount].Length)
                    throw new InvalidDataException("CSV field-span traversal changed a row width.");
                RowCount++;
            }

            _currentRecordIndex = recordIndex;
            _currentRow.Clear();
        }

        ValidateField(fieldIndex, value);
        var text = value.ToString();
        _currentRow.Add(text);
        TextLength += text.Length;
        FieldCount++;
    }

    public void VisitFieldValue(int recordIndex, int fieldIndex, string value)
    {
        if (recordIndex != _currentRecordIndex)
        {
            if (_currentRecordIndex >= 0)
            {
                if (_expectedRows is not null && _currentRow.Count != _expectedRows[RowCount].Length)
                    throw new InvalidDataException("CSV field-span traversal changed a row width.");
                RowCount++;
            }

            _currentRecordIndex = recordIndex;
            _currentRow.Clear();
        }

        ValidateField(fieldIndex, value);
        _currentRow.Add(value);
        TextLength += value.Length;
        FieldCount++;
    }

    public void Complete()
    {
        if (_currentRecordIndex >= 0)
        {
            RowCount++;
        }
        if (_expectedRows is not null && (RowCount != _expectedRows.Length || (_expectedRows.Length > 0 && _currentRow.Count != _expectedRows[^1].Length)))
            throw new InvalidDataException("CSV field-span traversal changed the complete row count or width.");
    }

    private void ValidateField(int fieldIndex, ReadOnlySpan<char> value)
    {
        if (_expectedRows is not null && ((uint)RowCount >= (uint)_expectedRows.Length
            || fieldIndex != _currentRow.Count || (uint)fieldIndex >= (uint)_expectedRows[RowCount].Length
            || !value.SequenceEqual(_expectedRows[RowCount][fieldIndex].AsSpan())))
            throw new InvalidDataException("CSV field-span traversal changed a raw field.");
    }
}
