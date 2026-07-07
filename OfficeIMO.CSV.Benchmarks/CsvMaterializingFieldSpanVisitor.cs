#nullable enable

namespace OfficeIMO.CSV.Benchmarks;

internal struct CsvMaterializingFieldSpanVisitor : ICsvFieldSpanVisitor
{
    private readonly List<string> _currentRow;
    private int _currentRecordIndex;

    public CsvMaterializingFieldSpanVisitor()
    {
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
                RowCount++;
            }

            _currentRecordIndex = recordIndex;
            _currentRow.Clear();
        }

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
                RowCount++;
            }

            _currentRecordIndex = recordIndex;
            _currentRow.Clear();
        }

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
    }
}
