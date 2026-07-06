#nullable enable

namespace OfficeIMO.CSV.Benchmarks;

internal struct CsvMaterializingRowFieldSpanVisitor : ICsvRowFieldSpanVisitor
{
    public int FieldCount { get; private set; }

    public int RowCount { get; private set; }

    public int TextLength { get; private set; }

    public void BeginRow(IReadOnlyList<string> header, int rowIndex)
    {
    }

    public void VisitField(int rowIndex, int fieldIndex, ReadOnlySpan<char> value)
    {
        var text = value.ToString();
        TextLength += text.Length;
        FieldCount++;
    }

    public void VisitFieldValue(int rowIndex, int fieldIndex, string value)
    {
        TextLength += value.Length;
        FieldCount++;
    }

    public void EndRow(int rowIndex, int fieldCount)
    {
        RowCount++;
    }
}
