namespace OfficeIMO.Excel;

/// <summary>Receives the current typed cell without retaining source rows or boxing primitive values.</summary>
internal abstract class ExcelTabularRowSink : IDisposable {
    protected int Column { get; private set; }
    protected uint Style { get; private set; }
    internal void BeginCell(int column, uint style) { Column = column; Style = style; }
    internal abstract void BeginRow(int oneBasedRow);
    internal abstract void EndRow();
    internal abstract void WriteText(string value);
    internal abstract void WriteNumber(double value);
    internal abstract void WriteBoolean(bool value);
    internal abstract void WriteBlank();
    internal abstract void Complete();
    public abstract void Dispose();

    protected static void ValidateNumber(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value))
            throw new ArgumentException("Binary Excel cells require a finite number.", nameof(value));
    }
}
