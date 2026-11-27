using OfficeIMO.Fluent;

namespace OfficeIMO.Fluent.Generators.Tests.Sample;

/// <summary>A worksheet-like store that owns the cell state.</summary>
public sealed class SampleSheet {
    private readonly Dictionary<(int Row, int Column), (object? Value, bool Bold)> _cells = new();

    public SampleCell Cell(int row, int column) => new SampleCell(this, row, column);

    internal (object? Value, bool Bold) Read(int row, int column) => _cells.TryGetValue((row, column), out var cell) ? cell : default;

    internal void Write(int row, int column, object? value) => _cells[(row, column)] = (value, Read(row, column).Bold);

    internal void WriteBold(int row, int column, bool bold) => _cells[(row, column)] = (Read(row, column).Value, bold);
}

/// <summary>A value-type handle: its identity is the sheet and position, so fluent calls write through to the sheet.</summary>
[Fluent]
public readonly struct SampleCell : IEquatable<SampleCell> {
    private readonly SampleSheet _sheet;

    internal SampleCell(SampleSheet sheet, int row, int column) {
        _sheet = sheet;
        Row = row;
        Column = column;
    }

    public int Row { get; }
    public int Column { get; }
    public object? CurrentValue => _sheet.Read(Row, Column).Value;
    public bool IsBold => _sheet.Read(Row, Column).Bold;

    [FluentMutator("Value")]
    internal void SetValueCore(object? value) => _sheet.Write(Row, Column, value);

    [FluentMutator("Bold")]
    internal void SetBoldCore(bool bold) => _sheet.WriteBold(Row, Column, bold);

    public bool Equals(SampleCell other) => ReferenceEquals(_sheet, other._sheet) && Row == other.Row && Column == other.Column;
    public override bool Equals(object? obj) => obj is SampleCell other && Equals(other);
    public override int GetHashCode() => HashCode.Combine(_sheet, Row, Column);
}
