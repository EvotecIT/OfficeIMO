using System;
using System.Collections.Generic;
using System.Globalization;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

/// <summary>
/// PowerShell-facing editable cell handle backed by the public OfficeIMO Excel cell model.
/// </summary>
public sealed class ExcelPowerShellCellEdit
{
    private readonly ExcelCell _cell;
    private readonly CultureInfo _culture;
    private object? _value;

    internal ExcelPowerShellCellEdit(ExcelCell cell, object? value, CultureInfo culture)
    {
        _cell = cell ?? throw new ArgumentNullException(nameof(cell));
        _value = value is DBNull ? null : value;
        _culture = culture ?? CultureInfo.InvariantCulture;
    }

    /// <summary>Gets the 1-based worksheet row index.</summary>
    public int RowIndex => _cell.Row;

    /// <summary>Gets the 1-based worksheet column index.</summary>
    public int ColumnIndex => _cell.Column;

    /// <summary>Gets the A1 cell address.</summary>
    public string Address => _cell.Address;

    /// <summary>Gets or sets the current cell value.</summary>
    public object? Value
    {
        get => _value;
        set
        {
            _cell.SetValue(value);
            _value = value;
        }
    }

    /// <summary>Converts the current value using PowerShell conversion rules.</summary>
    public T ConvertTo<T>() => _value is null
        ? default!
        : (T)LanguagePrimitives.ConvertTo(_value, typeof(T), _culture);

    /// <summary>Applies an Excel number format to the cell.</summary>
    public void NumberFormat(string format) => _cell.SetNumberFormat(format);

    /// <summary>Sets a formula on the cell.</summary>
    public void Formula(string formula) => _cell.SetFormula(formula);
}

/// <summary>
/// Header-aware editable worksheet row used by <c>Edit-OfficeExcelRow</c>.
/// </summary>
public sealed class ExcelPowerShellRowEdit
{
    private readonly IReadOnlyList<ExcelPowerShellCellEdit> _cells;
    private readonly Dictionary<string, int> _headerMap;

    internal ExcelPowerShellRowEdit(
        ExcelSheet sheet,
        int rowIndex,
        int firstColumn,
        IReadOnlyList<string> headers,
        IReadOnlyList<object?> values,
        CultureInfo culture)
    {
        RowIndex = rowIndex;
        _headerMap = new Dictionary<string, int>(headers.Count, StringComparer.OrdinalIgnoreCase);
        var cells = new ExcelPowerShellCellEdit[headers.Count];
        for (var index = 0; index < headers.Count; index++)
        {
            var header = headers[index];
            if (_headerMap.ContainsKey(header))
            {
                throw new InvalidOperationException($"Header '{header}' is duplicated in the editable range.");
            }

            _headerMap.Add(header, index);

            var value = index < values.Count ? values[index] : null;
            cells[index] = new ExcelPowerShellCellEdit(sheet.CellAt(rowIndex, firstColumn + index), value, culture);
        }

        _cells = cells;
    }

    /// <summary>Gets the 1-based worksheet row index.</summary>
    public int RowIndex { get; }

    /// <summary>Gets editable cells in range-column order.</summary>
    public IReadOnlyList<ExcelPowerShellCellEdit> Cells => _cells;

    /// <summary>Gets a cell by its 1-based position within the editable range.</summary>
    public ExcelPowerShellCellEdit this[int columnIndex] => columnIndex > 0 && columnIndex <= _cells.Count
        ? _cells[columnIndex - 1]
        : throw new ArgumentOutOfRangeException(nameof(columnIndex));

    /// <summary>Gets a cell by header name.</summary>
    public ExcelPowerShellCellEdit this[string header] => _headerMap.TryGetValue(header, out var index)
        ? _cells[index]
        : throw new KeyNotFoundException($"Header '{header}' was not found.");

    /// <summary>Gets a typed value by header name.</summary>
    public T Get<T>(string header) => this[header].ConvertTo<T>();

    /// <summary>Gets a typed value by header name or returns a default when conversion fails.</summary>
    public T GetOrDefault<T>(string header, T @default = default!)
    {
        try
        {
            return this[header].Value is null ? @default : this[header].ConvertTo<T>();
        }
        catch (Exception exception) when (exception is InvalidCastException or FormatException or OverflowException or PSInvalidCastException)
        {
            return @default;
        }
    }

    /// <summary>Sets a cell value by header name.</summary>
    public void Set(string header, object? value) => this[header].Value = value;

    /// <summary>Gets an editable cell by header name.</summary>
    public ExcelPowerShellCellEdit CellByHeader(string header) => this[header];

    /// <summary>Applies a number format to a cell selected by header name.</summary>
    public void NumberFormat(string header, string format) => this[header].NumberFormat(format);

    /// <summary>Sets a formula on a cell selected by header name.</summary>
    public void SetFormula(string header, string formula) => this[header].Formula(formula);
}
