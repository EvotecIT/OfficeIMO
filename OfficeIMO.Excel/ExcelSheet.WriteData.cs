using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Writes typed cells from a DataTable without creating a native table or applying table layout defaults.</summary>
    public void WriteData(DataTable data, int row = 1, int column = 1, bool includeHeaders = true,
        ExcelExecutionMode? mode = null, CancellationToken cancellationToken = default) =>
        InsertDataTable(data, row, column, includeHeaders, mode, cancellationToken);

    /// <summary>Writes flattened object graphs as plain cells. Null and DBNull remain blank; empty strings remain text.</summary>
    [RequiresUnreferencedCode("Uses reflection over arbitrary object graphs. For AOT-safe usage, use explicit WriteData column selectors.")]
    public void WriteData<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        IEnumerable<T> data, bool includeHeaders = true, int row = 1) =>
        InsertObjects(data, includeHeaders, row);

    /// <summary>Writes plain cells through explicit selectors without runtime reflection or table creation.</summary>
    public void WriteData<T>(IEnumerable<T> data, params (string Header, Func<T, object?> Selector)[] columns) =>
        InsertObjects(data, columns);

    /// <summary>
    /// Writes a header row and plain cells from row one through typed column selectors, such as
    /// <c>WriteData(rows, columns =&gt; columns.Add("Name", row =&gt; row.Name).Add("Id", row =&gt; row.Id))</c>.
    /// The cells match the equivalent <c>(header, selector)</c> columns, and numbers, dates and Boolean values are
    /// kept without boxing. Selectors run once per cell, after <paramref name="data"/> has been enumerated.
    /// </summary>
    public void WriteData<T>(IEnumerable<T> data, Action<ExcelDataColumns<T>> columns) {
        ArgumentNullException.ThrowIfNull(columns);
        var definitions = new ExcelDataColumns<T>();
        columns(definitions);
        InsertTypedObjects(data, definitions, includeHeaders: true, startRow: 1);
    }

    /// <summary>Writes plain cells at a chosen row through explicit selectors, preserving duplicate or blank display headers.</summary>
    public void WriteData<T>(IEnumerable<T> data, bool includeHeaders, int row,
        params (string Header, Func<T, object?> Selector)[] columns) =>
        InsertObjects(data, includeHeaders, row, columns);
}
