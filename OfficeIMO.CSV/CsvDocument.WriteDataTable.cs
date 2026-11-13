using System.Data;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Writes a <see cref="DataTable"/> as CSV: its column names as the header, then its rows in order without the
    /// deleted ones, each value from the row's current version. The output is the same as writing the table's
    /// <see cref="DataTable.CreateDataReader()"/> with <see cref="WriteDataReader(TextWriter, IDataReader, CsvSaveOptions?)"/>,
    /// without the data table reader's cost per value.
    /// </summary>
    /// <param name="writer">Destination text writer.</param>
    /// <param name="table">Source table with at least one column.</param>
    /// <param name="options">Optional save settings.</param>
    public static void WriteDataTable(TextWriter writer, DataTable table, CsvSaveOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(table);
        using var reader = new CsvDataTableRowReader(table);
        WriteDataReader(writer, reader, options);
    }

    /// <summary>Writes a <see cref="DataTable"/> as CSV to a stream, as <see cref="WriteDataTable(TextWriter, DataTable, CsvSaveOptions?)"/> does.</summary>
    /// <param name="destination">Writable destination stream.</param>
    /// <param name="table">Source table with at least one column.</param>
    /// <param name="options">CSV serialization options.</param>
    /// <param name="leaveOpen">Whether the destination stream remains open after writing.</param>
    public static void WriteDataTable(Stream destination, DataTable table, CsvSaveOptions? options = null, bool leaveOpen = true)
    {
        ArgumentNullException.ThrowIfNull(table);
        using var reader = new CsvDataTableRowReader(table);
        WriteDataReader(destination, reader, options, leaveOpen);
    }
}
