#nullable enable

namespace OfficeIMO.CSV;

public sealed partial class CsvObjectWriter
{
    /// <summary>
    /// Writes already-projected rows using one shared column order.
    /// </summary>
    /// <param name="columns">Column names shared by every row.</param>
    /// <param name="rows">Rows whose values use the same order as <paramref name="columns"/>.</param>
    /// <remarks>
    /// The method validates the schema once and validates every row width while streaming the rows.
    /// Use <see cref="WriteTextRows"/> when the caller already owns culture-aware value formatting.
    /// </remarks>
    public void WriteRows(IReadOnlyList<string> columns, IEnumerable<object?[]?> rows)
    {
        ThrowIfDisposed();
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        if (rows == null)
        {
            throw new ArgumentNullException(nameof(rows));
        }

        EnsureColumns(columns);
        foreach (var row in rows)
        {
            if (row == null)
            {
                throw new InvalidOperationException("Data rows cannot contain null entries.");
            }

            ValidateProjectedValueCount(columns, row.Length);
            WriteBuffered(row);
        }
    }

    /// <summary>
    /// Writes already-formatted text rows using one shared column order.
    /// </summary>
    /// <param name="columns">Column names shared by every row.</param>
    /// <param name="rows">Text rows whose values use the same order as <paramref name="columns"/>.</param>
    /// <remarks>
    /// The method validates the schema once, validates every row width, and applies CSV escaping to every value.
    /// Use this when the caller already owns culture-aware value formatting.
    /// </remarks>
    public void WriteTextRows(IReadOnlyList<string> columns, IEnumerable<string?[]?> rows)
    {
        ThrowIfDisposed();
        if (columns == null)
        {
            throw new ArgumentNullException(nameof(columns));
        }

        if (rows == null)
        {
            throw new ArgumentNullException(nameof(rows));
        }

        EnsureColumns(columns);
        foreach (var row in rows)
        {
            if (row == null)
            {
                throw new InvalidOperationException("Data rows cannot contain null entries.");
            }

            ValidateProjectedValueCount(columns, row.Length);
            WriteTextBuffered(row);
        }
    }
}
