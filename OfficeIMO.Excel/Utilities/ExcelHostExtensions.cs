using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Excel;

/// <summary>
/// Helper extensions for common Excel document tasks (sheet resolution, A1 parsing).
/// </summary>
public static class ExcelHostExtensions
{
    /// <summary>
    /// Returns the named sheet if it exists, otherwise creates it; falls back to the last sheet.
    /// </summary>
    public static ExcelSheet GetOrCreateSheet(this ExcelDocument document, string? name, SheetNameValidationMode validationMode)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var sheetsCollection = document.Sheets ?? new List<ExcelSheet>();

        if (!string.IsNullOrWhiteSpace(name))
        {
            var existing = sheetsCollection.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing;
            }

            return document.AddWorkSheet(name ?? string.Empty, validationMode);
        }

        if (sheetsCollection.Count == 0)
        {
            return document.AddWorkSheet(string.Empty, SheetNameValidationMode.None);
        }

        return sheetsCollection[sheetsCollection.Count - 1];
    }

    /// <summary>
    /// Resolves a cell address from either A1 notation or row/column coordinates.
    /// </summary>
    public static (int Row, int Column) ResolveCellAddress(int? row, int? column, string? address)
    {
        if (!string.IsNullOrWhiteSpace(address))
        {
            var trimmedAddress = address!.Trim();
            var (rowIndex, columnIndex) = A1.ParseCellRef(trimmedAddress);
            if (rowIndex <= 0 || columnIndex <= 0)
            {
                throw new ArgumentException($"Address '{address}' is not a valid A1 reference.", nameof(address));
            }

            return (rowIndex, columnIndex);
        }

        if (!row.HasValue || !column.HasValue)
        {
            throw new ArgumentException("Specify either -Address or both -Row and -Column.");
        }

        return (row.Value, column.Value);
    }

    /// <summary>
    /// Resolves a column index from either a numeric index or column letters.
    /// </summary>
    public static int ResolveColumnIndex(int? columnIndex, string? columnName)
    {
        if (!string.IsNullOrWhiteSpace(columnName))
        {
            var index = A1.ColumnLettersToIndex(columnName!.Trim());
            if (index <= 0)
            {
                throw new ArgumentException($"ColumnName '{columnName}' is not a valid column reference.", nameof(columnName));
            }
            return index;
        }

        if (!columnIndex.HasValue)
        {
            throw new ArgumentException("Specify either -Column or -ColumnName.");
        }

        return columnIndex.Value;
    }
}
