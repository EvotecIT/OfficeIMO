using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Reader;

/// <summary>Options for deterministic table cleanup.</summary>
public sealed class OfficeDocumentTableNormalizationOptions {
    /// <summary>Trim column names and cell values.</summary>
    public bool TrimValues { get; set; } = true;

    /// <summary>Generate deterministic names for blank columns.</summary>
    public bool FillMissingColumnNames { get; set; } = true;

    /// <summary>Pad rows so every row has the same width.</summary>
    public bool RectangularRows { get; set; } = true;

    internal OfficeDocumentTableNormalizationOptions Clone() => new OfficeDocumentTableNormalizationOptions {
        TrimValues = TrimValues,
        FillMissingColumnNames = FillMissingColumnNames,
        RectangularRows = RectangularRows
    };
}

/// <summary>Normalizes table widths, column labels, cells, and column profiles.</summary>
public sealed class OfficeDocumentTableNormalizationProcessor : OfficeDocumentProcessorBase {
    private readonly OfficeDocumentTableNormalizationOptions _options;

    /// <summary>Creates the processor.</summary>
    public OfficeDocumentTableNormalizationProcessor(OfficeDocumentTableNormalizationOptions? options = null)
        : base("officeimo.reader.normalize-tables") {
        _options = (options ?? new OfficeDocumentTableNormalizationOptions()).Clone();
    }

    /// <inheritdoc />
    public override OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        foreach (ReaderTable table in OfficeDocumentModelTraversal.TableInstances(document)) {
            context.CancellationToken.ThrowIfCancellationRequested();
            Normalize(table);
        }
        return document;
    }

    private void Normalize(ReaderTable table) {
        IReadOnlyList<string> sourceColumns = table.Columns ?? Array.Empty<string>();
        IReadOnlyList<IReadOnlyList<string>> sourceRows = table.Rows ?? Array.Empty<IReadOnlyList<string>>();
        int width = sourceColumns.Count;
        for (int rowIndex = 0; rowIndex < sourceRows.Count; rowIndex++) {
            width = Math.Max(width, sourceRows[rowIndex]?.Count ?? 0);
        }

        var columns = new string[width];
        for (int columnIndex = 0; columnIndex < width; columnIndex++) {
            string value = columnIndex < sourceColumns.Count ? sourceColumns[columnIndex] ?? string.Empty : string.Empty;
            if (_options.TrimValues) value = value.Trim();
            if (_options.FillMissingColumnNames && value.Length == 0) {
                value = "Column" + (columnIndex + 1).ToString(CultureInfo.InvariantCulture);
            }
            columns[columnIndex] = value;
        }

        var rows = new IReadOnlyList<string>[sourceRows.Count];
        for (int rowIndex = 0; rowIndex < sourceRows.Count; rowIndex++) {
            IReadOnlyList<string>? sourceRow = sourceRows[rowIndex];
            int rowWidth = _options.RectangularRows ? width : sourceRow?.Count ?? 0;
            var row = new string[rowWidth];
            for (int columnIndex = 0; columnIndex < rowWidth; columnIndex++) {
                string value = sourceRow != null && columnIndex < sourceRow.Count
                    ? sourceRow[columnIndex] ?? string.Empty
                    : string.Empty;
                row[columnIndex] = _options.TrimValues ? value.Trim() : value;
            }
            rows[rowIndex] = row;
        }

        table.Columns = columns;
        table.Rows = rows;
        table.TotalRowCount = Math.Max(table.TotalRowCount, rows.Length);
        table.ColumnProfiles = ReaderTableProfiler.CreateProfiles(columns, rows);
        if (_options.TrimValues && table.Title != null) table.Title = table.Title.Trim();
        if (_options.TrimValues && table.Kind != null) table.Kind = table.Kind.Trim();
    }
}
