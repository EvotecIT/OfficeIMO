using System.Globalization;
using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

/// <summary>Controls Excel workbook projection into Reader chunks.</summary>
public sealed class ReaderExcelOptions {
    /// <summary>Optional worksheet name. All worksheets are read when omitted.</summary>
    public string? SheetName { get; set; }

    /// <summary>Optional one-based worksheet position. Cannot be combined with SheetName.</summary>
    public int? SheetNumber { get; set; }

    /// <summary>Optional A1 range applied to each selected worksheet.</summary>
    public string? A1Range { get; set; }

    /// <summary>Treats the first range row as column names.</summary>
    public bool HeadersInFirstRow { get; set; } = true;

    /// <summary>Number of source rows requested per extraction chunk.</summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>Optional low-level Excel read policy.</summary>
    public ExcelReaderOptions? ReadOptions { get; set; }

    /// <summary>
    /// Chooses whether cell text is the value as each cell's number format displays it, such as <c>3.20%</c>,
    /// <c>$31,200.00</c> or <c>2024-01-15</c>, or the stored value. Defaults to <see cref="ReaderExcelCellText.Formatted"/>.
    /// </summary>
    public ReaderExcelCellText CellText { get; set; } = ReaderExcelCellText.Formatted;

    /// <summary>
    /// Culture used to display formatted cell text. Defaults to <see cref="CultureInfo.InvariantCulture"/>, so the same
    /// workbook produces the same text on every machine; workbooks do not record the locale they were edited in.
    /// Literal currency symbols and locale tags inside a number format are kept as written.
    /// </summary>
    public CultureInfo CellTextCulture { get; set; } = CultureInfo.InvariantCulture;
}

/// <summary>Selects how Excel cell values become extracted text.</summary>
public enum ReaderExcelCellText {
    /// <summary>
    /// Text as the cell's number format displays it: percentages, currency, dates, times, fractions and scientific
    /// notation follow the format, General numbers use up to 15 significant digits, and booleans read TRUE or FALSE.
    /// </summary>
    Formatted = 0,

    /// <summary>
    /// Stored values: numbers as invariant text, date-formatted cells as ISO 8601 timestamps and
    /// booleans as True or False.
    /// </summary>
    Raw = 1
}
