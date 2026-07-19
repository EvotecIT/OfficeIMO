using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

/// <summary>Controls Excel workbook projection into Reader chunks.</summary>
public sealed class ReaderExcelOptions {
    /// <summary>Optional worksheet name. All worksheets are read when omitted.</summary>
    public string? SheetName { get; set; }

    /// <summary>Optional A1 range applied to each selected worksheet.</summary>
    public string? A1Range { get; set; }

    /// <summary>Treats the first range row as column names.</summary>
    public bool HeadersInFirstRow { get; set; } = true;

    /// <summary>Number of source rows requested per extraction chunk.</summary>
    public int ChunkRows { get; set; } = 200;

    /// <summary>Optional low-level Excel read policy.</summary>
    public ExcelReadOptions? ReadOptions { get; set; }
}
