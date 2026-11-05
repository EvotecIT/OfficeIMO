using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.Csv;

/// <summary>Controls worksheet projection into CSV; serialization remains owned by the CSV engine.</summary>
public sealed class ExcelCsvExportOptions {
    /// <summary>A1 range to project; null uses the worksheet's used range.</summary>
    public string? Range { get; set; }
    /// <summary>Whether the first selected row supplies field names.</summary>
    public bool HeadersInFirstRow { get; set; } = true;
    /// <summary>Cell-value projection, culture and cancellation options.</summary>
    public ExcelReaderOptions? ReadOptions { get; set; }
    /// <summary>Encoding, compression, formatting and file conflict policy for streamed saves.</summary>
    public CsvSaveOptions? SaveOptions { get; set; }
    /// <summary>Rejects unassessed workbook fidelity before output is written.</summary>
    public bool FailOnLoss { get; set; }
}

/// <summary>Controls detached CSV-to-Excel conversion and optional persistence.</summary>
public sealed class CsvExcelExportOptions {
    /// <summary>CSV schema inference, headers, worksheet placement and table behavior.</summary>
    public ExcelCsvImportOptions ImportOptions { get; set; } = new ExcelCsvImportOptions();
    /// <summary>Options used to create the detached workbook.</summary>
    public ExcelCreateOptions? CreateOptions { get; set; }
    /// <summary>Native workbook persistence options.</summary>
    public ExcelSaveOptions? SaveOptions { get; set; }
    /// <summary>Rejects conversion when its completed report describes fidelity loss.</summary>
    public bool FailOnLoss { get; set; }
}
