using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

/// <summary>
/// Controls how CSV rows are projected into an Excel worksheet. CSV parsing and
/// schema behavior remain owned by <see cref="CsvLoadOptions"/> and
/// <see cref="CsvDataReaderOptions"/>.
/// </summary>
public sealed class ExcelCsvImportOptions {
    /// <summary>Gets or sets the CSV parsing options.</summary>
    public CsvLoadOptions LoadOptions { get; set; } = new CsvLoadOptions { DetectDelimiter = true };

    /// <summary>Gets or sets the CSV data-reader projection options.</summary>
    public CsvDataReaderOptions ReaderOptions { get; set; } = new CsvDataReaderOptions { InferSchema = true };

    /// <summary>Gets or sets the target worksheet name when importing into a workbook.</summary>
    public string SheetName { get; set; } = "Import";

    /// <summary>Gets or sets the 1-based target row.</summary>
    public int StartRow { get; set; } = 1;

    /// <summary>Gets or sets the 1-based target column.</summary>
    public int StartColumn { get; set; } = 1;

    /// <summary>Gets or sets whether field names are written above the data rows.</summary>
    public bool IncludeHeaders { get; set; } = true;

    /// <summary>Gets or sets whether an Excel table is created over the imported range.</summary>
    public bool CreateTable { get; set; } = true;

    /// <summary>Gets or sets the requested Excel table name.</summary>
    public string? TableName { get; set; }

    /// <summary>Gets or sets the Excel table style.</summary>
    public TableStyle TableStyle { get; set; } = TableStyle.TableStyleMedium2;

    /// <summary>Gets or sets whether table filter controls are included.</summary>
    public bool IncludeAutoFilter { get; set; } = true;

    /// <summary>Gets or sets whether imported columns are auto-fitted.</summary>
    public bool AutoFit { get; set; }
}
