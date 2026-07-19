using System.Globalization;
using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

/// <summary>Controls culture-aware CSV and TSV import into an Excel workbook.</summary>
public sealed class ExcelDelimitedImportOptions {
    /// <summary>Gets or sets the delimiter. When omitted, the delimiter is detected.</summary>
    public char? Delimiter { get; set; }

    /// <summary>Gets or sets whether the first imported record contains column names.</summary>
    public bool HeadersInFirstRow { get; set; } = true;

    /// <summary>Gets or sets the number of logical records skipped before import.</summary>
    public int SkipInitialRecords { get; set; }

    /// <summary>Gets or sets the culture used for number and date conversion.</summary>
    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    /// <summary>Gets or sets whether number and date text is converted to typed values.</summary>
    public bool ConvertNumbersAndDates { get; set; } = true;

    /// <summary>Gets or sets whether an Excel table is created over the imported range.</summary>
    public bool CreateTable { get; set; } = true;

    /// <summary>Gets or sets the worksheet name.</summary>
    public string? SheetName { get; set; }

    /// <summary>Gets or sets the Excel table name.</summary>
    public string? TableName { get; set; }

    /// <summary>Gets or sets the Excel table style.</summary>
    public TableStyle TableStyle { get; set; } = TableStyle.TableStyleMedium2;
}
