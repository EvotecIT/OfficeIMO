namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Formula values returned by <c>Import-OfficeExcel</c>.</summary>
public enum OfficeExcelFormulaMode
{
    /// <summary>Return the cached result stored in the workbook.</summary>
    CachedValue,
    /// <summary>Return formula expressions when present.</summary>
    FormulaText
}

/// <summary>Column-format presets exposed by <c>Set-OfficeExcelColumnStyleByHeader</c>.</summary>
public enum OfficeExcelColumnStylePreset
{
    /// <summary>General decimal number format.</summary>
    Number,
    /// <summary>Whole-number format.</summary>
    Integer,
    /// <summary>Percentage format.</summary>
    Percent,
    /// <summary>Culture-aware currency format.</summary>
    Currency,
    /// <summary>Date format.</summary>
    Date,
    /// <summary>Date and time format.</summary>
    DateTime,
    /// <summary>Time format.</summary>
    Time,
    /// <summary>Elapsed-hours duration format.</summary>
    DurationHours,
    /// <summary>Text format.</summary>
    Text,
    /// <summary>Caller-supplied Excel number format.</summary>
    NumberFormat
}

/// <summary>Supported column alignment shortcuts.</summary>
public enum OfficeExcelColumnAlignment
{
    /// <summary>Align cell content to the left.</summary>
    Left,
    /// <summary>Center cell content.</summary>
    Center,
    /// <summary>Align cell content to the right.</summary>
    Right
}

/// <summary>Package metadata kinds supported by <c>Add-OfficeExcelPackageMetadata</c>.</summary>
public enum OfficeExcelPackageMetadataKind
{
    /// <summary>Workbook connection metadata.</summary>
    Connection,
    /// <summary>Worksheet query-table metadata.</summary>
    QueryTable
}

/// <summary>Report callout styles supported by the Excel report composer.</summary>
public enum OfficeExcelReportCalloutKind
{
    /// <summary>Informational callout.</summary>
    Info,
    /// <summary>Successful outcome callout.</summary>
    Success,
    /// <summary>Warning callout.</summary>
    Warning,
    /// <summary>Error callout.</summary>
    Error,
    /// <summary>Critical issue callout.</summary>
    Critical
}
