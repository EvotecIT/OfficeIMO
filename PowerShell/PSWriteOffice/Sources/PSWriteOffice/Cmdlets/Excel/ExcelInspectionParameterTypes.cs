using OfficeIMO.Excel;
namespace PSWriteOffice.Cmdlets.Excel;
/// <summary>Native repair policy areas.</summary>
public enum ExcelRepairArea {
    /// <summary>Apply the native defined-name repair policy.</summary>
    DefinedNames,
    /// <summary>Apply the native table repair policy.</summary>
    Tables,
    /// <summary>Apply the native worksheet-view repair policy.</summary>
    SheetViews,
    /// <summary>Apply the native print-settings repair policy.</summary>
    PrintSettings,
    /// <summary>Apply the native protection repair policy.</summary>
    Protection,
    /// <summary>Apply the native auto-filter repair policy.</summary>
    AutoFilters,
    /// <summary>Apply the native drawing repair policy.</summary>
    Drawings,
    /// <summary>Apply the native link and comment repair policy.</summary>
    LinksAndComments,
    /// <summary>Apply the native calculation repair policy.</summary>
    Calculation,
    /// <summary>Apply the native workbook-artifact repair policy.</summary>
    WorkbookArtifacts,
}
/// <summary>Native workbook comparison areas.</summary>
public enum ExcelCompareArea {
    /// <summary>Compare cell values and formulas.</summary>
    Cells,
    /// <summary>Compare cell formatting.</summary>
    CellStyles,
    /// <summary>Compare workbook and worksheet names.</summary>
    NamedRanges,
    /// <summary>Compare table definitions.</summary>
    Tables,
    /// <summary>Compare worksheet metadata.</summary>
    WorksheetMetadata,
    /// <summary>Compare cell comments.</summary>
    Comments,
}
/// <summary>Excel-only diagnostic policy forwarded by Test-OfficeDocument.</summary>
public sealed class ExcelTestOptions {
    /// <summary>Native workbook diagnostic options; this factory never enables repair.</summary>
    public ExcelWorkbookDoctorOptions Doctor { get; set; } = new ExcelWorkbookDoctorOptions();
    /// <summary>Capabilities to qualify from the feature report.</summary>
    public ExcelPreflightCapability[]? Capability { get; set; }
    /// <summary>Include feature details in the returned report.</summary>
    public bool IncludeFeatures { get; set; }
    /// <summary>Include capability repair hints.</summary>
    public bool IncludeRepairHints { get; set; }
}
