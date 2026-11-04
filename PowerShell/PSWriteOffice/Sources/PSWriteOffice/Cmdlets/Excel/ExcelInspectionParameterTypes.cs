using OfficeIMO.Excel;
namespace PSWriteOffice.Cmdlets.Excel;
/// <summary>Native repair policy areas.</summary>
public enum ExcelRepairArea { DefinedNames, Tables, SheetViews, PrintSettings, Protection, AutoFilters, Drawings, LinksAndComments, Calculation, WorkbookArtifacts }
/// <summary>Native workbook comparison areas.</summary>
public enum ExcelCompareArea { Cells, CellStyles, NamedRanges, Tables, WorksheetMetadata, Comments }
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
