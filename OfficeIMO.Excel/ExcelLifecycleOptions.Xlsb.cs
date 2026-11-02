using OfficeIMO.Excel.Xlsb;

namespace OfficeIMO.Excel;

public sealed partial class ExcelLoadOptions {
    /// <summary>For XLSB sources, controls the largest accepted decompressed package part. The default is 128 MiB.</summary>
    public int MaxPartBytes { get => XlsbImportOptions.MaxPartBytes; set => XlsbImportOptions.MaxPartBytes = value; }

    /// <summary>For XLSB sources, controls the aggregate decompressed package budget. The default is 512 MiB.</summary>
    public long MaxPackageBytes { get => XlsbImportOptions.MaxPackageBytes; set => XlsbImportOptions.MaxPackageBytes = value; }

    /// <summary>For XLSB sources, controls the largest accepted BIFF12 record payload. The default is 64 MiB.</summary>
    public int MaxRecordBytes { get => XlsbImportOptions.MaxRecordBytes; set => XlsbImportOptions.MaxRecordBytes = value; }

    /// <summary>For XLSB sources, controls the aggregate BIFF12 record limit across one workbook import. The default is 8,000,000.</summary>
    public int MaxRecordCount { get => XlsbImportOptions.MaxRecordCount; set => XlsbImportOptions.MaxRecordCount = value; }

    /// <summary>For XLSB sources, controls the maximum number of worksheets accepted from one workbook.</summary>
    public int MaxWorksheets { get => XlsbImportOptions.MaxWorksheets; set => XlsbImportOptions.MaxWorksheets = value; }

    /// <summary>For XLSB sources, controls the maximum number of populated cells projected from one workbook.</summary>
    public int MaxCells { get => XlsbImportOptions.MaxCells; set => XlsbImportOptions.MaxCells = value; }

    /// <summary>For XLSB sources, controls the maximum logical rows emitted across one workbook data-reader operation.</summary>
    public int MaxLogicalRows { get => XlsbImportOptions.MaxLogicalRows; set => XlsbImportOptions.MaxLogicalRows = value; }

    /// <summary>For XLSB sources, controls the maximum number of row metadata definitions projected from one workbook.</summary>
    public int MaxRowDefinitions { get => XlsbImportOptions.MaxRowDefinitions; set => XlsbImportOptions.MaxRowDefinitions = value; }

    /// <summary>For XLSB sources, controls the maximum number of shared-string items accepted.</summary>
    public int MaxSharedStrings { get => XlsbImportOptions.MaxSharedStrings; set => XlsbImportOptions.MaxSharedStrings = value; }

    /// <summary>For XLSB sources, controls the maximum UTF-16 character length of one shared-string item.</summary>
    public int MaxSharedStringItemCharacters { get => XlsbImportOptions.MaxSharedStringItemCharacters; set => XlsbImportOptions.MaxSharedStringItemCharacters = value; }

    /// <summary>For XLSB sources, controls the aggregate UTF-16 character budget for the shared-string table.</summary>
    public long MaxSharedStringCharacters { get => XlsbImportOptions.MaxSharedStringCharacters; set => XlsbImportOptions.MaxSharedStringCharacters = value; }

    /// <summary>For XLSB sources, controls the maximum number of merged ranges accepted from one workbook.</summary>
    public int MaxMergedRanges { get => XlsbImportOptions.MaxMergedRanges; set => XlsbImportOptions.MaxMergedRanges = value; }

    /// <summary>For XLSB sources, controls the maximum number of worksheet hyperlinks accepted from one workbook.</summary>
    public int MaxHyperlinks { get => XlsbImportOptions.MaxHyperlinks; set => XlsbImportOptions.MaxHyperlinks = value; }

    /// <summary>For XLSB sources, controls the maximum number of UTF-16 characters accepted in one BIFF12 string.</summary>
    public int MaxStringCharacters { get => XlsbImportOptions.MaxStringCharacters; set => XlsbImportOptions.MaxStringCharacters = value; }

    /// <summary>For XLSB sources, controls whether unknown records are listed in the import preservation report.</summary>
    public bool ReportPreservedRecords { get => XlsbImportOptions.ReportPreservedRecords; set => XlsbImportOptions.ReportPreservedRecords = value; }

    /// <summary>For XLSB sources, controls the largest chart/VBA related-part graph projected from one workbook.</summary>
    public int MaxRelatedParts { get => XlsbImportOptions.MaxRelatedParts; set => XlsbImportOptions.MaxRelatedParts = value; }

    /// <summary>For XLSB sources, controls the largest relationship depth projected from chart and VBA carriers.</summary>
    public int MaxRelationshipDepth { get => XlsbImportOptions.MaxRelationshipDepth; set => XlsbImportOptions.MaxRelationshipDepth = value; }

}
