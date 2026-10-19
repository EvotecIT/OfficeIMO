using System.Threading;

namespace OfficeIMO.Excel;

/// <summary>Resource limits for the explicit XML Spreadsheet 2003 data interchange profile.</summary>
public sealed class ExcelSpreadsheetXmlOptions {
    /// <summary>Maximum input bytes, or maximum materialized XLSX snapshot bytes on export. Default: 16 MiB.</summary>
    public int MaxInputBytes { get; set; } = 16 * 1024 * 1024;
    /// <summary>Maximum XML output bytes. Default: 64 MiB.</summary>
    public int MaxOutputBytes { get; set; } = 64 * 1024 * 1024;
    /// <summary>Maximum explicitly stored cells across all worksheets. Default: 1,000,000.</summary>
    public int MaxCells { get; set; } = 1_000_000;
    /// <summary>Maximum worksheets. Default: 256.</summary>
    public int MaxWorksheets { get; set; } = 256;
    /// <summary>Maximum style definitions. Default: 4096.</summary>
    public int MaxStyles { get; set; } = 4096;
    /// <summary>Maximum XML elements admitted before materialization, including unsupported content. Default: 1,000,000.</summary>
    public int MaxXmlElements { get; set; } = 1_000_000;
    /// <summary>Maximum XML nesting depth. Default: 64.</summary>
    public int MaxXmlDepth { get; set; } = 64;
    /// <summary>Cancellation checked during input, worksheet and cell processing.</summary>
    public CancellationToken CancellationToken { get; set; }

    internal ExcelSpreadsheetXmlOptions Snapshot() {
        if (MaxInputBytes < 1 || MaxOutputBytes < 1 || MaxCells < 1 || MaxWorksheets < 1 || MaxStyles < 1 || MaxXmlElements < 1 || MaxXmlDepth < 1)
            throw new ArgumentOutOfRangeException(nameof(ExcelSpreadsheetXmlOptions), "All XML Spreadsheet limits must be positive.");
        CancellationToken.ThrowIfCancellationRequested();
        return (ExcelSpreadsheetXmlOptions)MemberwiseClone();
    }
}
