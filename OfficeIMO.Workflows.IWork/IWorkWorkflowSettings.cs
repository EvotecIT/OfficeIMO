using OfficeIMO.IWork;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.PowerPoint.Pdf;

namespace OfficeIMO.Workflows.IWork;

/// <summary>Per-request read bounds and acceptance choices for an Apple workflow conversion.</summary>
public sealed class IWorkWorkflowSettings : IOfficeWorkflowConversionSettings {
    /// <summary>Source package and materialization limits. The workflow input limit also applies.</summary>
    public IWorkReadOptions ReadOptions { get; set; } = new() { PreserveSourceRecords = false };
    /// <summary>Representation, partial reconstruction, preview coverage, and worksheet naming choices.</summary>
    public IWorkConversionOptions ConversionOptions { get; set; } = new() { RequireCompleteVisualCoverage = true };
    /// <summary>Pages PDF layout and font options, snapshotted for each request.</summary>
    public WordToPdfOptions PagesPdfOptions { get; set; } = new();
    /// <summary>Numbers PDF layout and font options. Sheet/row selection and print-area clipping are rejected.</summary>
    public ExcelToPdfOptions NumbersPdfOptions { get; set; } = new() {
        RespectWorkbookSheetVisibility = false, RespectWorksheetHiddenRowsAndColumns = false,
        UseWorksheetPrintAreas = false
    };
    /// <summary>Keynote PDF layout and font options. Complete routes render all slides.</summary>
    public PowerPointToPdfOptions KeynotePdfOptions { get; set; } = new() { IncludeHiddenSlides = true };

    /// <summary>Validates and copies all mutable options for one request.</summary>
    public IWorkWorkflowSettings Clone() => new() {
        ReadOptions = (ReadOptions ?? throw new ArgumentException("iWork read options cannot be null.")).Clone(),
        ConversionOptions = (ConversionOptions ?? throw new ArgumentException("iWork conversion options cannot be null.")).Clone(),
        PagesPdfOptions = (PagesPdfOptions ?? throw new ArgumentException("Pages PDF options cannot be null.")).Clone(),
        NumbersPdfOptions = (NumbersPdfOptions ?? throw new ArgumentException("Numbers PDF options cannot be null.")).Clone(),
        KeynotePdfOptions = (KeynotePdfOptions ?? throw new ArgumentException("Keynote PDF options cannot be null.")).Clone()
    };

    /// <inheritdoc />
    public IOfficeWorkflowConversionSettings Snapshot() => Clone();
}
