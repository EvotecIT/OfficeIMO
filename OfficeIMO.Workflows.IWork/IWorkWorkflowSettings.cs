using OfficeIMO.IWork;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.PowerPoint.Pdf;

namespace OfficeIMO.Workflows.IWork;

/// <summary>Per-request read bounds and acceptance choices for an Apple workflow conversion.</summary>
public sealed class IWorkWorkflowSettings : IOfficeConversionSettings {
    /// <summary>Source package and materialization limits. The workflow input limit also applies.</summary>
    public IWorkReadOptions ReadOptions { get; set; } = new() { PreserveSourceRecords = false };
    /// <summary>Representation, partial reconstruction, preview coverage, and worksheet naming choices.</summary>
    public IWorkConversionOptions ConversionOptions { get; set; } = new() { RequireCompleteVisualCoverage = true };
    /// <summary>Pages PDF layout and font options, snapshotted for each request.</summary>
    public WordPdfExportOptions PagesPdfOptions { get; set; } = new();
    /// <summary>Numbers PDF layout and font options. Sheet/row selection and print-area clipping are rejected.</summary>
    public ExcelPdfExportOptions NumbersPdfOptions { get; set; } = new() {
        RespectWorkbookSheetVisibility = false, RespectWorksheetHiddenRowsAndColumns = false,
        UseWorksheetPrintAreas = false
    };
    /// <summary>Keynote PDF layout and font options. Complete routes render all slides.</summary>
    public PowerPointPdfExportOptions KeynotePdfOptions { get; set; } = new() { IncludeHiddenSlides = true };

    /// <summary>Validates and copies all mutable options for one request.</summary>
    /// <exception cref="NotSupportedException">Registered PDF options request encryption. Use the PDF protection workflow after conversion.</exception>
    public IWorkWorkflowSettings Clone() {
        var snapshot = new IWorkWorkflowSettings {
            ReadOptions = (ReadOptions ?? throw new ArgumentException("iWork read options cannot be null.")).Clone(),
            ConversionOptions = (ConversionOptions ?? throw new ArgumentException("iWork conversion options cannot be null.")).Clone(),
            PagesPdfOptions = (PagesPdfOptions ?? throw new ArgumentException("Pages PDF options cannot be null.")).Clone(),
            NumbersPdfOptions = (NumbersPdfOptions ?? throw new ArgumentException("Numbers PDF options cannot be null.")).Clone(),
            KeynotePdfOptions = (KeynotePdfOptions ?? throw new ArgumentException("Keynote PDF options cannot be null.")).Clone()
        };
        if (snapshot.PagesPdfOptions.PdfOptions?.Encryption != null
            || snapshot.NumbersPdfOptions.PdfOptions?.Encryption != null
            || snapshot.KeynotePdfOptions.PdfOptions?.Encryption != null)
            throw new NotSupportedException("Registered iWork PDF options do not support encryption. Use the PDF protection workflow after conversion.");
        return snapshot;
    }

    /// <inheritdoc />
    public IOfficeConversionSettings Snapshot() => Clone();
}
