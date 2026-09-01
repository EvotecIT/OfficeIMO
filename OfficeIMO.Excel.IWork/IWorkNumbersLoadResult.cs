using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

/// <summary>Contains an opt-in Excel projection and the Numbers source/report that produced it.</summary>
public sealed class IWorkNumbersLoadResult : IDisposable {
    internal IWorkNumbersLoadResult(ExcelDocument document, IWorkSourceDocument source,
        IWorkNumbersProjection projection, IWorkImportReport report) {
        Document = document;
        Source = source;
        Projection = projection;
        ImportReport = report;
    }

    /// <summary>Gets the normal editable OfficeIMO Excel document.</summary>
    public ExcelDocument Document { get; }
    /// <summary>Gets the bounded source package and preserved IWA records.</summary>
    public IWorkSourceDocument Source { get; }
    /// <summary>Gets the typed Numbers source projection.</summary>
    public IWorkNumbersProjection Projection { get; }
    /// <summary>Gets the loss-aware projection report.</summary>
    public IWorkImportReport ImportReport { get; }
    /// <summary>Gets whether the result uses a visual preview rather than editable reconstruction.</summary>
    public bool IsVisualFallback => ImportReport.ProjectionKind == IWorkProjectionKind.VisualFallback;
    /// <summary>Gets whether known source records or visuals were flattened or omitted.</summary>
    public bool HasConversionLoss => ImportReport.HasConversionLoss;

    /// <inheritdoc />
    public void Dispose() => Document.Dispose();
}
