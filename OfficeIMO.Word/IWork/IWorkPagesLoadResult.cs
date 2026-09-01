using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

/// <summary>Contains a Word projection and the Pages source/report that produced it.</summary>
public sealed class IWorkPagesLoadResult : IDisposable {
    internal IWorkPagesLoadResult(WordDocument document, IWorkSourceDocument source,
        IWorkPagesProjection projection, IWorkImportReport report) {
        Document = document;
        Source = source;
        Projection = projection;
        ImportReport = report;
    }

    /// <summary>Gets the normal editable OfficeIMO Word document.</summary>
    public WordDocument Document { get; }
    /// <summary>Gets the bounded source package and preserved IWA records.</summary>
    public IWorkSourceDocument Source { get; }
    /// <summary>Gets the typed Pages source projection.</summary>
    public IWorkPagesProjection Projection { get; }
    /// <summary>Gets the loss-aware projection report.</summary>
    public IWorkImportReport ImportReport { get; }
    /// <summary>Gets whether the result uses a visual preview rather than editable reconstruction.</summary>
    public bool IsVisualFallback => ImportReport.ProjectionKind == IWorkProjectionKind.VisualFallback;
    /// <summary>Gets whether known source records or visuals were flattened or omitted.</summary>
    public bool HasConversionLoss => ImportReport.HasConversionLoss;

    /// <inheritdoc />
    public void Dispose() => Document.Dispose();
}
