using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

/// <summary>Contains a PowerPoint projection and the Keynote source/report that produced it.</summary>
public sealed class IWorkKeynoteLoadResult : IDisposable {
    internal IWorkKeynoteLoadResult(PowerPointPresentation document, IWorkSourceDocument source,
        IWorkKeynoteProjection projection, IWorkImportReport report) {
        Document = document;
        Source = source;
        Projection = projection;
        ImportReport = report;
    }

    /// <summary>Gets the normal editable OfficeIMO PowerPoint presentation.</summary>
    public PowerPointPresentation Document { get; }
    /// <summary>Gets the bounded source package and preserved IWA records.</summary>
    public IWorkSourceDocument Source { get; }
    /// <summary>Gets the typed Keynote source projection.</summary>
    public IWorkKeynoteProjection Projection { get; }
    /// <summary>Gets the loss-aware projection report.</summary>
    public IWorkImportReport ImportReport { get; }
    /// <summary>Gets whether the result uses a visual preview rather than editable reconstruction.</summary>
    public bool IsVisualFallback => ImportReport.ProjectionKind == IWorkProjectionKind.VisualFallback;
    /// <summary>Gets whether known source records or visuals were flattened or omitted.</summary>
    public bool HasConversionLoss => ImportReport.HasConversionLoss;

    /// <inheritdoc />
    public void Dispose() => Document.Dispose();
}
