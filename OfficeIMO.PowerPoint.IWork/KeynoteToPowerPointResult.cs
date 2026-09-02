using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

/// <summary>Contains a PowerPoint conversion and the Keynote source evidence that produced it.</summary>
public sealed class KeynoteToPowerPointResult : IDisposable {
    internal KeynoteToPowerPointResult(PowerPointPresentation value, IWorkSourceDocument source,
        IWorkKeynoteProjection projection, IWorkConversionReport report) {
        Value = value;
        Source = source;
        Projection = projection;
        Report = report;
    }

    /// <summary>Gets the converted editable OfficeIMO PowerPoint presentation.</summary>
    public PowerPointPresentation Value { get; }
    /// <summary>Gets the bounded source package and preserved IWA records.</summary>
    public IWorkSourceDocument Source { get; }
    /// <summary>Gets the typed Keynote source projection.</summary>
    public IWorkKeynoteProjection Projection { get; }
    /// <summary>Gets the loss-aware conversion report.</summary>
    public IWorkConversionReport Report { get; }
    /// <summary>Gets whether the result uses a visual preview rather than editable reconstruction.</summary>
    public bool IsVisualFallback => Report.ProjectionKind == IWorkProjectionKind.VisualFallback;
    /// <summary>Gets whether known source records or visuals were flattened, omitted, or reported as errors.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the converted PowerPoint presentation.</summary>
    public PowerPointPresentation RequireValue() => Value;

    /// <summary>Returns the converted PowerPoint presentation or throws when the conversion was lossy.</summary>
    public PowerPointPresentation RequireNoLoss() {
        try {
            Report.RequireNoLoss();
            return Value;
        } catch {
            Value.Dispose();
            throw;
        }
    }

    /// <inheritdoc />
    public void Dispose() => Value.Dispose();
}
