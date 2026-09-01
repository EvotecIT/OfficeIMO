using OfficeIMO.IWork;

namespace OfficeIMO.Word.IWork;

/// <summary>Contains a Word conversion and the Pages source evidence that produced it.</summary>
public sealed class PagesToWordResult : IDisposable {
    internal PagesToWordResult(WordDocument value, IWorkSourceDocument source,
        IWorkPagesProjection projection, IWorkConversionReport report) {
        Value = value;
        Source = source;
        Projection = projection;
        Report = report;
    }

    /// <summary>Gets the converted editable OfficeIMO Word document.</summary>
    public WordDocument Value { get; }
    /// <summary>Gets the bounded source package and preserved IWA records.</summary>
    public IWorkSourceDocument Source { get; }
    /// <summary>Gets the typed Pages source projection.</summary>
    public IWorkPagesProjection Projection { get; }
    /// <summary>Gets the loss-aware conversion report.</summary>
    public IWorkConversionReport Report { get; }
    /// <summary>Gets whether the result uses a visual preview rather than editable reconstruction.</summary>
    public bool IsVisualFallback => Report.ProjectionKind == IWorkProjectionKind.VisualFallback;
    /// <summary>Gets whether known source records or visuals were flattened, omitted, or reported as errors.</summary>
    public bool HasLoss => Report.HasLoss;

    /// <summary>Returns the converted Word document.</summary>
    public WordDocument RequireValue() => Value;

    /// <summary>Returns the converted Word document or throws when the conversion was lossy.</summary>
    public WordDocument RequireNoLoss() {
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
