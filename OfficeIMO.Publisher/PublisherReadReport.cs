namespace OfficeIMO.Publisher;

/// <summary>Immutable source-decoding and page-projection evidence. Recovered objects do not imply complete native fidelity.</summary>
public sealed class PublisherReadReport : IOfficeConversionReport {
    internal PublisherReadReport(IEnumerable<OfficeConversionFidelityDiagnostic> diagnostics, int sourceObjects,
        int projectedObjects, int stories, int images, int records, int? currentObjects = null) {
        FidelityDiagnostics = Array.AsReadOnly(diagnostics.ToArray());
        SourceObjectCount = sourceObjects; ProjectedObjectCount = projectedObjects;
        CurrentObjectCount = currentObjects ?? sourceObjects;
        TextStoryCount = stories; ImageCount = images; InspectedRecordCount = records;
    }
    /// <summary>Diagnostics identifying approximated, omitted, or unassessed source content.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <summary>Number of publication shape, group, and table records identified in the native directory.</summary>
    public int SourceObjectCount { get; }
    /// <summary>Number of current publication objects, including retained imported objects, authored tables, shapes and continuation frames.</summary>
    public int CurrentObjectCount { get; }
    /// <summary>Number of distinct current objects represented on a document or master page, including authored tables, shapes and frames. Repeated master use counts once.</summary>
    public int ProjectedObjectCount { get; }
    /// <summary>Number of current complete text stories, including authored table stories and stories not placed on a page.</summary>
    public int TextStoryCount { get; }
    /// <summary>Number of current embedded assets, including recovered unplaced images and assets retained by supported picture edits.</summary>
    public int ImageCount { get; }
    /// <summary>Number of inspected native records across all source streams.</summary>
    public int InspectedRecordCount { get; }
    /// <inheritdoc />
    public bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    public void RequireNoLoss() {
        if (HasLoss) throw new OfficeConversionException("Publisher recovery contains approximated, omitted, or unassessed source content. Inspect FidelityDiagnostics.", this);
    }
}
