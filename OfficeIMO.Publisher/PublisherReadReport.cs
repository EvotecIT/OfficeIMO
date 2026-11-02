namespace OfficeIMO.Publisher;

/// <summary>Immutable source-decoding and page-projection evidence. Recovered objects do not imply complete native fidelity.</summary>
public sealed class PublisherReadReport : OfficeConversionReport {
    internal PublisherReadReport(IEnumerable<OfficeDiagnostic> diagnostics, int sourceObjects,
        int projectedObjects, int stories, int images, int records) {
        FidelityDiagnostics = Array.AsReadOnly(diagnostics.ToArray());
        SourceObjectCount = sourceObjects; ProjectedObjectCount = projectedObjects;
        TextStoryCount = stories; ImageCount = images; InspectedRecordCount = records;
    }
    /// <summary>Diagnostics identifying approximated, omitted, or unassessed source content.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics { get; }
    /// <summary>Number of publication shape, group, and table records identified in the native directory.</summary>
    public int SourceObjectCount { get; }
    /// <summary>Number of distinct source objects represented on a document or master page. Repeated master use counts once.</summary>
    public int ProjectedObjectCount { get; }
    /// <summary>Number of decoded text stories, including stories not placed on a page.</summary>
    public int TextStoryCount { get; }
    /// <summary>Number of images with recovered payloads, including unplaced images.</summary>
    public int ImageCount { get; }
    /// <summary>Number of inspected native records across all source streams.</summary>
    public int InspectedRecordCount { get; }
    /// <inheritdoc />
    public override bool HasLoss => FidelityDiagnostics.Any(item => item.LossKind != OfficeConversionLossKind.None);
    /// <inheritdoc />
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("Publisher recovery contains approximated, omitted, or unassessed source content. Inspect FidelityDiagnostics.", this);
}
