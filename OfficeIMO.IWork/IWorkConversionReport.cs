namespace OfficeIMO.IWork;

/// <summary>Loss-aware summary of one iWork-to-OfficeIMO projection.</summary>
public sealed class IWorkConversionReport : global::OfficeIMO.IOfficeConversionReport {
    internal IWorkConversionReport(IWorkDocumentKind sourceKind, IWorkProjectionKind projectionKind,
        IReadOnlyList<string> buildVersions, IReadOnlyList<IWorkArchiveRecord> unsupportedRecords,
        IReadOnlyList<IWorkDiagnostic> diagnostics, IWorkPreviewAsset? visualPreview,
        int totalRecordCount, int unsupportedRecordCount, int reconstructedItemCount) {
        SourceKind = sourceKind;
        ProjectionKind = projectionKind;
        BuildVersions = Array.AsReadOnly(buildVersions.ToArray());
        UnsupportedRecords = Array.AsReadOnly(unsupportedRecords.ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        VisualPreview = visualPreview;
        TotalRecordCount = totalRecordCount;
        UnsupportedRecordCount = unsupportedRecordCount;
        ReconstructedItemCount = reconstructedItemCount;
    }

    /// <summary>Gets the source iWork application.</summary>
    public IWorkDocumentKind SourceKind { get; }
    /// <summary>Gets whether the result contains editable reconstruction or a visual preview fallback.</summary>
    public IWorkProjectionKind ProjectionKind { get; }
    /// <summary>Gets producer build-history strings stored by the package.</summary>
    public IReadOnlyList<string> BuildVersions { get; }
    /// <summary>Gets preserved IWA payload records not losslessly represented by the typed projection, including partially consumed and auxiliary payloads.</summary>
    public IReadOnlyList<IWorkArchiveRecord> UnsupportedRecords { get; }
    /// <summary>Gets parser and projection diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets the preview used by visual fallback, when applicable.</summary>
    public IWorkPreviewAsset? VisualPreview { get; }
    /// <summary>Gets the total number of IWA payload records in the source.</summary>
    public int TotalRecordCount { get; }
    /// <summary>Gets the number of unprojected IWA payloads even when payload details were excluded by the read options.</summary>
    public int UnsupportedRecordCount { get; }
    /// <summary>Gets the number of semantic paragraphs, cells, slides, or other items reconstructed by the adapter.</summary>
    public int ReconstructedItemCount { get; }
    /// <summary>Gets whether the projection is known to omit or flatten source content.</summary>
    public bool HasLoss => ProjectionKind == IWorkProjectionKind.VisualFallback || UnsupportedRecordCount > 0 || HasErrors;
    /// <summary>Gets whether the parser or semantic projection reported an error diagnostic.</summary>
    public bool HasErrors => Diagnostics.Any(diagnostic => diagnostic.Severity == IWorkDiagnosticSeverity.Error);
    /// <summary>Gets whether the visual fallback is known to cover the complete source rather than a first-page or composite preview.</summary>
    public bool HasCompleteVisualCoverage => VisualPreview?.Coverage == IWorkVisualCoverage.FullDocument;

    /// <summary>Throws when the result is a visual fallback rather than editable reconstruction.</summary>
    public IWorkConversionReport RequireEditableReconstruction() {
        if (ProjectionKind != IWorkProjectionKind.EditableReconstruction) {
            throw new InvalidOperationException("The iWork source was projected as a visual fallback, not editable content.");
        }
        return this;
    }

    /// <summary>Throws when the projection reported errors.</summary>
    public IWorkConversionReport RequireNoErrors() {
        if (HasErrors) {
            throw new InvalidOperationException("The iWork conversion reported errors: "
                + string.Join("; ", Diagnostics.Where(diagnostic => diagnostic.Severity == IWorkDiagnosticSeverity.Error).Take(8)));
        }
        return this;
    }

    /// <summary>Throws when the projection used a visual fallback, left preserved records unprojected, or reported errors.</summary>
    public void RequireNoLoss() {
        if (HasLoss) {
            throw new InvalidOperationException(
                "The iWork conversion contains errors, visual fallback content, or preserved records that are not represented in the editable destination.");
        }
    }
}
