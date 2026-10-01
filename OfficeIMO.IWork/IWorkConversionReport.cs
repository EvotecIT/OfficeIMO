namespace OfficeIMO.IWork;

/// <summary>Loss-aware summary of one iWork-to-OfficeIMO projection.</summary>
public sealed class IWorkConversionReport : global::OfficeIMO.IOfficeConversionReport {
    internal IWorkConversionReport(IWorkDocumentKind sourceKind, IWorkProjectionKind projectionKind,
        IReadOnlyList<string> buildVersions, IReadOnlyList<IWorkArchiveRecord> preservedRecords,
        IReadOnlyList<IWorkDiagnostic> diagnostics, IWorkPreviewAsset? visualPreview,
        int totalRecordCount, int preservedRecordCount, int reconstructedItemCount,
        IReadOnlyList<IWorkSourceUnit>? sourceUnits = null,
        IReadOnlyList<IWorkFormulaCellStatus>? formulaCells = null,
        IReadOnlyList<IWorkSourceReferenceIssue>? sourceReferenceIssues = null,
        IReadOnlyList<IWorkSourceDeclarationIssue>? sourceDeclarationIssues = null,
        IReadOnlyList<IWorkSourceCellIssue>? sourceCellIssues = null) {
        SourceKind = sourceKind;
        ProjectionKind = projectionKind;
        BuildVersions = Array.AsReadOnly(buildVersions.ToArray());
        PreservedRecords = Array.AsReadOnly(preservedRecords.ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        VisualPreview = visualPreview;
        TotalRecordCount = totalRecordCount;
        PreservedRecordCount = preservedRecordCount;
        ReconstructedItemCount = reconstructedItemCount;
        SourceUnits = Array.AsReadOnly((sourceUnits ?? Array.Empty<IWorkSourceUnit>()).ToArray());
        SourceUnitCounts = Array.AsReadOnly(((IWorkSourceUnitKind[])Enum.GetValues(typeof(IWorkSourceUnitKind)))
            .Select(kind => new IWorkSourceUnitCount(kind, SourceUnits)).ToArray());
        FormulaCells = Array.AsReadOnly((formulaCells ?? Array.Empty<IWorkFormulaCellStatus>()).ToArray());
        FormulaSummary = new IWorkFormulaSummary(FormulaCells);
        SourceReferenceIssues = Array.AsReadOnly((sourceReferenceIssues ?? Array.Empty<IWorkSourceReferenceIssue>()).ToArray());
        SourceDeclarationIssues = Array.AsReadOnly((sourceDeclarationIssues ?? Array.Empty<IWorkSourceDeclarationIssue>()).ToArray());
        SourceCellIssues = Array.AsReadOnly((sourceCellIssues ?? Array.Empty<IWorkSourceCellIssue>()).ToArray());
        var fidelityDiagnostics = new List<global::OfficeIMO.OfficeConversionFidelityDiagnostic>();
        foreach (IWorkDiagnostic diagnostic in Diagnostics) {
            fidelityDiagnostics.Add(new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.LossKind,
                "OfficeIMO.IWork",
                diagnostic.EntryPath));
        }
        if (ProjectionKind == IWorkProjectionKind.VisualFallback) {
            fidelityDiagnostics.Add(new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
                "IWORK_VISUAL_FALLBACK",
                "The source was represented by a visual preview instead of editable reconstruction.",
                VisualPreview?.Coverage == IWorkVisualCoverage.FullDocument
                    ? global::OfficeIMO.OfficeConversionLossKind.Approximation
                    : global::OfficeIMO.OfficeConversionLossKind.Omission,
                "OfficeIMO.IWork"));
        }
        if (UnassessedRecordCount > 0) {
            fidelityDiagnostics.Add(new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
                "IWORK_RECORD_FIDELITY_UNASSESSED",
                UnassessedRecordCount + " source record(s) remain available for preservation, but record-level fidelity has not been assessed. This count is not a count of omitted content.",
                global::OfficeIMO.OfficeConversionLossKind.Unassessed,
                "OfficeIMO.IWork"));
        }
        if (SourceReferenceIssues.Count > 0) {
            fidelityDiagnostics.Add(new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
                "IWORK_SOURCE_REFERENCES_UNRESOLVED",
                SourceReferenceIssues.Count + " declared reference occurrence(s) in assessed content paths could not be resolved. These occurrences do not identify omitted objects or establish visual coverage.",
                global::OfficeIMO.OfficeConversionLossKind.Unassessed,
                "OfficeIMO.IWork"));
        }
        if (SourceDeclarationIssues.Count > 0) {
            fidelityDiagnostics.Add(new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
                "IWORK_SOURCE_DECLARATIONS_UNASSESSED",
                SourceDeclarationIssues.Count + " selected source path(s) contain unreadable, rejected or unsupported declarations. Their nested references and content coverage remain unknown; these are not omitted-object counts.",
                global::OfficeIMO.OfficeConversionLossKind.Unassessed,
                "OfficeIMO.IWork"));
        }
        if (SourceCellIssues.Count > 0) {
            fidelityDiagnostics.Add(new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
                "IWORK_SOURCE_CELLS_UNDECODED",
                SourceCellIssues.Count + " materialized selected cell(s) could not be decoded. Their source contents remain unassessed; this is not a count of omitted cells or unreadable storage entries.",
                global::OfficeIMO.OfficeConversionLossKind.Unassessed,
                "OfficeIMO.IWork"));
        }
        if (FormulaSummary.UnassessedExpressionCount > 0) {
            fidelityDiagnostics.Add(new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
                "IWORK_FORMULA_CELLS_UNASSESSED",
                "Supported cell headers declare formulas whose contents could not be decoded; their expressions and caches remain unassessed.",
                global::OfficeIMO.OfficeConversionLossKind.Unassessed,
                "OfficeIMO.IWork"));
        }
        FidelityDiagnostics = fidelityDiagnostics.AsReadOnly();
    }

    /// <summary>Gets the source iWork application.</summary>
    public IWorkDocumentKind SourceKind { get; }
    /// <summary>Gets whether the result contains editable reconstruction or a visual preview fallback.</summary>
    public IWorkProjectionKind ProjectionKind { get; }
    /// <summary>Gets producer build-history strings stored by the package.</summary>
    public IReadOnlyList<string> BuildVersions { get; }
    /// <summary>Gets source IWA payloads retained for inspection, including consumed, partially consumed, and auxiliary records. Presence here does not imply content was omitted.</summary>
    public IReadOnlyList<IWorkArchiveRecord> PreservedRecords { get; }
    /// <summary>Gets parser and projection diagnostics.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
    /// <summary>Gets category-preserving diagnostics for downstream acceptance policies.</summary>
    public IReadOnlyList<global::OfficeIMO.OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; }
    /// <summary>Gets the preview used by visual fallback, when applicable.</summary>
    public IWorkPreviewAsset? VisualPreview { get; }
    /// <summary>Gets the total number of IWA payload records in the source.</summary>
    public int TotalRecordCount { get; }
    /// <summary>Gets the number of source IWA payloads covered by preservation accounting, even when report payload details were excluded by the read options.</summary>
    public int PreservedRecordCount { get; }
    /// <summary>Gets the source record count without a field-level fidelity assessment. Auxiliary and consumed records are included; this is not an omission count.</summary>
    public int UnassessedRecordCount => TotalRecordCount;
    /// <summary>Gets the number of semantic paragraphs, cells, slides, or other items reconstructed by the adapter.</summary>
    public int ReconstructedItemCount { get; }
    /// <summary>Gets identified primary document, sheet, slide, table-info, text-storage, image, and explicitly omitted unsupported records selected by the semantic projection, with known destination outcomes. Inactive records and unresolved references are excluded; this is not a complete content inventory.</summary>
    public IReadOnlyList<IWorkSourceUnit> SourceUnits { get; }
    /// <summary>Gets per-kind counts of identified selected source units. Reconstructed units can still contain omitted, approximated, or unassessed fields; visual fallback does not establish individual unit coverage.</summary>
    public IReadOnlyList<IWorkSourceUnitCount> SourceUnitCounts { get; }
    /// <summary>Gets unresolved declared reference occurrences in assessed document, drawable, table and text-formatting fields. This is not a complete dependency inventory or a count of omitted objects. Repeated occurrences remain separate, and visual fallback does not resolve them.</summary>
    public IReadOnlyList<IWorkSourceReferenceIssue> SourceReferenceIssues { get; }
    /// <summary>Gets unreadable or rejected declarations at selected source paths, separately from resolved object identities and failed references. Repeated selection of a shared path reports it once. Visual fallback does not establish coverage.</summary>
    public IReadOnlyList<IWorkSourceDeclarationIssue> SourceDeclarationIssues { get; }
    /// <summary>Gets decoding failures for materialized selected cells, including visual fallback. Native error markers, inactive tables and unmaterialized storage are excluded. The inventory is bounded by the source-wide materialized-cell limit and does not establish destination cell outcomes.</summary>
    public IReadOnlyList<IWorkSourceCellIssue> SourceCellIssues { get; }
    /// <summary>Gets source expression and cache assessments for declared formula cells, including supported headers in undecoded cells and visual fallback. These do not establish destination formula preservation or cache freshness.</summary>
    public IReadOnlyList<IWorkFormulaCellStatus> FormulaCells { get; }
    /// <summary>Gets aggregate source expression/cache assessment, including unassessed declared formulas. Unknown headers and inactive table records are excluded.</summary>
    public IWorkFormulaSummary FormulaSummary { get; }
    /// <summary>Gets whether any typed fidelity diagnostic reports omission, approximation, failure, or unassessed fidelity.</summary>
    public bool HasLoss => FidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != global::OfficeIMO.OfficeConversionLossKind.None);
    /// <summary>Gets whether the parser or semantic projection reported an error diagnostic.</summary>
    public bool HasErrors => Diagnostics.Any(diagnostic => diagnostic.Severity == IWorkDiagnosticSeverity.Error);
    /// <summary>Gets whether the explicit partial-reconstruction policy was needed to retain editable content.</summary>
    public bool IsPartialEditableReconstruction => ProjectionKind == IWorkProjectionKind.EditableReconstruction
        && Diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_PARTIAL_EDITABLE_RECONSTRUCTION");

    /// <summary>Throws unless the output is an editable reconstruction without explicitly partial content.</summary>
    public IWorkConversionReport RequireCompleteEditableReconstruction() {
        RequireEditableReconstruction();
        if (IsPartialEditableReconstruction) {
            throw new InvalidOperationException("The iWork source contains explicitly partial editable reconstruction.");
        }
        return this;
    }

    /// <summary>Throws unless a visual fallback is known to cover the complete source.</summary>
    public IWorkConversionReport RequireCompleteVisualCoverage() {
        if (ProjectionKind != IWorkProjectionKind.VisualFallback || !HasCompleteVisualCoverage) {
            throw new InvalidOperationException("The output is not a visual fallback with known complete document coverage.");
        }
        return this;
    }

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

    /// <summary>Throws when any typed fidelity diagnostic reports omission, approximation, failure, or unassessed fidelity.</summary>
    public void RequireNoLoss() {
        if (HasLoss) {
            throw new InvalidOperationException(
                "The iWork conversion contains a typed fidelity diagnostic reporting omission, approximation, failure, or unassessed fidelity.");
        }
    }
}
