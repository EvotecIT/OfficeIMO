namespace OfficeIMO.Html.Benchmarks;

internal sealed record HtmlQualificationBaselineReport(
    string Schema,
    string SchemaVersion,
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool TrackedSourceTreeDirty,
    int UntrackedSourcePathCount,
    IReadOnlyList<string> UntrackedSourceRoots,
    HtmlQualificationCorpusEvidence Corpus,
    HtmlQualificationEnvironmentEvidence Environment,
    IReadOnlyList<HtmlQualificationProviderEvidence> Providers,
    HtmlQualificationDocumentEvidence Document,
    IReadOnlyList<HtmlQualificationProfileEvidence> Profiles,
    HtmlQualificationCancellationEvidence Cancellation,
    IReadOnlyList<string> Failures);

internal sealed record HtmlQualificationCorpusEvidence(
    string Id,
    string ManifestSha256,
    string EntryPath,
    string BaseUri,
    string SourceKind,
    string License,
    IReadOnlyList<HtmlQualificationFileEvidence> Files);

internal sealed record HtmlQualificationFileEvidence(
    string Path,
    string Role,
    string MediaType,
    long Length,
    string Sha256);

internal sealed record HtmlQualificationEnvironmentEvidence(
    string Framework,
    string OperatingSystem,
    string Architecture,
    string Processor,
    string PhysicalCoreCount,
    int LogicalProcessorCount,
    string ProcessPriority,
    string ProcessAffinity,
    string ActivePowerPlan);

internal sealed record HtmlQualificationProviderEvidence(
    string Assembly,
    string Version,
    string InformationalVersion);

internal sealed record HtmlQualificationDocumentEvidence(
    int InputBytes,
    string InputSha256,
    int NodeCount,
    int ElementCount,
    IReadOnlyList<HtmlQualificationSelectorEvidence> Selectors,
    IReadOnlyDictionary<string, int> LogicalNodeCounts,
    IReadOnlyList<string> LogicalCapabilities,
    int SemanticSectionCount,
    int SemanticBlockCount,
    int SemanticRootTableCount,
    int SemanticTableBlockCount,
    int SemanticResourceCount,
    int SourceStyledElementCount,
    int SourceStylePropertyCount,
    int AllowedResourceCount,
    int BlockedResourceCount,
    string NormalizedHtmlSha256,
    IReadOnlyList<HtmlQualificationDiagnosticEvidence> Diagnostics,
    IReadOnlyList<HtmlQualificationStageObservation> Stages);

internal sealed record HtmlQualificationSelectorEvidence(
    string Selector,
    int ExpectedCount,
    int ActualCount,
    bool Passed);

internal sealed record HtmlQualificationProfileEvidence(
    string Id,
    string Mode,
    string Media,
    double ViewportWidth,
    double ViewportHeight,
    int ExpectedPageCount,
    int ActualPageCount,
    int TextCharacters,
    string TextSha256,
    int HeadingCount,
    int ResolvedResourceCount,
    long ResolvedResourceBytes,
    IReadOnlyList<HtmlQualificationStyleEvidence> Styles,
    IReadOnlyList<HtmlQualificationResourceEvidence> Resources,
    IReadOnlyList<HtmlQualificationResourceEvidence> RenderResources,
    IReadOnlyList<HtmlQualificationResourceEvidence> PdfResources,
    HtmlQualificationPdfResourcePolicyEvidence? PdfResourcePolicy,
    IReadOnlyList<HtmlQualificationPageEvidence> Pages,
    IReadOnlyList<HtmlQualificationArtifactEvidence> Artifacts,
    IReadOnlyList<HtmlQualificationDiagnosticEvidence> Diagnostics,
    IReadOnlyList<HtmlQualificationPdfWarningEvidence> PdfWarnings,
    IReadOnlyList<HtmlQualificationStageObservation> Stages,
    bool Passed);

internal sealed record HtmlQualificationStyleEvidence(
    string Selector,
    string Property,
    string Value,
    string ExpectedContains,
    HtmlQualificationCascadeEvidence? Cascade,
    bool Passed);

internal sealed record HtmlQualificationCascadeEvidence(
    bool Inherited,
    bool Important,
    bool Inline,
    bool Layered,
    int SpecificityIds,
    int SpecificityClasses,
    int SpecificityElements,
    int RuleOrder,
    int DeclarationOrder);

internal sealed record HtmlQualificationResourceEvidence(
    string Uri,
    string Kind,
    string Path,
    string MediaType,
    long Length,
    string Sha256);

internal sealed record HtmlQualificationPageEvidence(
    int PageNumber,
    double Width,
    double Height,
    int SceneVisualCount,
    int VisualCount,
    string SvgSha256,
    string PngSha256);

internal sealed record HtmlQualificationArtifactEvidence(
    string Path,
    string MediaType,
    long Length,
    string Sha256,
    int? PageNumber);

internal sealed record HtmlQualificationDiagnosticEvidence(
    string Code,
    string Severity,
    string LossKind,
    string Message,
    string? Source);

internal sealed record HtmlQualificationPdfWarningEvidence(
    string Converter,
    string Code,
    string Severity,
    string LossKind,
    string Message,
    string Source);

internal sealed record HtmlQualificationPdfResourcePolicyEvidence(
    bool AllowSystemFontEmbedding,
    bool AllowDocumentFontEmbedding,
    bool AllowLocalFileAccess,
    bool AllowRemoteResourceResolution,
    bool AllowDataUris,
    bool AllowEmbeddedPackageResources);

internal sealed record HtmlQualificationStageObservation(
    string Stage,
    string MeasurementKind,
    double ElapsedMilliseconds,
    long ProcessAllocatedBytes);

internal sealed record HtmlQualificationCancellationEvidence(
    bool PreCanceledRenderStopped,
    int ResolverCalls,
    bool ProducedRender);
