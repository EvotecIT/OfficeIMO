namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed record HtmlCorpusEvidenceReport(
    int SchemaVersion,
    DateTimeOffset GeneratedUtc,
    HtmlCorpusEvidenceEnvironment Environment,
    HtmlCorpusEvidenceSource Source,
    IReadOnlyList<HtmlCorpusCaseEvidence> Cases,
    HtmlCorpusAcceptanceEvidence? Acceptance,
    IReadOnlyList<string> Failures);

internal sealed record HtmlCorpusEvidenceEnvironment(
    string OsDescription,
    string ProcessArchitecture,
    string FrameworkDescription,
    string OfficeImoVersion,
    string PeachPdfVersion,
    string HtmlTinkerXVersion,
    string ChromiumVersion,
    string ExternalPdfRasterizer);

internal sealed record HtmlCorpusEvidenceSource(
    string CorpusId,
    string RelativeRoot,
    string? ManifestSha256,
    string? Commit,
    bool WorktreeDirty,
    int CaseCount);

internal sealed record HtmlCorpusCaseEvidence(
    string Id,
    string SourceRelativePath,
    long SourceBytes,
    string SourceSha256,
    IReadOnlyList<string> Capabilities,
    IReadOnlyList<string> TextMarkers,
    HtmlCorpusStaticEvidence? OfficeImo,
    HtmlCorpusPdfEvidence? PeachPdf,
    HtmlCorpusBrowserEvidence? Browser,
    HtmlCorpusComparisonEvidence Comparisons,
    IReadOnlyList<string> Failures);

internal sealed record HtmlCorpusStaticEvidence(
    HtmlCorpusOutputEvidence PrintPdf,
    HtmlCorpusOutputEvidence ScreenPng,
    HtmlCorpusOutputEvidence ScreenToPdf,
    HtmlCorpusSceneEvidence PrintScene,
    HtmlCorpusTextEvidence PrintText,
    HtmlCorpusTextEvidence ScreenText,
    IReadOnlyList<HtmlCorpusElementGeometry> ScreenElements,
    HtmlCorpusOperationMetrics PrintMetrics,
    HtmlCorpusOperationMetrics ScreenMetrics,
    HtmlCorpusOperationMetrics ScreenToPdfMetrics,
    double ElapsedMilliseconds,
    long ManagedAllocatedBytes,
    string ProfilePrint,
    string ProfileScreen,
    string ProfileScreenToPdf);

internal sealed record HtmlCorpusOperationMetrics(
    double ElapsedMilliseconds,
    long ManagedAllocatedBytes,
    long OutputBytes);

internal sealed record HtmlCorpusPdfEvidence(
    HtmlCorpusOutputEvidence Pdf,
    HtmlCorpusTextEvidence Text,
    double ElapsedMilliseconds,
    long ManagedAllocatedBytes);

internal sealed record HtmlCorpusBrowserEvidence(
    HtmlCorpusOutputEvidence PrintPdf,
    HtmlCorpusOutputEvidence ScreenPng,
    HtmlCorpusTextEvidence ScreenText,
    IReadOnlyList<HtmlCorpusElementGeometry> ScreenElements,
    double ElapsedMilliseconds);

internal sealed record HtmlCorpusComparisonEvidence(
    HtmlCorpusTextComparison? OfficeImoPrintToChromiumPrint,
    HtmlCorpusTextComparison? OfficeImoPrintToPeachPdf,
    HtmlCorpusGeometryComparison? ScreenGeometry,
    HtmlCorpusPixelComparison? ScreenPixels,
    HtmlCorpusScreenToPageComparison? ScreenToPage,
    IReadOnlyList<HtmlCorpusPageComparison> PrintPagesToChromium,
    IReadOnlyList<HtmlCorpusPageComparison> PrintPagesToPeachPdf);

internal sealed record HtmlCorpusOutputEvidence(
    string RelativePath,
    string MediaType,
    long SizeBytes,
    string Sha256,
    int PageCount,
    IReadOnlyList<HtmlCorpusPageArtifact> Pages);

internal sealed record HtmlCorpusPageArtifact(
    int PageNumber,
    string RelativePath,
    int Width,
    int Height,
    long SizeBytes,
    string Sha256,
    IReadOnlyList<string> Diagnostics);

internal sealed record HtmlCorpusSceneEvidence(
    int PageCount,
    int VisualCount,
    int DiagnosticCount,
    IReadOnlyList<string> DiagnosticCodes,
    IReadOnlyList<HtmlCorpusPageArtifact> RasterPages,
    IReadOnlyList<HtmlCorpusPageArtifact> SvgPages);

internal sealed record HtmlCorpusTextEvidence(
    int CharacterCount,
    string Sha256,
    IReadOnlyList<string> MissingMarkers,
    IReadOnlyList<HtmlCorpusMarkerEvidence> Markers);

internal sealed record HtmlCorpusMarkerEvidence(
    string Marker,
    string Policy,
    bool Matched);

internal sealed record HtmlCorpusTextComparison(
    double TokenRecall,
    double TokenPrecision,
    int ReferenceTokenCount,
    int CandidateTokenCount);

internal sealed record HtmlCorpusElementGeometry(
    string Key,
    string Source,
    int Index,
    double X,
    double Y,
    double Width,
    double Height);

internal sealed record HtmlCorpusGeometryComparison(
    int OfficeImoElementCount,
    int BrowserElementCount,
    int MatchedElementCount,
    double? MeanAbsoluteX,
    double? MeanAbsoluteY,
    double? MeanAbsoluteWidth,
    double? MeanAbsoluteHeight);

internal sealed record HtmlCorpusPixelComparison(
    bool DimensionsMatch,
    int ExpectedWidth,
    int ExpectedHeight,
    int ActualWidth,
    int ActualHeight,
    string Alignment,
    int ComparisonWidth,
    int ComparisonHeight,
    double? MeanAbsoluteError,
    double? RootMeanSquareError,
    double? MeanLuminanceError,
    string? DifferenceRelativePath);

internal sealed record HtmlCorpusScreenToPageComparison(
    int ScreenWidth,
    int ScreenHeight,
    int PageWidth,
    int CombinedPageHeight,
    int PageCount,
    int ExpectedPageWidth,
    int ExpectedPageHeight,
    int ExpectedPageCount,
    bool PageNumbersSequential,
    bool UniformPageDimensions,
    bool PageWidthsMatch,
    bool PageHeightsMatch,
    bool PageCountMatches,
    int ComparisonWidth,
    int ComparisonHeight,
    int ClippedWidth,
    int TrailingHeight,
    int ExpectedTrailingHeight,
    bool TrailingHeightMatches,
    bool CoversScreenHeight,
    double MeanAbsoluteError,
    double RootMeanSquareError,
    double MeanLuminanceError,
    string DifferenceRelativePath);

internal sealed record HtmlCorpusPageComparison(
    int PageNumber,
    bool PresentInBoth,
    HtmlCorpusPixelComparison? Pixels);

internal sealed class HtmlCorpusBrowserObservation {
    public string Text { get; set; } = string.Empty;
    public int ScrollWidth { get; set; }
    public int ScrollHeight { get; set; }
    public List<HtmlCorpusBrowserElement> Elements { get; set; } = new();
}

internal sealed class HtmlCorpusBrowserElement {
    public string Key { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
    public int Index { get; set; }
    public double X { get; set; }
    public double Y { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
}

internal sealed record HtmlCorpusAcceptanceEvidence(
    string ConfigurationSha256,
    string ReferencePolicy,
    bool Passed,
    IReadOnlyList<HtmlCorpusCaseAcceptance> Cases,
    IReadOnlyList<HtmlCorpusCapabilityAcceptance> Capabilities,
    IReadOnlyList<string> Failures);

internal sealed record HtmlCorpusCaseAcceptance(
    string Id,
    bool Passed,
    HtmlCorpusIntentAcceptance Screen,
    HtmlCorpusIntentAcceptance Print,
    HtmlCorpusIntentAcceptance ScreenToPage);

internal sealed record HtmlCorpusIntentAcceptance(
    string Intent,
    string Reference,
    string Classification,
    string Rationale,
    bool Passed,
    IReadOnlyList<HtmlCorpusAcceptanceCriterion> Criteria);

internal sealed record HtmlCorpusAcceptanceCriterion(
    string Id,
    double? Actual,
    string Comparison,
    double? Threshold,
    bool Passed,
    string? Detail = null);

internal sealed record HtmlCorpusCapabilityAcceptance(
    string ProfileId,
    string CapabilityId,
    IReadOnlyList<string> Intents,
    IReadOnlyList<string> RequiredCaseIds,
    bool Passed,
    IReadOnlyList<string> FailedCaseIds);
