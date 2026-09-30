using System.Globalization;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint.IWork;
using OfficeIMO.Word.IWork;

namespace OfficeIMO.Workflows.IWork;

/// <summary>Opt-in Apple iWork conversion using the bounded source reader and existing Office destination owners.</summary>
public static class IWorkWorkflow {
    /// <summary>Creates a runner with Pages-to-Word, Numbers-to-Excel, and Keynote-to-PowerPoint routes.</summary>
    /// <remarks>Defaults reject partial editable reconstruction and visual previews without known full-document coverage. Local ZIP files and directory bundles, and provider ZIP streams, use captured inputs and publication-time source verification.</remarks>
    public static OfficeWorkflowRunner CreateRunner(IWorkReadOptions? readOptions = null,
        IWorkConversionOptions? conversionOptions = null) => new(null, null, conversions: CreateRegistrations(readOptions, conversionOptions));

    /// <summary>Captures independent read and acceptance options for registration alongside other format adapters.</summary>
    public static IReadOnlyList<OfficeWorkflowConversionRegistration> CreateRegistrations(IWorkReadOptions? readOptions = null,
        IWorkConversionOptions? conversionOptions = null) {
        IWorkReadOptions reading = (readOptions ?? new IWorkReadOptions { PreserveSourceRecords = false }).Clone();
        IWorkConversionOptions conversion = (conversionOptions ?? new IWorkConversionOptions { RequireCompleteVisualCoverage = true }).Clone();
        return Array.AsReadOnly(new[] {
            OfficeWorkflowConversionRegistration.Create<IWorkWorkflowSettings>("pages-docx", (input, output, limits, settings, token) => {
                using PagesToWordResult result = WordIWorkConverter.ConvertPagesToWordResult(input, Bound(settings?.ReadOptions ?? reading, limits), settings?.ConversionOptions ?? conversion, token);
                result.Value.SaveAsync(output, token).GetAwaiter().GetResult();
                return Evidence(result.Report);
            }, (path, limits, settings) => DirectoryInput(path, Bound(settings?.ReadOptions ?? reading, limits), limits)),
            OfficeWorkflowConversionRegistration.Create<IWorkWorkflowSettings>("numbers-xlsx", (input, output, limits, settings, token) => {
                using NumbersToExcelResult result = ExcelIWorkConverter.ConvertNumbersToExcelResult(input, Bound(settings?.ReadOptions ?? reading, limits), settings?.ConversionOptions ?? conversion, token);
                result.Value.SaveAsync(output, token).GetAwaiter().GetResult();
                return Evidence(result.Report);
            }, (path, limits, settings) => DirectoryInput(path, Bound(settings?.ReadOptions ?? reading, limits), limits)),
            OfficeWorkflowConversionRegistration.Create<IWorkWorkflowSettings>("keynote-pptx", (input, output, limits, settings, token) => {
                using KeynoteToPowerPointResult result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(input, Bound(settings?.ReadOptions ?? reading, limits), settings?.ConversionOptions ?? conversion, token);
                result.Value.SaveAsync(output, token).GetAwaiter().GetResult();
                return Evidence(result.Report);
            }, (path, limits, settings) => DirectoryInput(path, Bound(settings?.ReadOptions ?? reading, limits), limits))
        });
    }

    private static OfficeWorkflowStreamInput DirectoryInput(string path, IWorkReadOptions reading, OfficeWorkflowLimits limits) {
        var snapshot = new IWorkDirectoryPackageSnapshot(path, reading, limits.MaximumInputBytes);
        return new OfficeWorkflowStreamInput(Path.GetFileName(path), token => Task.Run(() => snapshot.OpenStream(token), token), null, OfficeWorkflowSourceSnapshotKind.DirectoryPackage, new DirectorySourceGuard(snapshot));
    }

    private sealed class DirectorySourceGuard(IWorkDirectoryPackageSnapshot snapshot) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) =>
            ValueTask.FromResult(snapshot.CanPublish(path, isDirectory, token));
    }

    private static IWorkReadOptions Bound(IWorkReadOptions options, OfficeWorkflowLimits limits) {
        IWorkReadOptions copy = options.Clone();
        copy.MaximumPackageBytes = Math.Min(copy.MaximumPackageBytes, limits.MaximumInputBytes);
        return copy;
    }

    private static OfficeWorkflowConversionEvidence Evidence(IWorkConversionReport report) => new(report,
        new Dictionary<string, string>(StringComparer.Ordinal) {
            ["sourceKind"] = report.SourceKind.ToString(),
            ["projectionKind"] = report.ProjectionKind.ToString(),
            ["producerBuildVersions"] = string.Join(", ", report.BuildVersions),
            ["totalRecordCount"] = report.TotalRecordCount.ToString(CultureInfo.InvariantCulture),
            ["unassessedRecordCount"] = report.UnassessedRecordCount.ToString(CultureInfo.InvariantCulture),
            ["reconstructedItemCount"] = report.ReconstructedItemCount.ToString(CultureInfo.InvariantCulture),
            ["sourceUnitCount"] = report.SourceUnits.Count.ToString(CultureInfo.InvariantCulture),
            ["reconstructedSourceUnitCount"] = report.SourceUnitCounts.Sum(count => count.ReconstructedCount).ToString(CultureInfo.InvariantCulture),
            ["omittedSourceUnitCount"] = report.SourceUnitCounts.Sum(count => count.OmittedCount).ToString(CultureInfo.InvariantCulture),
            ["unassessedSourceUnitCount"] = report.SourceUnitCounts.Sum(count => count.UnassessedCount).ToString(CultureInfo.InvariantCulture),
            ["sourceReferenceIssueCount"] = report.SourceReferenceIssues.Count.ToString(CultureInfo.InvariantCulture),
            ["sourceDeclarationIssueCount"] = report.SourceDeclarationIssues.Count.ToString(CultureInfo.InvariantCulture),
            ["sourceMissingReferenceTargetCount"] = report.SourceReferenceIssues.Count(issue => issue.Kind == IWorkSourceReferenceIssueKind.MissingTarget).ToString(CultureInfo.InvariantCulture),
            ["sourceMalformedReferenceCount"] = report.SourceReferenceIssues.Count(issue => issue.Kind == IWorkSourceReferenceIssueKind.MalformedReference).ToString(CultureInfo.InvariantCulture),
            ["sourceRejectedReferenceSetCount"] = report.SourceReferenceIssues.Count(issue => issue.Kind == IWorkSourceReferenceIssueKind.RejectedReferenceSet).ToString(CultureInfo.InvariantCulture),
            ["sourceFormulaCellCount"] = report.FormulaSummary.TotalCount.ToString(CultureInfo.InvariantCulture),
            ["sourceCompleteFormulaExpressionCount"] = report.FormulaSummary.CompleteExpressionCount.ToString(CultureInfo.InvariantCulture),
            ["sourceIncompleteFormulaExpressionCount"] = report.FormulaSummary.IncompleteExpressionCount.ToString(CultureInfo.InvariantCulture),
            ["sourceCompleteFormulaCacheCount"] = report.FormulaSummary.CompleteCacheCount.ToString(CultureInfo.InvariantCulture),
            ["sourcePartialFormulaCacheCount"] = report.FormulaSummary.PartialCacheCount.ToString(CultureInfo.InvariantCulture),
            ["sourceApproximateFormulaCacheCount"] = report.FormulaSummary.ApproximateCacheCount.ToString(CultureInfo.InvariantCulture),
            ["sourceMissingFormulaCacheCount"] = report.FormulaSummary.MissingCacheCount.ToString(CultureInfo.InvariantCulture),
            ["partialEditableReconstruction"] = report.IsPartialEditableReconstruction.ToString(),
            ["visualCoverage"] = report.VisualPreview?.Coverage.ToString() ?? "NotApplicable"
        });
}
