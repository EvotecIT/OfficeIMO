using System.Globalization;
using OfficeIMO.Excel.IWork;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint.IWork;
using OfficeIMO.Word.IWork;

namespace OfficeIMO.Workflows.IWork;

/// <summary>Opt-in Apple iWork conversion using the bounded source reader and existing Office destination owners.</summary>
public static class IWorkWorkflow {
    /// <summary>Creates a runner with Pages-to-Word, Numbers-to-Excel, and Keynote-to-PowerPoint routes.</summary>
    /// <remarks>Defaults reject partial editable reconstruction and visual previews without known full-document coverage. Inputs are ZIP files or provider ZIP streams; directory bundles require a separate package snapshot contract.</remarks>
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
            }),
            OfficeWorkflowConversionRegistration.Create<IWorkWorkflowSettings>("numbers-xlsx", (input, output, limits, settings, token) => {
                using NumbersToExcelResult result = ExcelIWorkConverter.ConvertNumbersToExcelResult(input, Bound(settings?.ReadOptions ?? reading, limits), settings?.ConversionOptions ?? conversion, token);
                result.Value.SaveAsync(output, token).GetAwaiter().GetResult();
                return Evidence(result.Report);
            }),
            OfficeWorkflowConversionRegistration.Create<IWorkWorkflowSettings>("keynote-pptx", (input, output, limits, settings, token) => {
                using KeynoteToPowerPointResult result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(input, Bound(settings?.ReadOptions ?? reading, limits), settings?.ConversionOptions ?? conversion, token);
                result.Value.SaveAsync(output, token).GetAwaiter().GetResult();
                return Evidence(result.Report);
            })
        });
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
            ["partialEditableReconstruction"] = report.IsPartialEditableReconstruction.ToString(),
            ["visualCoverage"] = report.VisualPreview?.Coverage.ToString() ?? "NotApplicable"
        });
}
