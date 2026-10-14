using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.IWork;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionJobViewModel {
    internal static bool IsEditableIWorkRoute(string id) => id is "pages-docx" or "numbers-xlsx" or "keynote-pptx";
    public bool SupportsIWorkOptions => IsEditableIWorkRoute(Route.Route.Id);
    public bool IsNumbersImport => Route.Route.Id == "numbers-xlsx";
    public IReadOnlyList<IWorkConversionMode> IWorkModes { get; } = Enum.GetValues<IWorkConversionMode>();

    [ObservableProperty, NotifyPropertyChangedFor(nameof(Fidelity))]
    private IWorkConversionMode _iWorkMode = IWorkConversionMode.Auto;
    [ObservableProperty] private bool _allowPartialEditableReconstruction;
    [ObservableProperty] private bool _allowIncompleteVisualPreview;
    [ObservableProperty] private bool _normalizeWorksheetNames;

    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasConversionEvidence)), NotifyPropertyChangedFor(nameof(ConversionEvidenceSummary)), NotifyPropertyChangedFor(nameof(HasConversionEvidenceSummary))]
    private OfficeWorkflowConversionEvidence? _conversionEvidence;

    public bool HasConversionEvidence => ConversionEvidence is not null;

    /// <summary>True for routes that report how the result was reconstructed; other routes show only the source snapshot.</summary>
    public bool HasConversionEvidenceSummary => ConversionEvidenceSummary.Length > 0;
    public string SourceFingerprint => Diagnostics.FirstOrDefault(diagnostic => diagnostic.Code == "SourceSnapshot")?.Details
        .GetValueOrDefault("sha256") ?? string.Empty;

    public string ConversionEvidenceSummary {
        get {
            // Only reconstruction routes describe a projection; without one there is nothing to summarize.
            if (ConversionEvidence is null || !ConversionEvidence.Facts.ContainsKey("projectionKind")) return string.Empty;
            var facts = ConversionEvidence.Facts;
            string projection = facts.GetValueOrDefault("projectionKind") == "VisualFallback"
                ? _localizer.GetOrDefault("Conversion.IWork.VisualResult", "Embedded visual preview")
                : _localizer.GetOrDefault("Conversion.IWork.EditableResult", "Editable reconstruction");
            string coverage = facts.GetValueOrDefault("visualCoverage") switch {
                "FullDocument" => _localizer.GetOrDefault("Conversion.IWork.FullCoverage", "Known full-document preview coverage"),
                "NotApplicable" => string.Empty,
                _ => _localizer.GetOrDefault("Conversion.IWork.IncompleteCoverage", "Preview coverage is incomplete or unknown")
            };
            string partial = facts.GetValueOrDefault("partialEditableReconstruction") == "True"
                ? _localizer.GetOrDefault("Conversion.IWork.PartialResult", "Partial editable reconstruction was accepted") : string.Empty;
            string counts = facts.TryGetValue("reconstructedItemCount", out string? reconstructed) && facts.TryGetValue("unassessedRecordCount", out string? unassessed)
                ? _localizer.FormatOrDefault("Conversion.IWork.ReconstructionCounts",
                    "{0} reconstructed items · {1} records without a field-level fidelity assessment", reconstructed, unassessed)
                : string.Empty;
            string sourceUnits = facts.ContainsKey("sourceUnitCount")
                ? _localizer.FormatOrDefault("Conversion.IWork.SourceUnitCounts",
                    "{0} identified source units · {1} reconstructed · {2} omitted · {3} unassessed",
                    facts.GetValueOrDefault("sourceUnitCount") ?? "?", facts.GetValueOrDefault("reconstructedSourceUnitCount") ?? "?",
                    facts.GetValueOrDefault("omittedSourceUnitCount") ?? "?", facts.GetValueOrDefault("unassessedSourceUnitCount") ?? "?")
                : string.Empty;
            string formulas = facts.TryGetValue("sourceFormulaCellCount", out string? formulaCount) && formulaCount != "0"
                ? _localizer.FormatOrDefault("Conversion.IWork.FormulaCounts",
                    "Source formulas: {0} complete · {1} incomplete · {6} unassessed. Recovered caches: {2} complete · {3} partial · {4} approximate · {5} missing · {7} unassessed",
                    facts.GetValueOrDefault("sourceCompleteFormulaExpressionCount") ?? "?", facts.GetValueOrDefault("sourceIncompleteFormulaExpressionCount") ?? "?",
                    facts.GetValueOrDefault("sourceCompleteFormulaCacheCount") ?? "?", facts.GetValueOrDefault("sourcePartialFormulaCacheCount") ?? "?",
                    facts.GetValueOrDefault("sourceApproximateFormulaCacheCount") ?? "?", facts.GetValueOrDefault("sourceMissingFormulaCacheCount") ?? "?",
                    facts.GetValueOrDefault("sourceUnassessedFormulaExpressionCount") ?? "?", facts.GetValueOrDefault("sourceUnassessedFormulaCacheCount") ?? "?")
                : string.Empty;
            return string.Join(Environment.NewLine, new[] { projection, coverage, partial, sourceUnits, formulas, counts }.Where(text => text.Length > 0));
        }
    }

    internal IOfficeWorkflowConversionSettings? CreateRegisteredConversionSettings() => SupportsIWorkOptions
        ? new IWorkWorkflowSettings {
            ConversionOptions = new IWorkConversionOptions {
                Mode = IWorkMode,
                AllowPartialEditableReconstruction = AllowPartialEditableReconstruction,
                RequireCompleteVisualCoverage = !AllowIncompleteVisualPreview,
                NormalizeWorksheetNames = IsNumbersImport && NormalizeWorksheetNames
            }
        } : null;
}
