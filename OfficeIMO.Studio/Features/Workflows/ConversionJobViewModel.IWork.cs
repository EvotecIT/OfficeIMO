using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.IWork;
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class ConversionJobViewModel {
    public bool SupportsIWorkOptions => Route.Route.Id is "pages-docx" or "numbers-xlsx" or "keynote-pptx";
    public bool IsNumbersImport => Route.Route.Id == "numbers-xlsx";
    public IReadOnlyList<IWorkConversionMode> IWorkModes { get; } = Enum.GetValues<IWorkConversionMode>();

    [ObservableProperty, NotifyPropertyChangedFor(nameof(Fidelity))]
    private IWorkConversionMode _iWorkMode = IWorkConversionMode.Auto;
    [ObservableProperty] private bool _allowPartialEditableReconstruction;
    [ObservableProperty] private bool _allowIncompleteVisualPreview;
    [ObservableProperty] private bool _normalizeWorksheetNames;

    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasConversionEvidence)), NotifyPropertyChangedFor(nameof(ConversionEvidenceSummary))]
    private OfficeWorkflowConversionEvidence? _conversionEvidence;

    public bool HasConversionEvidence => ConversionEvidence is not null;
    public string SourceFingerprint => Diagnostics.FirstOrDefault(diagnostic => diagnostic.Code == "SourceSnapshot")?.Details
        .GetValueOrDefault("sha256") ?? string.Empty;

    public string ConversionEvidenceSummary {
        get {
            if (ConversionEvidence is null) return string.Empty;
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
            string counts = _localizer.FormatOrDefault("Conversion.IWork.ReconstructionCounts",
                "{0} reconstructed items · {1} records without a field-level fidelity assessment",
                facts.GetValueOrDefault("reconstructedItemCount") ?? "?", facts.GetValueOrDefault("unassessedRecordCount") ?? "?");
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
