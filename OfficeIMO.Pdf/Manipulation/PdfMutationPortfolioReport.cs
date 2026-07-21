namespace OfficeIMO.Pdf;

/// <summary>
/// Input-specific portfolio of mutation plans produced by the single shared mutation planner.
/// </summary>
public sealed class PdfMutationPortfolioReport {
    internal PdfMutationPortfolioReport(PdfDocumentPreflight preflight, IReadOnlyList<PdfMutationPlan> plans) {
        Preflight = preflight;
        Plans = plans;
    }

    /// <summary>Shared preflight snapshot used by every plan.</summary>
    public PdfDocumentPreflight Preflight { get; }

    /// <summary>Requested operations in stable enum order.</summary>
    public IReadOnlyList<PdfMutationPlan> Plans { get; }

    /// <summary>Plans with a proven full-rewrite or append-only path.</summary>
    public IReadOnlyList<PdfMutationPlan> ExecutablePlans => Plans.Where(static plan => plan.CanExecute).ToArray();

    /// <summary>Plans without a proven path for this input.</summary>
    public IReadOnlyList<PdfMutationPlan> BlockedPlans => Plans.Where(static plan => !plan.CanExecute).ToArray();

    /// <summary>True when every requested operation has a proven execution path.</summary>
    public bool CanExecuteAll => Plans.All(static plan => plan.CanExecute);

    /// <summary>Returns the plan for one requested mutation family.</summary>
    public PdfMutationPlan Get(PdfMutationOperation operation) =>
        Plans.FirstOrDefault(plan => plan.Operation == operation) ??
        throw new KeyNotFoundException("The mutation portfolio does not contain " + operation + ".");
}
