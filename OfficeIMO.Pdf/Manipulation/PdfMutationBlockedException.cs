namespace OfficeIMO.Pdf;

/// <summary>Thrown when the shared PDF mutation planner cannot prove an execution path.</summary>
public sealed class PdfMutationBlockedException : NotSupportedException {
    internal PdfMutationBlockedException(PdfMutationPlan plan)
        : base(BuildMessage(plan)) {
        Plan = plan;
    }

    /// <summary>Full mutation decision, including structures, permissions, proofs, blockers, warnings, and diagnostics.</summary>
    public PdfMutationPlan Plan { get; }

    private static string BuildMessage(PdfMutationPlan plan) {
        if (plan.Operation == PdfMutationOperation.Optimize && plan.Preflight.Probe.Security.HasSignatures) {
            return "Optimization of signed PDF files is refused because a full rewrite would invalidate their signatures. Optimize before signing. " + plan.Summary;
        }
        if (plan.Diagnostics.Count == 0) {
            return plan.Summary;
        }

        return plan.Summary + " " + string.Join(" ", plan.Diagnostics);
    }
}
