namespace OfficeIMO.Pdf;

/// <summary>Thrown when the shared PDF mutation planner cannot prove an execution path.</summary>
public sealed class PdfMutationBlockedException : NotSupportedException {
    internal PdfMutationBlockedException(PdfMutationPlan plan)
        : base(BuildMessage(plan)) {
        Plan = plan;
    }

    /// <summary>Full mutation decision, including structures, permissions, proofs, blockers, warnings, and diagnostics.</summary>
    public PdfMutationPlan Plan { get; }

    /// <summary>Turns a failed required output proof into a blocked decision without exposing document values.</summary>
    internal static PdfMutationBlockedException FailedProof(PdfMutationPlan plan, PdfMutationProof proof, string? diagnostic = null) => new(new PdfMutationPlan(
        plan.Operation, plan.ExecutionPreference, PdfMutationExecutionMode.Blocked,
        plan.Preflight, plan.AppendOnlyReport, fullRewriteAvailable: false, appendOnlyAvailable: false,
        plan.AffectedStructures, plan.CapabilityRecords, plan.PermissionChecks, plan.RequiredProofs,
        new[] { "FullRewrite.ProofFailed." + proof }, plan.Warnings,
        new[] { diagnostic ?? "The required PDF " + proof + " proof could not be established." }));

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
