namespace OfficeIMO.IWork.Internal;

internal static class IWorkProjectionDiagnostics {
    internal static IWorkDiagnostic SemanticProjectionSkipped { get; } = new(
        IWorkDiagnosticSeverity.Information,
        "IWORK_SEMANTIC_PROJECTION_SKIPPED",
        "Semantic reconstruction was skipped because visual-only import was requested.");
}
