namespace OfficeIMO.Email.Store;

internal sealed class EmailStoreStructuralValidationResult {
    internal EmailStoreStructuralValidationResult(bool supported,
        int pagesExamined, int blocksExamined, long bytesExamined,
        int failures, bool wasTruncated,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        Supported = supported;
        PagesExamined = pagesExamined;
        BlocksExamined = blocksExamined;
        BytesExamined = bytesExamined;
        Failures = failures;
        WasTruncated = wasTruncated;
        Diagnostics = diagnostics;
    }

    internal bool Supported { get; }
    internal int PagesExamined { get; }
    internal int BlocksExamined { get; }
    internal long BytesExamined { get; }
    internal int Failures { get; }
    internal bool WasTruncated { get; }
    internal IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    internal static EmailStoreStructuralValidationResult NotSupported() =>
        new EmailStoreStructuralValidationResult(
            supported: false, 0, 0, 0, 0, false,
            new[] {
                new EmailStoreDiagnostic(
                    "EMAIL_STORE_STRUCTURAL_VALIDATION_UNSUPPORTED",
                    "Trailer-level structural validation is currently available for PST and OST sources.",
                    EmailStoreDiagnosticSeverity.Warning,
                    "structure")
            });
}
