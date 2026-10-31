namespace OfficeIMO;

// Internal transport for the durable batch algorithm; public callers use OfficeBatchOptions.
internal interface IOfficeConversionExecutor {
    IReadOnlyList<OfficeConversionExecutionRoute> ConversionRoutes { get; }
    Task<OfficeConversionExecutionResult> RunAsync(OfficeConversionExecutionRequest request, CancellationToken cancellationToken = default);
    bool SupportsDirectoryPackageInput(string routeId);
}

internal sealed record OfficeConversionExecutionRequest {
    internal string InputPath { get; init; } = string.Empty;
    internal string? OutputPath { get; init; }
    internal OfficeConversionExecutionInput? InputStream { get; init; }
    internal string? ConversionRouteId { get; init; }
    internal OfficeConversionOptions? ConversionOptions { get; init; }
    internal OfficeConversionProfile OutputProfile { get; init; }
    internal OfficePublicationPolicy ConflictPolicy { get; init; }
    internal IOfficePublicationGuard? PublicationGuard { get; init; }
    internal OfficeOperationLimits Limits { get; init; } = new();
}

internal sealed record OfficeConversionExecutionInput(string Name, Func<CancellationToken, Task<Stream>> OpenRead, string? ExpectedSha256,
    IOfficePublicationGuard? SourcePublicationGuard = null);

internal interface IOfficeConversionDirectoryOwner {
    OfficeConversionExecutionInput CaptureDirectoryInput(string path, OfficeConversionOptions options, OfficeOperationLimits limits);
}
internal sealed record OfficeConversionExecutionResult(string? OutputPath, OfficeBatchItemStatus Status, string Summary,
    IReadOnlyList<OfficeDiagnostic> Diagnostics, OfficeConversionReport? Report = null) {
    internal bool Succeeded => Status == OfficeBatchItemStatus.Completed;
}

internal sealed class OfficeConversionExecutionRoute(OfficeConversionCapability capability, bool checkpointable) {
    internal string Id => capability.Id;
    internal string Source => capability.Source;
    internal string Target => capability.Target;
    internal string TargetExtension => capability.TargetExtension;
    internal IReadOnlyList<string> SourceExtensions => capability.SourceExtensions;
    internal bool CanExecute => true;
    internal bool Checkpointable => checkpointable;
    internal IReadOnlyList<OfficeConversionProfile> SupportedOutputProfiles => Id is "doc-pdf" or "docx-pdf" or "xls-pdf" or "xlsb-pdf" or "xlsx-pdf" or "ppt-pdf" or "pptx-pdf"
        ? EnumCompat.GetValues<OfficeConversionProfile>() : [OfficeConversionProfile.Faithful];
}

// First-party owners provide a stable fingerprint using their native source-generated metadata.
// Runtime third-party registrations remain ordinary-batch only.
internal interface IOfficeConversionCheckpointOwner {
    string CaptureConfiguration(OfficeConversionOptions options);
}

// A registered first-party owner supplies its output encryption context for reopen validation.
internal interface IOfficeConversionOutputValidation {
    OfficeIMO.Pdf.PdfCreateOptions? GetOutputCreationOptions(OfficeConversionOptions options);
    OfficeIMO.Pdf.PdfLoadOptions? CreateOutputPdfLoadOptions(OfficeConversionOptions options);
}
