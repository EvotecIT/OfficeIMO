namespace OfficeIMO;

internal sealed class OfficeConversionExecutor : IOfficeConversionExecutor {
    public IReadOnlyList<OfficeConversionExecutionRoute> ConversionRoutes => OfficeConverter.ExecutionRoutes;
    public bool SupportsDirectoryPackageInput(string routeId) => OfficeConverter.SupportsDirectoryRoute(routeId);

    public async Task<OfficeConversionExecutionResult> RunAsync(OfficeConversionExecutionRequest request, CancellationToken cancellationToken = default) {
        OfficeConversionOptions options = (request.ConversionOptions ?? new()).Snapshot();
        options.RouteId = request.ConversionRouteId;
        options.OutputProfile = request.OutputProfile;
        options.FileConflictPolicy = request.ConflictPolicy;
        options.MaximumInputBytes = request.Limits.MaximumInputBytes;
        options.MaximumOutputBytes = request.Limits.MaximumOutputBytes;
        options.MaximumXmlCharactersInPart = request.Limits.MaximumXmlCharactersInPart;
        options.PublicationGuard = request.PublicationGuard;
        try {
            OfficeConversionReport report = await OfficeConverter.ConvertFileAsync(request.InputPath,
                request.OutputPath ?? throw new ArgumentException("A conversion destination is required."), options, request.InputStream, cancellationToken).ConfigureAwait(false);
            return new(report.OutputPath, OfficeBatchItemStatus.Completed, "Converted and verified the published artifact.", report.Diagnostics, report);
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            return new(null, OfficeBatchItemStatus.Cancelled, "Conversion was cancelled before publication.", []);
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            OfficeConversionReport? partial = error is OfficeConversionException conversion ? conversion.Report as OfficeConversionReport : null;
            var diagnostics = (partial?.Diagnostics ?? Array.Empty<OfficeDiagnostic>()).ToList();
            diagnostics.Add(new OfficeDiagnostic(OfficeDiagnosticCodes.ConversionFailed, OfficeDiagnosticSeverity.Error, error.Message,
                format: "conversion", lossKind: OfficeConversionLossKind.Failure));
            return new(null, OfficeBatchItemStatus.Failed, error.Message, diagnostics, partial);
        }
    }
}
