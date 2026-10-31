using OfficeIMO.Core.Internal;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private class WorkflowRoute(string id, OfficeWorkflowConversionRegistration? registration = null) : IOfficeConversionRoute, IOfficeConversionOutputValidation {
        protected OfficeWorkflowConversionRegistration? Registration { get; } = registration;
        public string Id => id;

        public async Task<OfficeConversionReport> ConvertAsync(Stream source, Stream destination, OfficeConversionOptions options,
            CancellationToken cancellationToken = default) {
            if (Registration == null && options.RegisteredSettings != null && options.RegisteredSettings is not OfficeWorkflowRouteSettings)
                throw new ArgumentException("This route requires OfficeWorkflowRouteSettings.", nameof(options));
            OfficeConversionOptions capturedOptions = options.Clone();
            // The front door owns post-conversion compression and final encryption.
            capturedOptions.CompressPdfOutput = false;
            string directory = OfficeTemporaryDirectory.Create("officeimo-workflow-conversion-");
            try {
                string inputPath = Path.Combine(directory, Path.GetFileName(options.SourceName!));
                string targetDirectory = Directory.CreateDirectory(Path.Combine(directory, "target")).FullName;
                string outputPath = Path.Combine(targetDirectory, Path.GetFileName(options.TargetName!));
                byte[] input = OfficeConversionInputReader.ReadAllBytes(source, options.SourceName!, options.MaximumInputBytes, cancellationToken);
                using (var captured = OfficeTemporaryFile.CreateAtPath(inputPath, 81920, FileOptions.Asynchronous))
                    await captured.WriteAsync(input, 0, input.Length, cancellationToken).ConfigureAwait(false);
                var request = new OfficeWorkflowRequest {
                    Operation = OfficeWorkflowOperation.Convert, InputPath = inputPath, OutputPath = outputPath,
                    InputStream = new OfficeWorkflowStreamInput(Path.GetFileName(inputPath), token => {
                        token.ThrowIfCancellationRequested(); return Task.FromResult<Stream>(new MemoryStream(input, false));
                    }),
                    ConversionRouteId = Id, ConversionOptions = capturedOptions, OutputProfile = options.OutputProfile,
                    RegisteredConversionSettings = Registration?.SnapshotSettings(options.RegisteredSettings),
                    ConflictPolicy = OfficePublicationPolicy.Fail,
                    Limits = new OfficeOperationLimits { MaximumInputBytes = options.MaximumInputBytes,
                        MaximumOutputBytes = options.MaximumOutputBytes, MaximumXmlCharactersInPart = options.MaximumXmlCharactersInPart }
                };
                var runner = Registration == null ? new OfficeWorkflowRunner()
                    : new OfficeWorkflowRunner(null, null, conversions: new[] { Registration });
                OfficeWorkflowResult result = await runner.RunAsync(request, cancellationToken: cancellationToken).ConfigureAwait(false);
                OfficeConversionReport report = PdfConversionEvidence.CreateReport(options.SourceName!,
                    result.ConversionEvidence == null ? Array.Empty<IOfficeConversionReport>() : new IOfficeConversionReport[] { result.ConversionEvidence },
                    elapsed: result.Duration, itemsWritten: result.Succeeded ? 1 : 0, target: options.TargetName!);
                var additional = result.Diagnostics.Where(d => d.Details.ContainsKey("lossKind") &&
                    !report.Diagnostics.Any(existing => existing.Code == d.Code && existing.Message == d.Message)).Select(d =>
                    new OfficeDiagnostic(d.Code, d.Severity == OfficeWorkflowDiagnosticSeverity.Error ? OfficeDiagnosticSeverity.Error :
                        d.Severity == OfficeWorkflowDiagnosticSeverity.Information ? OfficeDiagnosticSeverity.Information : OfficeDiagnosticSeverity.Warning,
                        d.Message, d.Details.TryGetValue("location", out string? location) ? location : null,
                        d.Details.TryGetValue("source", out string? producer) ? producer : Id,
                        Enum.TryParse(d.Details["lossKind"], out OfficeConversionLossKind loss) ? loss : OfficeConversionLossKind.None));
                report = new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten,
                    elapsed: report.Elapsed, diagnostics: report.Diagnostics.Concat(additional));
                if (result.Succeeded && result.Diagnostics.Any(finding => finding.Code == global::OfficeIMO.OfficeDiagnosticCodes.ConversionOutputReopened))
                    report = new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten,
                        elapsed: report.Elapsed, diagnostics: report.Diagnostics.Concat(new[] {
                            new OfficeDiagnostic(OfficeDiagnosticCodes.ConversionOutputReopened, OfficeDiagnosticSeverity.Information,
                                "The conversion output was reopened by its native format owner.", format: Id)
                        }));
                if (result.Status == OfficeWorkflowStatus.Cancelled) throw new OperationCanceledException(cancellationToken);
                if (!result.Succeeded) {
                    string code = result.FailureKind switch {
                        OfficeWorkflowFailureKind.ValidationFailed => OfficeDiagnosticCodes.ConversionRequestInvalid,
                        OfficeWorkflowFailureKind.InputNotFound => OfficeDiagnosticCodes.ConversionInputNotFound,
                        OfficeWorkflowFailureKind.UnsupportedInput => OfficeDiagnosticCodes.ConversionInputUnsupported,
                        OfficeWorkflowFailureKind.OutputFailed => OfficeDiagnosticCodes.ConversionOutputFailed,
                        _ => OfficeDiagnosticCodes.ConversionFailed
                    };
                    report = new OfficeConversionReport(report.Source, report.Target, elapsed: report.Elapsed,
                        diagnostics: report.Diagnostics.Concat(new[] {
                            new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Error, result.Summary,
                                format: Id, lossKind: OfficeConversionLossKind.Failure)
                        }));
                    throw new OfficeConversionException(result.Summary, report);
                }
                if (options.FailOnLoss) report.ThrowIfLoss();
                using var output = OfficeConversionInputReader.OpenLocalRead(outputPath);
                if (output.Length > options.MaximumOutputBytes) throw new InvalidDataException("The workflow artifact exceeds its output limit.");
                await output.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
                return report;
            } finally { Directory.Delete(directory, true); }
        }

        public PdfCreateOptions? GetOutputCreationOptions(OfficeConversionOptions options) {
            if (options.RegisteredSettings is not OfficeWorkflowRouteSettings settings) return null;
            return Id switch {
                "rtf-pdf" => settings.Rtf?.PdfOptions, "odg-pdf" => settings.Draw?.PdfOptions,
                "visio-pdf" => settings.Visio?.PdfOptions, "xps-pdf" => settings.Xps?.PdfOptions,
                "djvu-pdf" => settings.DjVu?.PdfOptions, "publisher-pdf" => settings.PublisherPdf,
                "chm-pdf" => settings.ChmHtml?.PdfOptions, _ => null
            };
        }

        public PdfLoadOptions? CreateOutputPdfLoadOptions(OfficeConversionOptions options) {
            PdfCreateOptions? creation = GetOutputCreationOptions(options);
            if (creation?.Encryption == null) return null;
            var reading = new PdfLoadOptions { Password = creation.Encryption.OwnerPassword ?? creation.Encryption.UserPassword };
            return creation.Encryption.AesCryptographyProvider is { } cryptography
                ? PdfLoadOptions.WithAesCryptographyProvider(reading, cryptography) : reading;
        }
    }
}
