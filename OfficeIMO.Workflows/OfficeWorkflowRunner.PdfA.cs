using System.Globalization;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static OperationArtifact ConvertToPdfA(ValidatedRequest request, List<OfficeWorkflowDiagnostic> diagnostics, CancellationToken token) {
        byte[] input = ReadInput(request.InputPath, request.Limits, token);
        PdfHealthSnapshot before = CreateHealthSnapshot(input, request.PdfLoadOptions, token);
        token.ThrowIfCancellationRequested();
        PdfAConversionResult conversion = PdfDocument.Load(input, request.PdfLoadOptions).ConvertToPdfA(request.PdfAConversion, token);
        token.ThrowIfCancellationRequested();
        PdfAConversionReport report = conversion.Report;
        string profile = report.Target;
        var facts = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["profile"] = profile,
            ["changes"] = report.Changes.Count.ToString(CultureInfo.InvariantCulture),
            ["lossyChanges"] = report.Changes.Count(change => change.LossKind != OfficeConversionLossKind.None).ToString(CultureInfo.InvariantCulture),
            ["violations"] = report.Violations.Count.ToString(CultureInfo.InvariantCulture)
        };
        var evidence = new OfficeWorkflowConversionEvidence(report, facts);
        AddPdfADiagnostics(report, diagnostics);
        byte[] output;
        try {
            output = conversion.ToBytes();
        } catch (PdfAConversionException refusal) {
            throw new WorkflowConversionFailureException(refusal, evidence, diagnosticsAdded: true);
        }

        PdfHealthSnapshot after = CreateHealthSnapshot(output, request.OutputPdfLoadOptions, token);
        bool verified = after.CanRead && !after.HasEncryption && after.PageCount == before.PageCount;
        string summary = report.Changes.Count == 0
            ? profile + " copy created; the source needed no changes."
            : profile + " copy created with " + report.Changes.Count.ToString(CultureInfo.InvariantCulture) + " change(s); review the structured diagnostics.";
        var health = new PdfHealthReport(request.Operation, before, after, summary, verified, facts);
        return new OperationArtifact(output, summary, health, ConversionEvidence: evidence);
    }

    // The library report keeps the object number; workflow diagnostics carry it alongside the location.
    private static void AddPdfADiagnostics(PdfAConversionReport report, List<OfficeWorkflowDiagnostic> diagnostics) {
        foreach (PdfAConversionViolation violation in report.Violations) {
            diagnostics.Add(new OfficeWorkflowDiagnostic("PdfA." + violation.Reason, violation.Message, OfficeWorkflowDiagnosticSeverity.Error, "convert",
                PdfADetails(violation.Location, violation.ObjectNumber, OfficeConversionLossKind.Failure)));
        }
        foreach (PdfAConversionChange change in report.Changes) {
            diagnostics.Add(new OfficeWorkflowDiagnostic("PdfA." + change.Kind, change.Message,
                change.LossKind == OfficeConversionLossKind.None ? OfficeWorkflowDiagnosticSeverity.Information : OfficeWorkflowDiagnosticSeverity.Warning,
                "convert", PdfADetails(change.Location, change.ObjectNumber, change.LossKind)));
        }
    }

    private static Dictionary<string, string> PdfADetails(string location, int? objectNumber, OfficeConversionLossKind lossKind) {
        var details = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["source"] = "PDF/A", ["lossKind"] = lossKind.ToString(), ["location"] = location
        };
        if (objectNumber is int number) details["objectNumber"] = number.ToString(CultureInfo.InvariantCulture);
        return details;
    }

    internal static PdfAConversionOptions? SnapshotPdfAConversion(PdfAConversionOptions? options) {
        if (options is null) return null;
        PdfOutputIntent? intent = options.OutputIntent;
        return new PdfAConversionOptions {
            Profile = options.Profile,
            Attachments = options.Attachments,
            FontSubstitution = options.FontSubstitution,
            OutputIntent = intent is null ? null : new PdfOutputIntent(intent.IccProfile, intent.OutputConditionIdentifier, intent.Policy, intent.Subtype) {
                OutputCondition = intent.OutputCondition, RegistryName = intent.RegistryName, Info = intent.Info
            }
        };
    }
}
