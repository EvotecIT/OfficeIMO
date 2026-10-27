namespace OfficeIMO.Pdf;

/// <summary>Freezes internal stage evidence at the public conversion boundary.</summary>
internal static partial class PdfConversionEvidence {
    internal static PdfDocument CompleteModel(PdfDocumentConversionResult conversion, string source, bool failOnLoss,
        System.Diagnostics.Stopwatch timer, System.Threading.CancellationToken token, out OfficeConversionReport report) {
        try {
            byte[] bytes = conversion.ToBytes(token);
            PdfLoadOptions reading = PdfLoadOptions.WithMinimumInputBytes(conversion.Value.ReadOptions, bytes.LongLength);
            using PdfDocument model = PdfDocument.LoadOwned(bytes, reading, token);
            int pageCount = model.Pages.Count;
            report = CreateReport(source, conversion.ConversionReports, elapsed: timer.Elapsed, itemsWritten: pageCount);
            if (failOnLoss) report.ThrowIfLoss();
            return conversion.Value;
        } catch (Exception error) when (error is not OperationCanceledException and not OfficeConversionException and not OutOfMemoryException and not StackOverflowException) {
            conversion.Value.Dispose();
            throw OutputFailure(conversion, source, error);
        } catch {
            conversion.Value.Dispose();
            throw;
        }
    }

    internal static OfficeConversionReport CreateReport(string source, IEnumerable<PdfConversionWarning> warnings,
        string? outputPath = null, TimeSpan elapsed = default, long itemsWritten = 0, string target = "Pdf") =>
        new OfficeConversionReport(source, target, outputPath, itemsWritten, elapsed, warnings.Select(warning => new OfficeDiagnostic(
            warning.Code, warning.Severity switch {
                PdfConversionWarningSeverity.Information => OfficeDiagnosticSeverity.Information,
                PdfConversionWarningSeverity.Error => OfficeDiagnosticSeverity.Error,
                _ => OfficeDiagnosticSeverity.Warning
            }, warning.Message, warning.Source, warning.Converter, warning.LossKind)));

    internal static OfficeConversionReport CreateReport(string source, IEnumerable<IOfficeConversionReport> stages,
        string? outputPath = null, TimeSpan elapsed = default, long itemsWritten = 0, string target = "Pdf") {
        var diagnostics = new List<OfficeDiagnostic>();
        foreach (IOfficeConversionReport stage in stages) {
            if (stage is OfficeConversionReport shared) {
                diagnostics.AddRange(shared.Diagnostics);
            } else if (stage is PdfConversionReport pdf) {
                foreach (PdfConversionWarning warning in pdf.Warnings) {
                    OfficeDiagnosticSeverity severity = warning.Severity switch {
                        PdfConversionWarningSeverity.Information => OfficeDiagnosticSeverity.Information,
                        PdfConversionWarningSeverity.Error => OfficeDiagnosticSeverity.Error,
                        _ => OfficeDiagnosticSeverity.Warning
                    };
                    diagnostics.Add(new OfficeDiagnostic(warning.Code, severity, warning.Message,
                        warning.Source, warning.Converter, warning.LossKind));
                }
            } else {
                foreach (OfficeConversionFidelityDiagnostic diagnostic in stage.FidelityDiagnostics) {
                    diagnostics.Add(new OfficeDiagnostic(diagnostic.Code, OfficeDiagnosticSeverity.Warning,
                        diagnostic.Message, diagnostic.Location, diagnostic.Source, diagnostic.LossKind));
                }
            }
        }
        return new OfficeConversionReport(source, target, outputPath, itemsWritten, elapsed, diagnostics);
    }
}
