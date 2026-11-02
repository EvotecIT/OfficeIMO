using System;
using System.Linq;
using System.Threading.Tasks;

namespace OfficeIMO.Core.Internal;

/// <summary>Retains completed conversion evidence when a later output operation fails.</summary>
internal static class OfficeConversionFailures {
    internal static OfficeConversionException Output(OfficeConversionReport report, Exception error) {
        var diagnostics = report.Diagnostics.AsEnumerable();
        if (error is OfficeConversionException conversion && !ReferenceEquals(conversion.Report, report)) {
            diagnostics = diagnostics.Concat(conversion.Report is OfficeConversionReport shared ? shared.Diagnostics
                : conversion.Report.FidelityDiagnostics.Select(d => new OfficeDiagnostic(d.Code,
                    OfficeDiagnosticSeverity.Warning, d.Message, d.Location, d.Source, d.LossKind)));
        }
        var failed = new OfficeConversionReport(report.Source, report.Target, itemsWritten: report.ItemsWritten,
            elapsed: report.Elapsed, diagnostics: diagnostics.Concat(new[] {
                new OfficeDiagnostic(OfficeDiagnosticCodes.ConversionOutputFailed, OfficeDiagnosticSeverity.Error,
                    error.Message, format: report.Target, lossKind: OfficeConversionLossKind.Failure)
            }));
        return new OfficeConversionException(error.Message, failed, error);
    }

    internal static void Run(OfficeConversionReport report, Action write) {
        try { write(); }
        catch (Exception error) when (CanRetain(error)) { throw Output(report, error); }
    }

    internal static async Task RunAsync(OfficeConversionReport report, Func<Task> write) {
        try { await write().ConfigureAwait(false); }
        catch (Exception error) when (CanRetain(error)) { throw Output(report, error); }
    }

    private static bool CanRetain(Exception error) => error is not OperationCanceledException
        and not OutOfMemoryException and not StackOverflowException;
}
