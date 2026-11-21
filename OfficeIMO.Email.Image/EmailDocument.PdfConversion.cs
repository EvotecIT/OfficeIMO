using System.Diagnostics;
using OfficeIMO.Html.Pdf;

namespace OfficeIMO.Email;

public static partial class EmailImageExportExtensions {
    /// <summary>Renders the message with its header block, body and attachment list to a caller-owned PDF stream.</summary>
    /// <remarks>The body uses the same selection, sanitization and inline-resource policy as image export. Body
    /// projection findings, such as an RTF-only body, are reported with the HTML-to-PDF findings.</remarks>
    public static async Task<OfficeConversionReport> SaveAsPdfAsync(this EmailDocument document, Stream destination,
        EmailPdfExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(destination);
        cancellationToken.ThrowIfCancellationRequested();
        var timer = Stopwatch.StartNew();
        EmailPdfExportOptions effective = options?.Clone() ?? new EmailPdfExportOptions();
        HtmlPdfExportOptions render = effective.Pdf;
        EmailMessageHtml message = PrepareMessage(document, new EmailMessageLayout(
            effective.IncludeMessageHeaders, effective.IncludeAttachmentList, Print: true, effective.PreferHtmlBody,
            effective.IncludeInlineResources, effective.RemoteResourcePolicy, effective.MaxInlineResourceCount,
            effective.MaxTotalInlineResourceBytes), render);
        OfficeDiagnostic[] bodyFindings = message.Diagnostics.Select(MapBodyPdfDiagnostic).ToArray();
        if (render.FailOnLoss && bodyFindings.Any(finding => finding.LossKind != OfficeConversionLossKind.None)) {
            throw new OfficeConversionException("The email body projection reported fidelity loss.",
                new OfficeConversionReport("Email", "PDF", elapsed: timer.Elapsed, diagnostics: bodyFindings));
        }
        try {
            OfficeConversionReport rendered = await message.Document.SaveAsPdfAsync(destination, render, cancellationToken).ConfigureAwait(false);
            return new OfficeConversionReport("Email", rendered.Target, rendered.OutputPath, rendered.ItemsWritten, timer.Elapsed,
                bodyFindings.Concat(rendered.Diagnostics));
        } catch (OfficeConversionException error) when (error.Report is { } report) {
            throw new OfficeConversionException(error.Message, new OfficeConversionReport("Email", report.Target, report.OutputPath,
                report.ItemsWritten, timer.Elapsed, bodyFindings.Concat(report.Diagnostics)), error);
        }
    }

    private static OfficeDiagnostic MapBodyPdfDiagnostic(EmailDiagnostic diagnostic) => new(
        diagnostic.Code,
        diagnostic.Severity == EmailDiagnosticSeverity.Error ? OfficeDiagnosticSeverity.Error : OfficeDiagnosticSeverity.Warning,
        diagnostic.Message,
        "Email body",
        "Email",
        diagnostic.Code == "EMAIL_BODY_RTF_PROJECTED" ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.Omission);
}
