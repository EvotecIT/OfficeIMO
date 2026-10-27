using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO;

public static partial class OfficeConverter {
    internal static WordMarkdownImportOptions ConfigureMarkdownHtml(WordMarkdownImportOptions? options, CancellationToken token) {
        WordMarkdownImportOptions configured = WordMarkdownConverterExtensions.CopyOptions(options ?? new());
        configured.RawHtmlRenderer = (text, destination, sink) => {
            token.ThrowIfCancellationRequested();
            var html = HtmlConversionDocument.Parse(text, null, token);
            var import = new WordHtmlImportOptions();
            destination.AddHtmlToBody(html, import);
            if (import.ConversionReport is { } report) {
                foreach (HtmlDiagnostic finding in report.Diagnostics)
                    sink?.Invoke(new WordMarkdownConversionDiagnostic(finding.Code, finding.Message, finding.LossKind,
                        finding.Severity switch {
                            HtmlDiagnosticSeverity.Info => OfficeDiagnosticSeverity.Information,
                            HtmlDiagnosticSeverity.Error => OfficeDiagnosticSeverity.Error,
                            _ => OfficeDiagnosticSeverity.Warning
                        }, finding.Source ?? finding.Provenance.SourceAddress, finding.Component));
            }
            return true;
        };
        configured.RawHtmlParser = (text, sink) => {
            token.ThrowIfCancellationRequested();
            var html = HtmlConversionDocument.Parse(text, null, token);
            MarkdownDocument markdown = html.ToMarkdown(null, out OfficeConversionReport report, token);
            foreach (OfficeDiagnostic finding in report.Diagnostics)
                sink?.Invoke(new WordMarkdownConversionDiagnostic(finding.Code, finding.Message, finding.LossKind,
                    finding.Severity, finding.Location, finding.Format));
            return markdown;
        };
        return configured;
    }
}
