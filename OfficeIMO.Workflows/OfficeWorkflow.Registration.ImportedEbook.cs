using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private static readonly ImportedEbookRoute[] ImportedEbookRoutes = CreateImportedEbookRoutes().ToArray();

    private static IEnumerable<ImportedEbookRoute> CreateImportedEbookRoutes() {
        var routes = new List<ImportedEbookRoute>();
        return routes;
    }

    private static void AddImportedEbookRoutes(List<ImportedEbookRoute> routes, string format,
        Func<Stream, string, long, CancellationToken, OfficeConversionResult<EpubPublication>> import) {
        foreach (string target in new[] { "epub", "pdf", "html", "markdown" })
            routes.Add(new ImportedEbookRoute("ebook-" + format + "-" + target, target, import));
    }

    private sealed class ImportedEbookRoute(string id, string target,
        Func<Stream, string, long, CancellationToken, OfficeConversionResult<EpubPublication>> import)
        : IOfficeConversionRoute, IOfficeConversionArtifactValidation {
        public string Id => id;

        public Task<OfficeConversionReport> ConvertAsync(Stream source, Stream destination,
            OfficeConversionOptions options, CancellationToken cancellationToken = default) {
            var imported = import(source, options.SourceName!, options.MaximumInputBytes, cancellationToken);
            if (options.FailOnLoss) imported.Report.ThrowIfLoss();
            var diagnostics = new List<OfficeDiagnostic>(imported.Report.Diagnostics);
            long itemsWritten = 1;
            if (target == "epub") {
                EpubWriteReport write = imported.Value.WriteStaged(destination,
                    new EpubSaveOptions { MaxOutputBytes = options.MaximumOutputBytes }, cancellationToken);
                diagnostics.AddRange(write.Diagnostics);
            } else {
                EpubDocument book = EpubImportedPublication.Open(imported.Value, cancellationToken);
                if (target == "pdf") {
                    OfficeConversionReport pdf = EpubPdfExport.Write(book, destination, options.Html, false, cancellationToken);
                    diagnostics.AddRange(pdf.Diagnostics);
                    itemsWritten = pdf.ItemsWritten;
                    if (book.Metadata.Any(item => item.Kind == EpubMetadataKind.DublinCore &&
                        item.Name is not ("title" or "creator" or "language" or "identifier")))
                        AddMetadataLoss(diagnostics, "PDF retains the title, primary creator and language from the publication metadata.");
                } else {
                    foreach (EpubDiagnostic item in book.Diagnostics)
                        if (item.Severity != EpubDiagnosticSeverity.Info)
                            diagnostics.Add(new OfficeDiagnostic(item.Code,
                                item.Severity == EpubDiagnosticSeverity.Error ? OfficeDiagnosticSeverity.Error : OfficeDiagnosticSeverity.Warning,
                                item.Message, item.Path, "OfficeIMO.Epub",
                                item.Severity == EpubDiagnosticSeverity.Error ? OfficeConversionLossKind.Failure : OfficeConversionLossKind.Omission));
                    if (target == "html") {
                        EpubImportedPublication.WriteHtml(book, destination, diagnostics, cancellationToken);
                    } else {
                        using var html = new OfficeBoundedMemoryStream(Math.Min(options.MaximumOutputBytes, 128L * 1024 * 1024));
                        EpubImportedPublication.WriteHtml(book, html, diagnostics, cancellationToken);
                        html.Position = 0;
                        using var text = new StreamReader(html, Encoding.UTF8);
                        var markup = HtmlConversionDocument.Parse(text.ReadToEnd(), null, cancellationToken);
                        diagnostics.AddRange(markup.SaveAsMarkdown(destination, cancellationToken: cancellationToken).Diagnostics);
                        AddMetadataLoss(diagnostics, "Markdown carries book content and navigation but omits publication metadata fields.");
                    }
                }
            }
            var report = new OfficeConversionReport(options.SourceName!, options.TargetName!,
                itemsWritten: itemsWritten, diagnostics: diagnostics);
            if (options.FailOnLoss || diagnostics.Any(item => item.LossKind == OfficeConversionLossKind.Failure))
                report.ThrowIfLoss();
            return Task.FromResult(report);
        }

        private static void AddMetadataLoss(List<OfficeDiagnostic> diagnostics, string message) =>
            diagnostics.Add(new OfficeDiagnostic("EBOOK_OUTPUT_METADATA_OMITTED", OfficeDiagnosticSeverity.Warning,
                message, null, "OfficeIMO.Epub", OfficeConversionLossKind.Omission));

        public async Task<bool> TryValidateArtifactAsync(Stream output, OfficeConversionOptions options, CancellationToken token) {
            if (target != "epub") return false;
            output.Position = 0;
            await ValidateManuscriptArtifactAsync(output, options.MaximumOutputBytes, null, token).ConfigureAwait(false);
            output.Position = output.Length;
            return true;
        }
    }
}
