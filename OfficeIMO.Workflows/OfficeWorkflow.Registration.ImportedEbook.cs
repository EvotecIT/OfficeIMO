using System.Text;
using OfficeIMO.Comic;
using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private static readonly ImportedEbookRoute[] ImportedEbookRoutes = CreateImportedEbookRoutes().ToArray();

    private static IEnumerable<ImportedEbookRoute> CreateImportedEbookRoutes() {
        var routes = new List<ImportedEbookRoute>();
        AddImportedEbookRoutes(routes, "cbz", ComicReader.Import);
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
        // Up to 120 MiB of comic images can expand to 160 MiB of base64 before adding markup.
        private const int MaximumWorkingHtmlBytes = 192 * 1024 * 1024;

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
                    OfficeConversionReport pdf;
                    try {
                        pdf = EpubPdfExport.Write(book, destination, options.Html, false, cancellationToken);
                    } catch (OfficeConversionException error) {
                        diagnostics.AddRange(error.Report.Diagnostics);
                        var combined = new OfficeConversionReport(options.SourceName!, options.TargetName!,
                            itemsWritten: error.Report.ItemsWritten, diagnostics: diagnostics);
                        throw new OfficeConversionException(error.Message, combined, error);
                    }
                    diagnostics.AddRange(pdf.Diagnostics);
                    itemsWritten = pdf.ItemsWritten;
                    if (EpubImportedPublication.MetadataFields(book).Any(item =>
                        item.Name is not ("dc.title" or "dc.creator" or "dc.language") &&
                        !(item.Name == "dc.identifier" && imported.Value.HasGeneratedIdentifier &&
                          item.Value == imported.Value.Identifier)) ||
                        book.Metadata.Count(item => item.Name == "creator") > 1)
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
                        // The converter's output limit applies to the final Markdown, which can be smaller than this HTML.
                        using var html = new OfficeBoundedMemoryStream(MaximumWorkingHtmlBytes);
                        EpubImportedPublication.WriteHtml(book, html, diagnostics, cancellationToken, markdownAnchors: true);
                        html.Position = 0;
                        using var text = new StreamReader(html, Encoding.UTF8);
                        // The static projection can generate extra nodes for navigation and anchor targets.
                        var limits = HtmlConversionLimits.UntrustedDefaults;
                        limits.MaxInputCharacters = MaximumWorkingHtmlBytes;
                        limits.MaxHtmlNodes = 1_000_000;
                        var markup = HtmlConversionDocument.Parse(text.ReadToEnd(),
                            new HtmlConversionDocumentOptions { Limits = limits }, cancellationToken);
                        var markdownOptions = new HtmlMarkdownExportOptions();
                        markdownOptions.PassThroughTags.Add("a");
                        using MarkdownDocument markdown = markup.ToMarkdown(markdownOptions, out OfficeConversionReport converted, cancellationToken);
                        var metadata = EpubImportedPublication.MetadataFields(book)
                            .GroupBy(item => item.Name, StringComparer.Ordinal)
                            .ToDictionary(group => group.Key, group => group.Count() == 1
                                ? (object)group.First().Value : group.Select(item => item.Value).ToArray(), StringComparer.Ordinal);
                        if (metadata.Count != 0) markdown.FrontMatter(metadata);
                        diagnostics.AddRange(converted.Diagnostics);
                        OfficeStreamWriter.CopyBytes(destination, markdown.ToBytes(cancellationToken: cancellationToken), cancellationToken);
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
