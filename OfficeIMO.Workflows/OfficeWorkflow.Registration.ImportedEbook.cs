using System.Text;
using OfficeIMO.Comic;
using OfficeIMO.Core.Internal;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Fb2;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Mobi;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private static readonly ImportedEbookRoute[] ImportedEbookRoutes = CreateImportedEbookRoutes().ToArray();

    private static IEnumerable<ImportedEbookRoute> CreateImportedEbookRoutes() {
        var routes = new List<ImportedEbookRoute>();
        AddImportedEbookRoutes(routes, "cbz", ComicReader.Import, maximumWorkingHtmlBytes: 192 * 1024 * 1024);
        AddImportedEbookRoutes(routes, "fb2", Fb2Reader.Import, maximumWorkingHtmlBytes: 128 * 1024 * 1024,
            maximumHtmlNodes: EpubImportedPublication.MaximumHtmlNodes);
        AddImportedEbookRoutes(routes, "mobi", MobiReader.Import, maximumWorkingHtmlBytes: 128 * 1024 * 1024,
            maximumHtmlNodes: EpubImportedPublication.MaximumHtmlNodes);
        return routes;
    }

    private static void AddImportedEbookRoutes(List<ImportedEbookRoute> routes, string format,
        Func<Stream, string, long, CancellationToken, OfficeConversionResult<EpubPublication>> import,
        int maximumWorkingHtmlBytes, int? maximumHtmlNodes = null) {
        foreach (string target in new[] { "epub", "pdf", "html", "markdown" })
            routes.Add(new ImportedEbookRoute("ebook-" + format + "-" + target, target, import, maximumWorkingHtmlBytes, maximumHtmlNodes));
    }

    private sealed class ImportedEbookRoute(string id, string target,
        Func<Stream, string, long, CancellationToken, OfficeConversionResult<EpubPublication>> import,
        int maximumWorkingHtmlBytes, int? maximumHtmlNodes)
        : IOfficeConversionRoute, IOfficeConversionArtifactValidation, IOfficeConversionOutputValidation {
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
                if (imported.Value.PageProgressionDirection is "ltr" or "rtl")
                    diagnostics.Add(new OfficeDiagnostic("EBOOK_PAGE_PROGRESSION_OMITTED", OfficeDiagnosticSeverity.Warning,
                        "The target does not carry the ebook's page progression direction.", null, "OfficeIMO.Epub",
                        OfficeConversionLossKind.Omission));
                EpubDocument book = EpubImportedPublication.Open(imported.Value, cancellationToken);
                if (target == "pdf") {
                    OfficeConversionReport pdf;
                    try {
                        pdf = EpubPdfExport.Write(book, destination, options.Html, false, cancellationToken,
                            importedHtmlNodeLimit: maximumHtmlNodes);
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
                        HtmlConversionDocument markup = CreateMarkdownProjection(book, options, diagnostics, cancellationToken);
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

        private HtmlConversionDocument CreateMarkdownProjection(EpubDocument book, OfficeConversionOptions options,
            List<OfficeDiagnostic> diagnostics, CancellationToken token) {
            // The caller's output limit applies to the final Markdown. The combined HTML has its own safety budget.
            using var html = new OfficeBoundedMemoryStream(maximumWorkingHtmlBytes);
            try {
                EpubImportedPublication.WriteHtml(book, html, diagnostics, token, markdownAnchors: true);
                html.Position = 0;
                using var text = new StreamReader(html, Encoding.UTF8);
                return HtmlConversionDocument.Parse(text.ReadToEnd(), ProjectionHtmlOptions(maximumWorkingHtmlBytes), token);
            } catch (HtmlDomLimitException error) {
                throw Reject(error.Code, "The combined ebook HTML exceeds the Markdown limit (" + error.LimitSource + ").", error);
            } catch (InvalidDataException error) when (OfficeOutputLimit.Is(error)) {
                throw Reject("EBOOK_MARKDOWN_WORKING_HTML_LIMIT",
                    $"The combined ebook HTML exceeds the Markdown working limit of {maximumWorkingHtmlBytes} bytes.", error);
            }

            OfficeConversionException Reject(string code, string message, Exception error) {
                diagnostics.Add(new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Error, message,
                    null, "OfficeIMO.Html", OfficeConversionLossKind.Failure));
                var report = new OfficeConversionReport(options.SourceName!, options.TargetName!, diagnostics: diagnostics);
                return new OfficeConversionException(message, report, error);
            }
        }

        private HtmlConversionDocumentOptions ProjectionHtmlOptions(int maximumCharacters) {
            HtmlConversionDocumentOptions options = HtmlConversionDocumentOptions.UntrustedDefaults;
            options.Limits.MaxInputCharacters = maximumCharacters;
            // Count the complete projection, including generated navigation and anchors.
            if (maximumHtmlNodes.HasValue) options.Limits.MaxHtmlNodes = maximumHtmlNodes;
            return options;
        }

        private static void AddMetadataLoss(List<OfficeDiagnostic> diagnostics, string message) =>
            diagnostics.Add(new OfficeDiagnostic("EBOOK_OUTPUT_METADATA_OMITTED", OfficeDiagnosticSeverity.Warning,
                message, null, "OfficeIMO.Epub", OfficeConversionLossKind.Omission));

        public PdfCreateOptions? GetOutputCreationOptions(OfficeConversionOptions options) =>
            target == "pdf" ? options.Html?.PdfOptions : null;

        public PdfLoadOptions? CreateOutputPdfLoadOptions(OfficeConversionOptions options) =>
            ReadEncryptedOutput(GetOutputCreationOptions(options));

        public async Task<bool> TryValidateArtifactAsync(Stream output, OfficeConversionOptions options, CancellationToken token) {
            if (target == "epub") {
                output.Position = 0;
                await ValidateManuscriptArtifactAsync(output, options.MaximumOutputBytes, null, token).ConfigureAwait(false);
            } else if (target == "html") {
                output.Position = 0;
                int maximumCharacters = (int)Math.Min(options.MaximumOutputBytes, int.MaxValue);
                await HtmlConversionDocument.LoadAsync(output, ProjectionHtmlOptions(maximumCharacters), cancellationToken: token).ConfigureAwait(false);
            } else {
                return false;
            }
            output.Position = output.Length;
            return true;
        }
    }
}
