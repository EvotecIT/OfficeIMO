using OfficeIMO.Pdf;

namespace OfficeIMO.Reader.Pdf;

internal static partial class PdfReaderAdapter {
    private static PdfLoadOptions GetEffectiveReadOptions(ReaderOptions readerOptions, ReaderPdfOptions pdfOptions) {
        PdfLoadOptions source = (pdfOptions.ReadOptions ?? PdfLoadOptions.Default).Clone();
        if (readerOptions.Projection != ReaderProjection.TextAndMetadata) return source;
        return new PdfLoadOptions {
            Profile = source.Profile, PageSelection = source.PageSelection,
            LayoutOptions = source.LayoutOptions, Pipeline = source.Pipeline,
            IncludeImages = false, IncludeFormWidgets = false
        };
    }

    private static IEnumerable<ReaderChunk> ReadTextAndMetadataChunks(IReadOnlyList<PdfLogicalPage> pages,
        SourceMetadata source, ReaderOptions options, CancellationToken token, bool chunkByPage = true) {
        if (!chunkByPage) {
            string documentText = string.Join(Environment.NewLine + Environment.NewLine,
                pages.Select(page => string.Join(Environment.NewLine, page.TextBlocks.Select(block => block.Text))));
            IReadOnlyList<string> documentParts = DocumentReaderEngine.SplitAdapterProjection(documentText, options.MaxChars);
            for (int index = 0; index < documentParts.Count; index++) {
                token.ThrowIfCancellationRequested();
                yield return EnrichChunk(new ReaderChunk {
                    Id = "pdf-document-text-" + index.ToString("D4", CultureInfo.InvariantCulture),
                    Kind = ReaderInputKind.Pdf, Text = documentParts[index], ContinuesPreviousChunk = index > 0,
                    Location = new ReaderLocation { Path = source.Path, SourceBlockIndex = 0, SourceBlockKind = "document", BlockAnchor = "document" }
                }, source, options.ComputeHashes);
            }
            yield break;
        }
        for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
            token.ThrowIfCancellationRequested();
            PdfLogicalPage page = pages[pageIndex];
            int selectionIndex = GetPageSelectionIndex(source, pageIndex);
            string text = string.Join(Environment.NewLine, page.TextBlocks.Select(block => block.Text));
            IReadOnlyList<string> parts = DocumentReaderEngine.SplitAdapterProjection(text, options.MaxChars);
            for (int index = 0; index < parts.Count; index++) {
                token.ThrowIfCancellationRequested();
                string anchor = "page-" + page.PageNumber.ToString(CultureInfo.InvariantCulture) + "-selection-" + selectionIndex.ToString("D4", CultureInfo.InvariantCulture);
                yield return EnrichChunk(new ReaderChunk {
                    Id = "pdf-text-" + selectionIndex.ToString("D4", CultureInfo.InvariantCulture) + "-part-" + index.ToString("D4", CultureInfo.InvariantCulture),
                    Kind = ReaderInputKind.Pdf, Text = parts[index], ContinuesPreviousChunk = index > 0,
                    Location = BuildLocation(source, page.PageNumber, selectionIndex, "page", anchor)
                }, source, options.ComputeHashes);
            }
        }
    }

    private static OfficeDocumentReadResult BuildTextAndMetadataResult(PdfDocumentReadResult document, SourceMetadata source,
        ReaderOptions options, ReaderPdfOptions pdfOptions, PdfDocumentPreflight? preflight, bool applyPageRanges,
        CancellationToken token, PdfInspectionReadResult? inspection) {
        IReadOnlyList<PdfLogicalPage> pages = applyPageRanges ? GetReaderPages(document, pdfOptions) : document.Pages;
        ReaderChunk[] chunks = ReadTextAndMetadataChunks(pages, source, options, token, pdfOptions.ChunkByPage)
            .Concat(BuildInspectionWarningChunks(inspection, source)).ToArray();
        var metadata = BuildDocumentMetadata(document, source, preflight, applyPageRanges ? pages : null,
            Array.Empty<ReaderTable>(), Array.Empty<OfficeDocumentOcrCandidate>(), pdfOptions, textAndMetadata: true).ToList();
        AddMetadata(metadata, "pdf-projection-images", "pdf.projection", "Images", "not-requested", "coverage");
        AddMetadata(metadata, "pdf-projection-widgets", "pdf.projection", "LogicalFormWidgets", "not-requested; bounded source field facts and values retained", "coverage");
        AddCountMetadata(metadata, "pdf-source-form-widget-count", "pdf.form", "SourceWidgetCount", document.FormFields.Sum(field => field.Widgets.Count));
        var result = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Source = new OfficeDocumentSource {
                Path = source.Path, SourceId = source.SourceId, SourceHash = source.SourceHash,
                LastWriteUtc = source.LastWriteUtc, LengthBytes = source.LengthBytes,
                Title = document.Metadata.Title, Author = document.Metadata.Author,
                Subject = document.Metadata.Subject, Keywords = document.Metadata.Keywords
            },
            CapabilitiesUsed = new[] { "officeimo.reader.pdf", "officeimo.pdf.logical-document", "officeimo.reader.pdf.pages.native", "officeimo.reader.pdf.text-and-metadata" },
            Chunks = chunks, Metadata = metadata.AsReadOnly(),
            Pages = BuildDocumentPages(pages, source, Array.Empty<OfficeDocumentBlock>(), Array.Empty<ReaderTable>(),
                Array.Empty<OfficeDocumentAsset>(), Array.Empty<OfficeDocumentLink>(), Array.Empty<OfficeDocumentFormField>(), Array.Empty<OfficeDocumentOcrCandidate>()),
            Diagnostics = BuildDocumentDiagnostics(chunks, Array.Empty<OfficeDocumentOcrCandidate>(), preflight)
        };
        return AddInspectionEvidence(result, pages, inspection, options, source, token);
    }
}
