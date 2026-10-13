using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentReaderAdapter {
    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions options,
        ReaderOpenDocumentOptions formatOptions, CancellationToken token) {
        ReaderInputLimits.EnforceFileSize(path, options.MaxInputBytes);
        token.ThrowIfCancellationRequested();
        OdfDocument document = string.Equals(Path.GetExtension(path), ".fodg", StringComparison.OrdinalIgnoreCase)
            ? OdgDocument.LoadFlatXml(path, CreateOpenOptions(options, formatOptions))
            : OdfDocument.Load(path, CreateOpenOptions(options, formatOptions));
        return ProjectDocument(document, Path.GetFullPath(path), options, formatOptions, token);
    }

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? name, ReaderOptions options,
        ReaderOpenDocumentOptions formatOptions, CancellationToken token) {
        Stream input = ReaderInputLimits.EnsureSeekableReadStream(stream, options.MaxInputBytes, token, out bool owns);
        try {
            token.ThrowIfCancellationRequested();
            string source = string.IsNullOrWhiteSpace(name) ? "document.odf" : name!;
            OdfDocument document = string.Equals(Path.GetExtension(source), ".fodg", StringComparison.OrdinalIgnoreCase)
                ? OdgDocument.LoadFlatXml(input, CreateOpenOptions(options, formatOptions))
                : OdfDocument.Load(input, CreateOpenOptions(options, formatOptions));
            return ProjectDocument(document, source, options, formatOptions, token);
        } finally { if (owns) input.Dispose(); }
    }

    private static OfficeDocumentReadResult ProjectDocument(OdfDocument document, string path, ReaderOptions options,
        ReaderOpenDocumentOptions formatOptions, CancellationToken token) {
        ReaderChunk[] chunks = ReadDocument(document, path, options, formatOptions, token).ToArray();
        var result = DocumentReaderEngine.CreateDocumentResult(chunks, ReaderInputKind.OpenDocument,
            capabilities: new[] { "officeimo.reader.opendocument.native-rich-v1" });
        result.Source.Title = document.Metadata.Title;
        result.Source.Author = document.Metadata.Creator;
        result.Source.Subject = document.Metadata.Subject;
        result.Diagnostics = result.Diagnostics.Concat(document.Diagnostics.Select(diagnostic => new OfficeDocumentDiagnostic {
            Code = diagnostic.Id, Message = diagnostic.Message, Source = "officeimo.opendocument",
            Category = OfficeDocumentDiagnosticCategory.Parsing,
            Severity = diagnostic.Severity == OdfDiagnosticSeverity.Error ? OfficeDocumentDiagnosticSeverity.Error : OfficeDocumentDiagnosticSeverity.Warning,
            Location = new ReaderLocation { Path = path, StartLine = diagnostic.LineNumber },
            Attributes = ReaderInspectionMetadataBuilder.Attributes(("partPath", diagnostic.PartPath), ("linePosition", diagnostic.LinePosition?.ToString(CultureInfo.InvariantCulture)))
        })).ToArray();
        var resources = new OpenDocumentRichResources(document, formatOptions, result, token);
        var rich = new OpenDocumentRichBlocks(formatOptions, token);
        if (document is OdtDocument text) {
            rich.Retain(result.Blocks);
            ProjectTextRich(text, path, options, resources, rich, result, token);
        } else if (document is OdpPresentation presentation) {
            ProjectPresentationRich(presentation, path, options, formatOptions, resources, rich, result, token);
        } else if (document is OdgDocument drawing) {
            ProjectDrawingRich(drawing, path, resources, rich, result, token);
        } else rich.Retain(result.Blocks);
        result.Blocks = rich.Blocks;
        foreach (OfficeDocumentPage page in result.Pages) page.Blocks = result.Blocks.Where(block =>
            page.Location.Slide.HasValue ? block.Location.Slide == page.Location.Slide :
            page.Location.Page.HasValue ? block.Location.Page == page.Location.Page : block.Location.Sheet == page.Location.Sheet).ToArray();
        resources.Complete();
        DocumentReaderEngine.EnrichDocumentResult(result, Array.Empty<string>(), result.Blocks, result.Tables,
            result.Links, result.Visuals, result.Pages);
        if (options.Inspection != null) return OpenDocumentInspectionProjection.Apply(document, path, options, formatOptions, result, token);
        return result;
    }

}
