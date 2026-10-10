using OfficeIMO.Chm;
using OfficeIMO.Reader.Html;
using System.Globalization;

namespace OfficeIMO.Reader.Chm;

internal static class ChmReaderAdapter {
    internal static OfficeDocumentReadResult ReadDocument(string path, ReaderOptions reader, ReaderChmOptions options, CancellationToken token) =>
        Project(DocumentReaderEngine.ReadAdapterInput(path, reader, token, options.ReadOptions.MaxInputBytes), reader, options, token);

    internal static OfficeDocumentReadResult ReadDocument(Stream stream, string? name, ReaderOptions reader, ReaderChmOptions options, CancellationToken token) =>
        Project(DocumentReaderEngine.ReadAdapterInput(stream, name ?? "document.chm", reader, token, options.ReadOptions.MaxInputBytes), reader, options, token);

    private static OfficeDocumentReadResult Project(ReaderAdapterInputSnapshot input, ReaderOptions reader, ReaderChmOptions options, CancellationToken token) {
        ChmDocument book = ChmDocument.Load(input.Bytes, options.ReadOptions, token);
        IReadOnlyList<ChmTopic> topics = book.SelectTopics(options.ConversionOptions);
        var documents = new List<OfficeDocumentReadResult>(topics.Count);
        var chunks = new List<ReaderChunk>();
        long characters = 0;
        int nodes = 0;
        int topicIndex = 0, tableIndex = 0;
        foreach (ChmTopic topic in topics) {
            token.ThrowIfCancellationRequested();
            string html = topic.ReadHtml(token); ChmDocument.ReserveCharacters(ref characters, html.Length, options.ConversionOptions);
            string virtualPath = input.Source.Path + "!" + topic.Path;
            ReaderHtmlOptions htmlOptions = options.HtmlOptions.Clone();
            htmlOptions.ConversionOptions = book.CreateHtmlOptions(topic.Path, htmlOptions.ConversionOptions);
            if (htmlOptions.HtmlToMarkdownOptions != null)
                htmlOptions.HtmlToMarkdownOptions.BaseUri = htmlOptions.ConversionOptions.BaseUri;
            ChmDocument.LimitHtmlNodes(htmlOptions.ConversionOptions, nodes, options.ConversionOptions);
            OfficeDocumentReadResult projected = HtmlReaderAdapter.ReadContentDocument(html, virtualPath, reader, htmlOptions, token, out int topicNodes);
            ChmDocument.ReserveHtmlNodes(ref nodes, topicNodes, options.ConversionOptions);
            string prefix = "chm-topic-" + (++topicIndex).ToString("D5", CultureInfo.InvariantCulture) + "-";
            HtmlReaderAdapter.PrefixProjection(prefix, projected, ref tableIndex);
            foreach (ReaderChunk chunk in projected.Chunks) {
                chunk.Kind = ReaderInputKind.Chm; chunk.Id = prefix + chunk.Id;
                chunk.Location.SourceBlockIndex = topicIndex - 1; chunk.Location.SourceBlockKind = "topic";
                chunk.Location.HeadingPath = string.IsNullOrEmpty(chunk.Location.HeadingPath) ? topic.Title : topic.Title + " / " + chunk.Location.HeadingPath;
                DocumentReaderEngine.ApplyAdapterSource(chunk, input, reader.ComputeHashes); chunks.Add(chunk);
            }
            documents.Add(projected);
        }
        input.Source.Title = book.Title;
        var result = DocumentReaderEngine.CreateDocumentResult(chunks, ReaderInputKind.Chm, input.Source,
            new[] { "officeimo.reader.chm.rich-v5", "officeimo.chm.archive", "officeimo.chm.contents", "officeimo.chm.index" }
                .Concat(documents.SelectMany(item => item.CapabilitiesUsed)), documents.SelectMany(item => item.Assets).ToArray());
        result.Blocks = documents.SelectMany(item => item.Blocks).ToArray(); result.Tables = documents.SelectMany(item => item.Tables).ToArray();
        result.Links = documents.SelectMany(item => item.Links).ToArray(); result.Forms = documents.SelectMany(item => item.Forms).ToArray();
        result.Visuals = documents.SelectMany(item => item.Visuals).ToArray();
        result.Diagnostics = documents.SelectMany(item => item.Diagnostics).Concat(book.Diagnostics.Select(item => new OfficeDocumentDiagnostic {
            Code = item.Code, Message = item.Message, Source = OfficeDocumentReaderBuilderChmExtensions.HandlerId,
            Category = OfficeDocumentDiagnosticCategory.Parsing, Severity = OfficeDocumentDiagnosticSeverity.Warning,
            IsRecoverable = true, Location = new ReaderLocation { Path = input.Source.Path + "!" + (item.Path ?? "/") }
        })).ToArray();
        token.ThrowIfCancellationRequested();
        return result;
    }
}
