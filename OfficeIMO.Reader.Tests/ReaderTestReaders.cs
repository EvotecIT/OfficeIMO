using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Markdown;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Word;

namespace OfficeIMO.Reader.Tests;

internal static class ReaderTestReaders {
    internal static OfficeDocumentReader All { get; } = new OfficeDocumentReaderBuilder()
        .AddAllOfficeIMOHandlers()
        .Build();

    internal static OfficeDocumentReader Excel(string? sheetName = null, string? a1Range = null, int? chunkRows = null) =>
        new OfficeDocumentReaderBuilder()
            .AddExcelHandler(new ReaderExcelOptions {
                SheetName = sheetName,
                A1Range = a1Range,
                ChunkRows = chunkRows ?? 200
            })
            .Build();

    internal static OfficeDocumentReader Word(
        bool includeFootnotes = true,
        bool includePageLocations = false) =>
        new OfficeDocumentReaderBuilder()
            .AddWordHandler(new ReaderWordOptions {
                IncludeFootnotes = includeFootnotes,
                IncludePageLocations = includePageLocations
            })
            .Build();

    internal static OfficeDocumentReader PowerPoint(
        bool includeNotes = true,
        bool includeHiddenShapes = true) =>
        new OfficeDocumentReaderBuilder()
            .AddPowerPointHandler(new ReaderPowerPointOptions {
                IncludeNotes = includeNotes,
                IncludeHiddenShapes = includeHiddenShapes
            })
            .Build();

    internal static OfficeDocumentReader Markdown(OfficeIMO.Markdown.MarkdownReaderOptions parserOptions) =>
        new OfficeDocumentReaderBuilder()
            .AddMarkdownHandler(new ReaderMarkdownOptions { ParserOptions = parserOptions })
            .Build();

    internal static OfficeDocumentReader OpenDocument(string? sheetName = null, string? a1Range = null) =>
        new OfficeDocumentReaderBuilder()
            .AddOpenDocumentHandler(new ReaderOpenDocumentOptions { SheetName = sheetName, A1Range = a1Range })
            .Build();
}
