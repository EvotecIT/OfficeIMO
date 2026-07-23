using OfficeIMO.Reader.AsciiDoc;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Image;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Latex;
using OfficeIMO.Reader.Markdown;
using OfficeIMO.Reader.Notebook;
using OfficeIMO.Reader.OneNote;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Reader.Subtitles;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Reader.Word;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Yaml;
using OfficeIMO.Reader.Zip;

namespace OfficeIMO.Reader.All;

/// <summary>Adds the local OfficeIMO format adapters to an isolated reader builder.</summary>
public static class OfficeDocumentReaderBuilderAllExtensions {
    /// <summary>
    /// Adds all local, in-process OfficeIMO handlers included by this package.
    /// </summary>
    /// <param name="builder">The isolated reader builder to configure.</param>
    /// <param name="options">Optional format-specific settings. Defaults are bounded and deterministic.</param>
    /// <returns>The same builder for fluent composition.</returns>
    /// <remarks>
    /// This preset intentionally excludes OCR engines, process adapters, network clients, and hosted providers.
    /// It only composes the handler packages referenced by <c>OfficeIMO.Reader.All</c>.
    /// </remarks>
    public static OfficeDocumentReaderBuilder AddAllOfficeIMOHandlers(
        this OfficeDocumentReaderBuilder builder,
        ReaderAllOptions? options = null) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));

        ReaderAllOptions configured = options ?? new ReaderAllOptions();
        ReaderOneNoteOptions oneNoteOptions = configured.OneNote ?? new ReaderOneNoteOptions {
            AllowTableOfContentsSiblingFileReads = false
        };
        return builder
            .AddPlainTextHandlers()
            .AddAsciiDocHandler(configured.AsciiDoc)
            .AddCsvHandler(configured.Csv)
            .AddEmailHandlers(configured.Email)
            .AddEpubHandler(configured.Epub)
            .AddExcelHandler(configured.Excel)
            .AddHtmlHandler(configured.Html)
            .AddImageHandler(configured.Image)
            .AddJsonHandler(configured.Json)
            .AddLatexHandler(configured.Latex)
            .AddMarkdownHandler(configured.Markdown)
            .AddNotebookHandler(configured.Notebook)
            .AddOneNoteHandler(oneNoteOptions)
            .AddOpenDocumentHandler(configured.OpenDocument)
            .AddPdfHandler(configured.Pdf)
            .AddPowerPointHandler(configured.PowerPoint)
            .AddRtfHandler(configured.Rtf)
            .AddSubtitleHandler(configured.Subtitles)
            .AddVisioHandler(configured.Visio)
            .AddWordHandler(configured.Word)
            .AddXmlHandler(configured.Xml)
            .AddYamlHandler(configured.Yaml)
            .AddZipHandler(configured.ZipTraversal, configured.Zip);
    }
}
