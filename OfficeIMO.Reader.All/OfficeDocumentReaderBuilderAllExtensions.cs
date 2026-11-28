using OfficeIMO.Archives;
using OfficeIMO.Comic;
using OfficeIMO.Fb2;
using OfficeIMO.Mobi;
using OfficeIMO.Reader.AsciiDoc;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Chm;
using OfficeIMO.Reader.Dbf;
using OfficeIMO.Reader.DocBook;
using OfficeIMO.Reader.Email;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Reader.Image;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Latex;
using OfficeIMO.Reader.Markdown;
using OfficeIMO.Reader.Notebook;
using OfficeIMO.Reader.OneNote;
using OfficeIMO.Reader.Opml;
using OfficeIMO.Reader.OpenDocument;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Reader.Publisher;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Reader.Subtitles;
using OfficeIMO.Reader.Toml;
using OfficeIMO.Reader.Visio;
using OfficeIMO.Reader.Xps;
using OfficeIMO.Reader.DjVu;
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
        ArgumentNullException.ThrowIfNull(builder);

        ReaderAllOptions configured = options ?? new ReaderAllOptions();
        ReaderOneNoteOptions oneNoteOptions = configured.OneNote ?? new ReaderOneNoteOptions {
            AllowTableOfContentsSiblingFileReads = false
        };
        builder
            .AddPlainTextHandlers()
            .AddSourceCodeHandler()
            .AddAsciiDocHandler(configured.AsciiDoc)
            .AddCsvHandler(configured.Csv)
            .AddHandler(EpubReaderAdapter.CreateImportHandler("cbr", new[] { ".cbr" },
                ReaderInputKind.Comic, ComicReader.MaximumInputBytes, ComicReader.Import, ComicReader.Probe,
                classifyReadError: static error => error is Rar5EncryptionNotSupportedException ? ReaderErrorCodes.UnsupportedEncryption : null))
            .AddHandler(EpubReaderAdapter.CreateImportHandler("cb7", new[] { ".cb7" },
                ReaderInputKind.Comic, ComicReader.MaximumInputBytes, ComicReader.Import, ComicReader.Probe))
            .AddHandler(EpubReaderAdapter.CreateImportHandler("cbz", new[] { ".cbz" },
                ReaderInputKind.Comic, ComicReader.MaximumInputBytes, ComicReader.Import, ComicReader.Probe))
            .AddChmHandler(configured.Chm)
            .AddDbfHandler(configured.Dbf)
            .AddDocBookHandler(configured.DocBook)
            .AddDjVuHandler(configured.DjVu)
            .AddEmailHandlers(configured.Email)
            .AddEpubHandler(configured.Epub)
            .AddHandler(EpubReaderAdapter.CreateImportHandler("fb2", new[] { ".fb2", ".fb2.zip" },
                ReaderInputKind.Fb2, Fb2Reader.MaximumInputBytes, Fb2Reader.Import, Fb2Reader.Probe))
            .AddExcelHandler(configured.Excel);
        if (configured.IncludeLegacySpreadsheet) {
            builder.AddLegacySpreadsheetHandler(configured.LegacySpreadsheet, configured.Excel);
        }
        builder
            .AddHtmlHandler(configured.Html)
            .AddIWorkHandler(configured.IWork)
            .AddImageHandler(configured.Image)
            .AddJsonHandler(configured.Json)
            .AddLatexHandler(configured.Latex)
            .AddMarkdownHandler(configured.Markdown)
            .AddHandler(EpubReaderAdapter.CreateImportHandler("mobi", new[] { ".mobi", ".azw", ".azw3" },
                ReaderInputKind.Mobi, MobiReader.MaximumInputBytes, MobiReader.Import, MobiReader.Probe, useDetectedKindFallback: true))
            .AddFallbackHandler(EpubReaderAdapter.CreateImportHandler("mobi-prc", new[] { ".prc" },
                ReaderInputKind.Mobi, MobiReader.MaximumInputBytes, MobiReader.Import, MobiReader.Probe))
            .AddFallbackHandler(EpubReaderAdapter.CreateImportHandler("palmdoc", new[] { ".pdb" },
                ReaderInputKind.Mobi, MobiReader.MaximumInputBytes, MobiReader.Import, MobiReader.Probe))
            .AddNotebookHandler(configured.Notebook)
            .AddOneNoteHandler(oneNoteOptions)
            .AddOpmlHandler(configured.Opml)
            .AddOpenDocumentHandler(configured.OpenDocument)
            .AddPdfHandler(configured.Pdf)
            .AddPowerPointHandler(configured.PowerPoint)
            .AddPublisherHandler(configured.Publisher)
            .AddRtfHandler(configured.Rtf)
            .AddSubtitleHandler(configured.Subtitles)
            .AddTomlHandler(configured.Toml)
            .AddVisioHandler(configured.Visio)
            .AddXpsHandler(configured.Xps)
            .AddWordHandlerWithLegacyRouting(configured.Word, configured.LegacyWord, configured.IncludeLegacyWord);
        if (configured.IncludeLegacyWord) {
            builder.AddLegacyWordHandler(configured.LegacyWord, configured.Word);
        }
        return builder
            .AddXmlHandler(configured.Xml)
            .AddYamlHandler(configured.Yaml)
            .AddZipHandler(configured.ZipTraversal, configured.Zip);
    }
}
