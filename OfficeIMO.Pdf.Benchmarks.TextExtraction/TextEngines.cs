using System.Text;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using Tabula;
using Tabula.Extractors;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace OfficeIMO.Pdf.Benchmarks.TextExtraction;

/// <summary>
/// The measured operation for each engine: open the file from disk and return the plain text of every page.
/// Each delegate is the cheapest public whole-document text call the engine offers.
/// </summary>
internal static class TextEngines {
    internal static readonly string[] All = ["officeimo", "officeimo-read", "pdfoxide", "pdfium", "pdfpig", "pdfpig-tabula"];

    internal static Func<string, string> Resolve(string engine) => engine switch {
        "officeimo" => OfficeImoExtractText,
        "officeimo-read" => OfficeImoRead,
        "pdfoxide" => PdfOxideExtractText,
        "pdfium" => PdfiumExtractText,
        "pdfpig" => PdfPigExtractText,
        "pdfpig-tabula" => PdfPigTabulaExtractText,
        _ => throw new ArgumentException($"Unknown engine '{engine}'. Known engines: {string.Join(", ", All)}.")
    };

    /// <summary><c>PdfDocument.ExtractText()</c>: plain page text through the bounded read pipeline.</summary>
    private static string OfficeImoExtractText(string path) {
        using global::OfficeIMO.Pdf.PdfDocument document = global::OfficeIMO.Pdf.PdfDocument.Load(path);
        return document.ExtractText();
    }

    /// <summary>
    /// <c>PdfDocument.Read()</c> with the default structured profile: semantic reconstruction of blocks,
    /// headings, lists, tables and document-wide evidence. It does more work than plain text extraction.
    /// </summary>
    private static string OfficeImoRead(string path) {
        using global::OfficeIMO.Pdf.PdfDocument document = global::OfficeIMO.Pdf.PdfDocument.Load(path);
        return document.Read().Text;
    }

    /// <summary>
    /// <c>ExtractText(page)</c> for every page. <c>ExtractAllText()</c> drops the text of a failing page without
    /// reporting it, so the per-page call is used to make page failures visible.
    /// </summary>
    private static string PdfOxideExtractText(string path) {
        using var document = PdfOxide.Core.PdfDocument.Open(path);
        int pages = document.PageCount;
        var text = new StringBuilder();
        for (int page = 0; page < pages; page++) {
            text.Append(document.ExtractText(page)).Append("\n\n");
        }
        return text.ToString();
    }

    /// <summary>PDFium <c>FPDFText_GetText</c> over each page through Docnet.Core.</summary>
    private static string PdfiumExtractText(string path) {
        using IDocReader document = DocLib.Instance.GetDocReader(path, new PageDimensions(1.0));
        int pages = document.GetPageCount();
        var text = new StringBuilder();
        for (int page = 0; page < pages; page++) {
            using IPageReader reader = document.GetPageReader(page);
            text.Append(reader.GetText()).Append("\n\n");
        }
        return text.ToString();
    }

    /// <summary>PdfPig <c>Page.Text</c> (content-stream order) for every page.</summary>
    private static string PdfPigExtractText(string path) {
        using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(path);
        var text = new StringBuilder();
        foreach (UglyToad.PdfPig.Content.Page page in document.GetPages()) {
            text.Append(page.Text).Append("\n\n");
        }
        return text.ToString();
    }

    /// <summary>
    /// The usual .NET open-source pairing for structured reading: PdfPig's line-preserving
    /// <c>ContentOrderTextExtractor</c> for page text, then Tabula's lattice extractor for ruled tables, with each
    /// table row appended after the page text as tab-separated cells.
    /// </summary>
    private static string PdfPigTabulaExtractText(string path) {
        using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(path, new UglyToad.PdfPig.ParsingOptions { ClipPaths = true });
        var text = new StringBuilder();
        var lattice = new SpreadsheetExtractionAlgorithm();
        for (int pageNumber = 1; pageNumber <= document.NumberOfPages; pageNumber++) {
            // One parse of the page serves both extractors.
            UglyToad.PdfPig.Content.Page pdfPage = document.GetPage(pageNumber);
            text.Append(ContentOrderTextExtractor.GetText(pdfPage)).Append('\n');
            PageArea page = ObjectExtractor.ExtractPage(pdfPage);
            foreach (Table table in lattice.Extract(page)) {
                foreach (IReadOnlyList<Cell> row in table.Rows) {
                    text.AppendJoin('\t', row.Select(static cell => cell.GetText())).Append('\n');
                }
            }
            text.Append('\n');
        }
        return text.ToString();
    }
}
