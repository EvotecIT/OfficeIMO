using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO;
using OfficeIMO.CSV;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Word.Pdf;

namespace PSWriteOffice.Services.Ocr;

/// <summary>Host output dispatch over the existing OfficeIMO recognition and format adapters.</summary>
internal static class OfficeImageExporter
{
    internal static void ValidateExtension(string extension)
    {
        switch (extension)
        {
            case ".xlsx": case ".docx": case ".pdf": case ".txt":
            case ".md": case ".html": case ".json": case ".csv": return;
            default: throw new ArgumentException("Choose .xlsx, .docx, .pdf, .txt, .md, .html, .json or .csv output.");
        }
    }

    internal static OfficeImageConversionResult Export(PdfSearchableOcrReview review, string outputPath,
        bool overwrite, PdfTablesToExcelOptions excelOptions, PdfToWordOptions? wordOptions, CancellationToken cancellationToken,
        PdfToHtmlOptions? htmlOptions = null, CsvSaveOptions? csvOptions = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (review == null) throw new ArgumentNullException(nameof(review));
        review.Ocr.RequireAcceptedOcrContent();
        string extension = Path.GetExtension(outputPath).ToLowerInvariant();
        ValidateExtension(extension);
        var document = review.Ocr.Document;
        if ((extension == ".xlsx" || extension == ".csv") && document.Tables.Count == 0)
            throw new InvalidOperationException("No tables were detected. Review the recognition settings or export the recognized content as Word, text or JSON.");
        if (extension == ".csv" && document.Tables.Count != 1)
            throw new InvalidOperationException("CSV requires exactly one detected table. Use Excel or JSON to preserve multiple tables.");
        IOfficeConversionReport? report = null;
        AtomicFileWriter.Write(outputPath, overwrite, temporaryPath =>
        {
            cancellationToken.ThrowIfCancellationRequested();
            switch (extension)
            {
                case ".xlsx": report = document.SaveTablesAsExcel(temporaryPath, excelOptions, cancellationToken).RequireSuccess().Report; break;
                case ".docx": report = document.SaveAsWord(temporaryPath,
                    wordOptions ?? new PdfToWordOptions { ImportImages = false, IncludeImagePlaceholders = false }, cancellationToken).RequireSuccess().Report; break;
                case ".html": report = document.SaveAsHtml(temporaryPath,
                    htmlOptions ?? new PdfToHtmlOptions { ImageExportMode = PdfHtmlImageExportMode.PlaceholderOnly, IncludeImagePlaceholders = false }, cancellationToken).RequireSuccess().Report; break;
                case ".pdf": review.ApplyAll(cancellationToken).Document.Save(temporaryPath).RequireSuccess(); break;
                case ".txt": WriteText(temporaryPath, review.Ocr.Text); break;
                case ".md": WriteText(temporaryPath, document.ToMarkdown(new PdfLogicalMarkdownOptions { IncludeImagePlaceholders = false })); break;
                case ".json": WriteText(temporaryPath, document.ExportStructured(PdfStructuredExportFormat.Json)); break;
                case ".csv":
                    var table = document.Tables[0];
                    var csvSettings = csvOptions?.Clone() ?? new CsvSaveOptions { FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape };
                    // The first source row is emitted as data; do not synthesize another header.
                    csvSettings.IncludeHeader = false;
                    using (var text = new StreamWriter(temporaryPath, false, new UTF8Encoding(false)))
                    using (var csv = new CsvRowWriter(text, csvSettings))
                        foreach (var row in table.Rows)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                            csv.WriteTextRow(table.Rows[0], row.Select(value => (string?)value).ToArray());
                        }
                    break;
            }
            cancellationToken.ThrowIfCancellationRequested();
        });
        return new OfficeImageConversionResult(outputPath, review, report,
            extension == ".xlsx" || extension == ".csv" ? PdfLogicalTableAnalysis.AnalyzeExtractionScope(document) : null);
    }

    private static void WriteText(string path, string text) => File.WriteAllText(path, text, new UTF8Encoding(false));
}
