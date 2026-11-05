using OfficeIMO.Pdf.Word;
using OfficeIMO.Pdf.Excel;
using OfficeIMO.Core.Internal;
using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
        bool overwrite, PdfExcelExportOptions excelOptions, PdfWordExportOptions? wordOptions, CancellationToken cancellationToken,
        PdfHtmlExportOptions? htmlOptions = null, CsvSaveOptions? csvOptions = null)
        => ExportAsync(review, outputPath, new OfficeConversionOptions {
            FileConflictPolicy = overwrite ? OfficePublicationPolicy.Replace : OfficePublicationPolicy.Fail,
            PdfExcel = excelOptions, PdfWord = wordOptions, PdfHtml = htmlOptions, CsvSave = csvOptions
        }, cancellationToken).GetAwaiter().GetResult();

    /// <summary>Publishes existing native OCR projections through the canonical atomic conflict and resource policies.</summary>
    internal static async Task<OfficeImageConversionResult> ExportAsync(PdfSearchableOcrReview review, string outputPath,
        OfficeConversionOptions options, CancellationToken cancellationToken) {
        var captured = options.Clone();
        var excelOptions = captured.PdfExcel ?? new PdfExcelExportOptions();
        var wordOptions = captured.PdfWord;
        var htmlOptions = captured.PdfHtml;
        var csvOptions = captured.CsvSave;
        cancellationToken.ThrowIfCancellationRequested();
        ArgumentNullException.ThrowIfNull(review);
        review.Ocr.RequireAcceptedOcrContent();
        string extension = Path.GetExtension(outputPath).ToLowerInvariant();
        ValidateExtension(extension);
        var document = review.Ocr.Document;
        if ((extension == ".xlsx" || extension == ".csv") && document.Tables.Count == 0)
            throw new InvalidOperationException("No tables were detected. Review the recognition settings or export the recognized content as Word, text or JSON.");
        if (extension == ".csv" && document.Tables.Count != 1)
            throw new InvalidOperationException("CSV requires exactly one detected table. Use Excel or JSON to preserve multiple tables.");
        string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath))!;
        if (captured.PublicationGuard is IOfficeStagingGuard staging) await staging.EnsureStagingDirectoryAllowedAsync(directory, cancellationToken).ConfigureAwait(false);
        Directory.CreateDirectory(directory);
        if (captured.PublicationGuard is IOfficeStagingGuard created) await created.EnsureStagingDirectoryAllowedAsync(directory, cancellationToken).ConfigureAwait(false);
        string temporaryPath = OfficeFileOutput.CreateTemporaryPath(directory, Path.GetFileName(outputPath));
        OfficeConversionReport? report = null;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            switch (extension)
            {
                case ".xlsx": report = review.Ocr.SaveAsExcel(temporaryPath, excelOptions, cancellationToken); break;
                case ".docx": report = review.Ocr.SaveAsWord(temporaryPath,
                    wordOptions ?? new PdfWordExportOptions { ImportImages = false, IncludeImagePlaceholders = false }, cancellationToken); break;
                case ".html": report = review.Ocr.SaveAsHtml(temporaryPath,
                    htmlOptions ?? new PdfHtmlExportOptions { ImageExportMode = PdfHtmlImageExportMode.PlaceholderOnly, IncludeImagePlaceholders = false }, cancellationToken); break;
                case ".pdf": using (var searchable = review.ApplyAll(cancellationToken).Document) searchable.Save(temporaryPath); break;
                case ".txt": WriteText(temporaryPath, review.Ocr.Text); break;
                case ".md": WriteText(temporaryPath, document.ExtractMarkdown(new PdfLogicalMarkdownOptions { IncludeImagePlaceholders = false })); break;
                case ".json": WriteText(temporaryPath, document.ExportStructured(PdfStructuredExportFormat.Json)); break;
                case ".csv":
                    using (var text = new StreamWriter(temporaryPath, false, new UTF8Encoding(false)))
                        WriteCsv(text, document.Tables[0], csvOptions, cancellationToken);
                    break;
            }
            cancellationToken.ThrowIfCancellationRequested();
            var scope = extension == ".xlsx" || extension == ".csv" ? PdfLogicalTableAnalysis.AnalyzeExtractionScope(document) : null;
            var evidence = new OfficeImageConversionResult(outputPath, review, report, scope);
            if (captured.FailOnLoss) evidence.ThrowIfLoss();
            if (new FileInfo(temporaryPath).Length > captured.MaximumOutputBytes)
                throw new InvalidOperationException("The OCR conversion output exceeds MaximumOutputBytes.");
            string published = await OfficeFilePublisher.PublishAsync(temporaryPath, Path.GetFullPath(outputPath),
                captured.FileConflictPolicy, captured.PublicationGuard, cancellationToken).ConfigureAwait(false);
            return new OfficeImageConversionResult(published, review, report?.WithOutputPath(published), scope);
        } finally { if (File.Exists(temporaryPath)) File.Delete(temporaryPath); }
    }

    internal static void ValidateTextExtension(string extension) {
        if (extension is not (".txt" or ".md" or ".json" or ".html" or ".csv"))
            throw new ArgumentException("OCR text output requires Text, Markdown, Json, Html or Csv; binary output requires OutputPath.");
    }

    internal static string ExportText(PdfSearchableOcrReview review, string extension, OfficeConversionOptions options,
        CancellationToken token, out OfficeConversionReport report) {
        ValidateTextExtension(extension);
        token.ThrowIfCancellationRequested();
        review.Ocr.RequireAcceptedOcrContent();
        var captured = options.Clone();
        var document = review.Ocr.Document;
        OfficeConversionReport? projection = null;
        string text;
        switch (extension) {
            case ".txt": text = review.Ocr.Text; break;
            case ".md": text = document.ExtractMarkdown(new PdfLogicalMarkdownOptions { IncludeImagePlaceholders = false }); break;
            case ".json": text = document.ExportStructured(PdfStructuredExportFormat.Json); break;
            case ".html": text = review.Ocr.ToHtml(captured.PdfHtml ?? new PdfHtmlExportOptions {
                ImageExportMode = PdfHtmlImageExportMode.PlaceholderOnly, IncludeImagePlaceholders = false }, out projection, token); break;
            default:
                if (document.Tables.Count != 1) throw new InvalidOperationException("CSV requires exactly one detected table. Use Excel or JSON to preserve multiple tables.");
                using (var writer = new StringWriter(System.Globalization.CultureInfo.InvariantCulture)) {
                    WriteCsv(writer, document.Tables[0], captured.CsvSave, token);
                    text = writer.ToString();
                }
                break;
        }
        token.ThrowIfCancellationRequested();
        var scope = extension == ".csv" ? PdfLogicalTableAnalysis.AnalyzeExtractionScope(document) : null;
        report = new OfficeConversionReport("OCR review", extension, itemsWritten: review.Ocr.Pages.Count,
            diagnostics: OfficeImageConversionResult.DiagnosticsFor(review, projection, scope));
        if (captured.FailOnLoss) report.ThrowIfLoss();
        if (Encoding.UTF8.GetByteCount(text) > captured.MaximumOutputBytes) throw new InvalidOperationException("The OCR conversion output exceeds MaximumOutputBytes.");
        return text;
    }

    private static void WriteCsv(TextWriter writer, PdfLogicalTable table, CsvSaveOptions? options, CancellationToken token) {
        var settings = options?.Clone() ?? new CsvSaveOptions { FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape };
        settings.IncludeHeader = false;
        using var csv = new CsvRowWriter(writer, settings);
        foreach (var row in table.Rows) {
            token.ThrowIfCancellationRequested();
            csv.WriteTextRow(table.Rows[0], row.Select(value => (string?)value).ToArray());
        }
    }

    private static void WriteText(string path, string text) => File.WriteAllText(path, text, new UTF8Encoding(false));
}
