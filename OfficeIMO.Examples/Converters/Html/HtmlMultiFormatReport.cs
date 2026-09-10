using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html;

internal static partial class Html {
    /// <summary>Exports one prepared HTML report to PDF, editable Word, and typed Excel.</summary>
    public static void Example_HtmlMultiFormatReport(string folderPath, string? sourcePath = null) {
        sourcePath = Path.GetFullPath(sourcePath ?? Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "Reports", "service-review.html"));
        string name = Path.GetFileNameWithoutExtension(sourcePath);
        string outputFolder = Path.Combine(folderPath, "MultiFormatReport", name);
        Directory.CreateDirectory(outputFolder);
        string stem = Path.Combine(outputFolder, name);

        // Parsing, resource policy, and semantic analysis are shared by the format adapters.
        HtmlConversionDocument report = HtmlConversionDocument.Load(sourcePath);
        File.WriteAllText(stem + ".html", report.SourceHtml);

        var pdfOptions = new HtmlToPdfOptions {
            Mode = HtmlRenderMode.Paged,
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(36D),
            BackgroundColor = OfficeColor.White
        };
        var pdf = report.SaveAsPdf(stem + ".pdf", pdfOptions).RequireSuccess();

        HtmlToWordResult wordResult = report.ToWordDocumentResult(new HtmlToWordOptions {
            DefaultPageSize = WordPageSize.A4,
            SectionTagHandling = SectionTagHandling.Block
        });
        using WordDocument word = wordResult.Value;
        word.Margins.Type = WordMargin.Narrow;
        wordResult.RequireValue().Save(stem + ".docx");

        HtmlToExcelResult excelResult = report.ToExcelDocumentResult(new HtmlToExcelOptions {
            Mode = HtmlImportMode.Generic,
            ImportTypedCellValues = true
        });
        using ExcelDocument excel = excelResult.Value;
        excelResult.RequireValue();
        foreach (ExcelSheet sheet in excel.Sheets) {
            sheet.AutoFitColumns();
            var (firstRow, firstColumn, lastRow, lastColumn) = A1.ParseRange(sheet.UsedRangeA1);
            if (firstColumn == lastColumn) {
                sheet.WrapCells(firstRow, lastRow, firstColumn, 80D);
                for (int row = firstRow; row <= lastRow; row++) sheet.CellVerticalAlign(row, firstColumn, ExcelVerticalAlignment.Top);
            }
            sheet.AutoFitRows();
            sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0, paperSize: ExcelPaperSize.A4);
            sheet.SetMarginsPreset(ExcelMarginPreset.Narrow);
        }
        excel.Save(stem + ".xlsx");

        var index = new List<string> { "# " + name, "", "Exported from `" + Path.GetFileName(sourcePath) + "`.", "" };
        var conversionNotes = new SortedSet<string>(StringComparer.Ordinal);
        foreach (var warning in pdf.Warnings) conversionNotes.Add("PDF: " + warning.Code + ": " + warning.Message);
        foreach (HtmlDiagnostic diagnostic in wordResult.Report.Diagnostics.Concat(excelResult.Report.Diagnostics))
            conversionNotes.Add(diagnostic.Component + ": " + diagnostic.Code + ": " + diagnostic.Message);
        foreach (string extension in new[] { ".html", ".pdf", ".docx", ".xlsx" })
            index.Add("- [" + extension.TrimStart('.').ToUpperInvariant() + "](" + Uri.EscapeDataString(name + extension) + ")");

        // Saved-PDF previews and Office layout previews are labelled separately, with every page retained.
        void Previews(string label, string route, Action<OfficeImageExportConsumer> export) {
            index.Add(""); index.Add("## " + label); index.Add("");
            int page = 0;
            export(image => {
                string fileName = route + "-" + (++page).ToString("D3") + ".png";
                File.WriteAllBytes(Path.Combine(outputFolder, fileName), image.Bytes);
                index.Add("- [Page or sheet " + page + "](" + fileName + ")");
                foreach (var diagnostic in image.Diagnostics)
                    conversionNotes.Add(label + ": " + diagnostic.Code + ": " + diagnostic.Message);
            });
        }
        Previews("Saved PDF pages", "pdf", consume => {
            foreach (var image in PdfDocument.Load(File.ReadAllBytes(stem + ".pdf")).Render.ExportImages(OfficeImageExportFormat.Png)) consume(image);
        });
        Previews("HTML layout pages", "html", consume => {
            foreach (var image in report.ExportImages(OfficeImageExportFormat.Png, pdfOptions)) consume(image);
        });
        Previews("Word layout pages", "word", consume => word.ExportImages(OfficeImageExportFormat.Png, consume));
        Previews("Excel worksheet layouts", "excel", consume => excel.ExportImages(OfficeImageExportFormat.Png, consume));
        index.Add(""); index.Add("## Conversion notes"); index.Add("");
        index.Add("Word and Excel exports use their own editable document layouts; page counts can differ from the source report.");
        index.Add("");
        foreach (string note in conversionNotes) index.Add("- " + note);
        if (conversionNotes.Count == 0) index.Add("No conversion or preview diagnostics were reported.");
        File.WriteAllLines(Path.Combine(outputFolder, "index.md"), index);

        Console.WriteLine($"Report outputs: {outputFolder}");
        Console.WriteLine($"PDF pages: {pdf.Serialization?.PageCount}; Excel sheets: {excelResult.Sheets}; typed import cells: {excelResult.Cells}");
        foreach (HtmlDiagnostic diagnostic in wordResult.Report.Diagnostics.Concat(excelResult.Report.Diagnostics)) {
            Console.WriteLine($"{diagnostic.Component}: {diagnostic.Code}: {diagnostic.Message} [{diagnostic.Source}] {diagnostic.Detail}");
        }
    }
}
