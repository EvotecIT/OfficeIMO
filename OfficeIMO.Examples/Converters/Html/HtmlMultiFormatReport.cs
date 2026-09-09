using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html;

internal static partial class Html {
    /// <summary>Exports one prepared HTML report to PDF, editable Word, and typed Excel.</summary>
    public static void Example_HtmlMultiFormatReport(string folderPath) {
        string sourcePath = Path.Combine(AppContext.BaseDirectory, "Converters", "Html", "Content", "Reports", "service-review.html");
        string outputFolder = Path.Combine(folderPath, "MultiFormatReport");
        Directory.CreateDirectory(outputFolder);
        string stem = Path.Combine(outputFolder, "service-review");

        // Parsing, resource policy, and semantic analysis are shared by the format adapters.
        HtmlConversionDocument report = HtmlConversionDocument.Load(sourcePath);
        File.WriteAllText(stem + ".html", report.SourceHtml);

        var pdfOptions = new HtmlToPdfOptions {
            PageSize = OfficePageSizes.A4,
            Margins = HtmlRenderMargins.All(36D),
            BackgroundColor = OfficeColor.White
        };
        var pdf = report.SaveAsPdf(stem + ".pdf", pdfOptions).RequireSuccess();

        HtmlToWordResult wordResult = report.ToWordDocumentResult();
        using WordDocument word = wordResult.Value;
        wordResult.RequireValue().Save(stem + ".docx");

        HtmlToExcelResult excelResult = report.ToExcelDocumentResult(new HtmlToExcelOptions {
            Mode = HtmlImportMode.Generic,
            ImportTypedCellValues = true
        });
        using ExcelDocument excel = excelResult.Value;
        excelResult.RequireValue();
        foreach (ExcelSheet sheet in excel.Sheets) sheet.AutoFitColumns();
        excel.Save(stem + ".xlsx");

        // A preview is optional; the actual PDF is written by the same first-party renderer.
        report.ToImage(pdfOptions).AsPng().OnFileConflict(OfficeImageExportFileConflictPolicy.Replace).Save(stem + ".png");
        word.ToImage().AsPng().OnFileConflict(OfficeImageExportFileConflictPolicy.Replace).Save(stem + "-word.png");
        excel.Sheets[0].ToImage().AsPng().OnFileConflict(OfficeImageExportFileConflictPolicy.Replace).Save(stem + "-excel.png");

        Console.WriteLine($"Report outputs: {outputFolder}");
        Console.WriteLine($"PDF pages: {pdf.Serialization?.PageCount}; Excel sheets: {excelResult.Sheets}; typed import cells: {excelResult.Cells}");
        foreach (HtmlDiagnostic diagnostic in wordResult.Report.Diagnostics.Concat(excelResult.Report.Diagnostics)) {
            Console.WriteLine($"{diagnostic.Component}: {diagnostic.Code}: {diagnostic.Message} [{diagnostic.Source}] {diagnostic.Detail}");
        }
    }
}
