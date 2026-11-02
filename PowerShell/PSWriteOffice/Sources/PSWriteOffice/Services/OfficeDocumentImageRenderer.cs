using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Threading;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

/// <summary>Maps shared selectors to the native image owners; Core retains batch budgets and deadlines.</summary>
internal static class OfficeDocumentImageRenderer {
    internal static IReadOnlyList<OfficeImageExportResult> Render(object document, OfficeImageExportFormat format,
        OfficeDocumentImageOptions options, OfficePageRange? pages, OfficeSheetSelector[]? sheets,
        string? range, string? chart, CancellationToken token) {
        var results = new List<OfficeImageExportResult>();
        switch (document) {
            case WordDocument word:
                if (pages != null) options.Word.PageSelection = pages.Selection;
                word.ExportImages(format, results.Add, options.Word, token);
                break;
            case PowerPointDocument slides:
                if (pages != null) options.PowerPoint.SlideNumbers = pages.Selection.Resolve(slides.Slides.Count);
                slides.ExportImages(format, results.Add, options.PowerPoint, token);
                break;
            case PdfDocument pdf:
                options.Pdf.Format = format;
                if (pages != null) options.Pdf.PageSelector = PdfPageSelector.Parse(pages.ToString());
                results.AddRange(pdf.ExportImages(options.Pdf, token));
                break;
            case HtmlConversionDocument html:
                if (pages != null) options.Html.PageSelection = pages.Selection;
                html.ExportImages(format, results.Add, options.Html, token);
                break;
            case ExcelDocument excel:
                if (sheets != null) options.Excel.SheetNames = sheets.Select(sheet => sheet.Resolve(excel.Sheets, value => value.Name).Name).ToArray();
                if (range != null) options.Excel.Range = range;
                if (chart == null) excel.ExportImages(format, results.Add, options.Excel, token);
                else {
                    var selected = options.Excel.SheetNames == null ? excel.Sheets.Where(sheet => !sheet.Hidden || options.Excel.IncludeHiddenSheets)
                        : options.Excel.SheetNames.Select(name => excel[name]);
                    var charts = selected.Select(sheet => sheet.GetChart(chart)
                        ?? throw new PSArgumentException("Chart '" + chart + "' does not exist on sheet '" + sheet.Name + "'.")).ToArray();
                    OfficeImageExportBatchProcessor.ForEachOrdered(charts, 1,
                        (item, _, cancellation) => item.ExportImage(format, options.Excel, cancellation), results.Add, token, options.Excel);
                }
                break;
            case VisioDocument visio:
                if (pages == null) visio.ExportImages(format, results.Add, options.Visio, token);
                else {
                    IReadOnlyList<int> numbers = pages.Selection.Resolve(visio.Pages.Count);
                    var budget = new OfficeImageExportEncodingBudget(options.Visio.MaximumTotalEncodedBytes);
                    OfficeImageExportBatchProcessor.ForEachOrdered(numbers, options.Visio.MaximumDegreeOfParallelism,
                        (number, _, cancellation) => VisioImageExportEngine.Render(visio.Pages[number - 1], format, options.Visio,
                            visio.Pages[number - 1].Name, "Visio page " + number, cancellation, budget),
                        results.Add, token, options.Visio);
                }
                break;
            default: throw new PSArgumentException("The document has no native image exporter.");
        }
        return results.AsReadOnly();
    }
}
