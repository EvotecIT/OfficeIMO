using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Qualifies the small source-linked workbook format subset that chart snapshots can render exactly.</summary>
internal static class OfficeOpenXmlChartAxisNumberFormats {
    internal static string ResolveSourceLinkedGeneral(ChartPart chartPart) {
        // A source-linked value axis takes the number format of the plotted cells, which the
        // chart caches beside every value. General caches need no workbook projection.
        if (HasOnlyGeneralValueCaches(chartPart)) return "General";
        EmbeddedPackagePart? embedded = OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(chartPart);
        if (embedded == null)
            throw new NotSupportedException("The source-linked chart axis has no qualified embedded workbook.");
        using var source = embedded.GetStream(FileMode.Open, FileAccess.Read);
        byte[] bytes = OfficeOpenXmlChartWorkbookSecurity.ReadAndValidate(source);
        using var stream = new MemoryStream(bytes, writable: false);
        using var workbook = SpreadsheetDocument.Open(stream, false, OfficeOpenXmlChartWorkbookSecurity.CreateOpenSettings());
        if (workbook.WorkbookPart?.Workbook == null ||
            workbook.WorkbookPart.WorkbookStylesPart?.Stylesheet?.CellFormats?.Elements<S.CellFormat>()
                .FirstOrDefault()?.NumberFormatId?.Value is uint defaultFormat && defaultFormat != 0 ||
            workbook.WorkbookPart.WorksheetParts.Any(part => part.Worksheet == null ||
                part.Worksheet.Descendants<S.Cell>().Any(cell => cell.StyleIndex?.Value is uint index && index != 0) ||
                part.Worksheet.Descendants<S.Row>().Any(row => row.StyleIndex?.Value is uint index && index != 0) ||
                part.Worksheet.Descendants<S.Column>().Any(column => column.Style?.Value is uint index && index != 0)))
            throw new NotSupportedException("The source-linked chart axis requires workbook number-format projection.");
        return "General";
    }

    private static bool HasOnlyGeneralValueCaches(ChartPart chartPart) {
        C.PlotArea? plot = chartPart.ChartSpace?.GetFirstChild<C.Chart>()?.PlotArea;
        if (plot == null) return false;
        C.NumberingCache[] caches = plot.Descendants<C.NumberingCache>()
            .Where(cache => cache.Parent?.Parent is C.Values or C.YValues).ToArray();
        return caches.Length > 0 && caches.All(cache =>
            cache.FormatCode?.Text is not string code || string.IsNullOrWhiteSpace(code) ||
            string.Equals(code, "General", StringComparison.OrdinalIgnoreCase));
    }
}
