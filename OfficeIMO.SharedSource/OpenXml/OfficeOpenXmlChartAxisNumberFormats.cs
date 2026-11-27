using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal.OpenXml;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Qualifies the small source-linked workbook format subset that chart snapshots can render exactly.</summary>
internal static class OfficeOpenXmlChartAxisNumberFormats {
    /// <param name="chartPart">The chart part.</param>
    /// <param name="axis">The linked axis; its own values decide, such as the X values of a scatter X axis. None checks
    /// every series' values.</param>
    internal static string ResolveSourceLinkedGeneral(IOfficeChartPart chartPart, OfficeXmlNode axis = default) {
        // A source-linked value axis takes the number format of the plotted cells, which the
        // chart caches beside every value. General caches need no workbook projection.
        if (!axis.IsNone ? HasOnlyGeneralCodes(OfficeOpenXmlChartSecondaryAxis.ReadCachedValueFormats(axis).ToArray())
                : HasOnlyGeneralValueCaches(chartPart)) return "General";
        IOfficeChartEmbeddedPackage? embedded = OfficeOpenXmlChartWriter.GetSharedEmbeddedWorkbook(chartPart);
        if (embedded == null)
            throw new NotSupportedException("The source-linked chart axis has no qualified embedded workbook.");
        using var source = embedded.OpenRead();
        byte[] bytes = OfficeOpenXmlChartWorkbookSecurity.ReadAndValidate(source);
        // The embedded workbook is still read through the SDK: bytes in, a verdict out.
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

    private static bool HasOnlyGeneralCodes(string?[] codes) =>
        codes.Length > 0 && codes.All(code => string.IsNullOrWhiteSpace(code) || string.Equals(code, "General", StringComparison.OrdinalIgnoreCase));

    private static bool HasOnlyGeneralValueCaches(IOfficeChartPart chartPart) {
        OfficeXmlNode plot = chartPart.ChartSpace.Child(NsC, "chart").Child(NsC, "plotArea");
        if (plot.IsNone) return false;
        OfficeXmlNode[] caches = plot.ContentDescendants(NsC, "numCache")
            .Where(cache => cache.Parent.Parent is var owner && (owner.IsChart("val") || owner.IsChart("yVal"))).ToArray();
        return caches.Length > 0 && caches.All(cache =>
            cache.Child(NsC, "formatCode") is var format && (format.IsNone || string.IsNullOrWhiteSpace(format.InnerText) ||
            string.Equals(format.InnerText, "General", StringComparison.OrdinalIgnoreCase)));
    }

    // SDK callers
    internal static string ResolveSourceLinkedGeneral(ChartPart chartPart, C.ValueAxis? axis = null) =>
        ResolveSourceLinkedGeneral(new OfficeSdkChartPart(chartPart), OfficeXmlNode.FromSdk(axis));
}
