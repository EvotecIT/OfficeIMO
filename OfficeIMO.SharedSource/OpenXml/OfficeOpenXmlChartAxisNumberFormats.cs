using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.OpenXml.Internal;

/// <summary>Qualifies the small source-linked workbook format subset that chart snapshots can render exactly.</summary>
internal static class OfficeOpenXmlChartAxisNumberFormats {
    internal static string ResolveSourceLinkedGeneral(ChartPart chartPart) {
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
                part.Worksheet.Descendants<S.Cell>().Any(cell => cell.StyleIndex?.Value is uint index && index != 0)))
            throw new NotSupportedException("The source-linked chart axis requires workbook number-format projection.");
        return "General";
    }
}
