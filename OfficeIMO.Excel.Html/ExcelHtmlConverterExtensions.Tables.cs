using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    // Both HTML profiles project cells or paint; neither stores native table definitions.
    private static void ReportNamedTableLoss(IEnumerable<ExcelTableInfo> tables, IList<HtmlDiagnostic> diagnostics) {
        foreach (ExcelTableInfo table in tables) {
            diagnostics.Add(new HtmlDiagnostic(
                "OfficeIMO.Excel.Html",
                HtmlConversionDiagnosticCodes.ContentOmitted,
                "Named Excel table '" + table.Name + "' was exported without its native table definition.",
                HtmlDiagnosticSeverity.Warning,
                "excel:table:" + table.SheetName + "/" + table.Name,
                detail: "range=" + table.Range + "; metadata=name, columns, table style, filter, totals",
                lossKind: OfficeConversionLossKind.Omission));
        }
    }
}
