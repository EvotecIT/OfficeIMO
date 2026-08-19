using System.Globalization;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class ExcelHtmlConverterExtensions {
    private static void AppendPivotInventory(StringBuilder body, IEnumerable<ExcelPivotTableInfo> pivots,
        IList<HtmlDiagnostic> diagnostics) {
        List<ExcelPivotTableInfo> pivotList = pivots.ToList();
        if (pivotList.Count == 0) return;

        body.Append("<section class=\"officeimo-feature officeimo-pivots\"><h3>Pivot tables</h3>")
            .Append("<div class=\"officeimo-diagnostic\" data-officeimo-loss=\"simplified\">")
            .Append("Pivot definitions are exposed as inert review metadata; refresh, drill, caches, slicers, and timelines remain native workbook behavior.")
            .Append("</div><ul class=\"officeimo-feature-list\">");
        foreach (ExcelPivotTableInfo pivot in pivotList) {
            diagnostics.Add(new HtmlDiagnostic(
                "OfficeIMO.Excel.Html",
                HtmlConversionDiagnosticCodes.ExcelPivotReviewApproximated,
                "Pivot table '" + pivot.Name + "' was exported as inert review metadata; refresh, drill, caches, slicers, and timelines remain native workbook behavior.",
                HtmlDiagnosticSeverity.Warning,
                "excel:pivot:" + pivot.Name,
                lossKind: OfficeConversionLossKind.Approximation));
            body.Append("<li class=\"officeimo-feature-item\" data-officeimo-feature=\"pivot-table\" data-officeimo-pivot-name=\"")
                .Append(OfficeHtmlText.EscapeAttribute(pivot.Name))
                .Append("\" data-officeimo-sheet=\"")
                .Append(OfficeHtmlText.EscapeAttribute(pivot.SheetName))
                .Append("\" data-officeimo-cache-id=\"")
                .Append(pivot.CacheId.ToString(CultureInfo.InvariantCulture))
                .Append("\">")
                .Append("<span class=\"officeimo-feature-label\">")
                .Append(OfficeHtmlText.Escape(pivot.Name))
                .Append("</span><dl class=\"officeimo-feature-meta\">");
            AppendPivotValue(body, "Location", pivot.Location);
            AppendPivotValue(body, "Source", JoinReference(pivot.SourceSheet, pivot.SourceRange));
            AppendPivotValue(body, "Layout", pivot.Layout.ToString());
            AppendPivotValue(body, "Style", pivot.PivotStyle);
            AppendPivotValue(body, "Rows", string.Join(", ", pivot.RowFields));
            AppendPivotValue(body, "Columns", string.Join(", ", pivot.ColumnFields));
            AppendPivotValue(body, "Filters", string.Join(", ", pivot.PageFields));
            AppendPivotValue(body, "Values", string.Join(", ", pivot.DataFields.Select(field =>
                string.IsNullOrWhiteSpace(field.DisplayName) ? field.FieldName : field.DisplayName)));
            body.Append("</dl></li>");
        }
        body.Append("</ul></section>");
    }

    private static void AppendPivotValue(StringBuilder body, string label, string? value) {
        if (string.IsNullOrWhiteSpace(value)) return;
        body.Append("<dt>").Append(OfficeHtmlText.Escape(label)).Append("</dt><dd>")
            .Append(OfficeHtmlText.Escape(value!)).Append("</dd>");
    }

    private static string? JoinReference(string? sheet, string? range) {
        if (string.IsNullOrWhiteSpace(sheet)) return range;
        if (string.IsNullOrWhiteSpace(range)) return sheet;
        return sheet + "!" + range;
    }
}
