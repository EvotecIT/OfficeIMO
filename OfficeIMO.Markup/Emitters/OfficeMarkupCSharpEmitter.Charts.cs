namespace OfficeIMO.Markup;

public sealed partial class OfficeMarkupCSharpEmitter {
    private static void EmitPowerPointChart(OfficeMarkupChartBlock chart, string slideVariable, StringBuilder sb, int chartIndex) {
        EmitPlacementComment(chart.Placement, sb);
        if (chart.Data.Count < 2 || chart.Data[0].Count < 2) {
            sb.AppendLine($"// Add {chart.ChartType} chart from source {CsString(chart.Source ?? string.Empty)}.");
            return;
        }

        var dataVariable = $"chartData{chartIndex}";
        EmitPowerPointChartData(chart, dataVariable, sb);
        var kind = ToPowerPointChartKind(chart.ChartType);
        var chartVariable = $"chart{chartIndex}";
        sb.AppendLine($"var {chartVariable} = {slideVariable}.AddChart(OfficeChartKind.{kind}, {dataVariable});");
        if (!string.IsNullOrWhiteSpace(chart.Title)) {
            sb.AppendLine($"{chartVariable}.SetTitle({CsString(chart.Title!)});");
        }

        EmitPowerPointChartSemanticOptions(chart, chartVariable, sb);
    }

    private static void EmitPowerPointChartSemanticOptions(OfficeMarkupChartBlock chart, string chartVariable, StringBuilder sb) {
        if (GetAttribute(chart.Attributes, "category-title", "categoryTitle", "x-title", "xTitle", "x-axis-title", "xAxisTitle") is { Length: > 0 } categoryTitle) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisTitle({CsString(categoryTitle)});");
        }

        if (GetAttribute(chart.Attributes, "value-title", "valueTitle", "y-title", "yTitle", "y-axis-title", "yAxisTitle") is { Length: > 0 } valueTitle) {
            sb.AppendLine($"{chartVariable}.SetValueAxisTitle({CsString(valueTitle)});");
        }

        if (GetAttribute(chart.Attributes, "category-format", "categoryFormat", "x-format", "xFormat", "category-number-format", "categoryNumberFormat") is { Length: > 0 } categoryFormat) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisNumberFormat({CsString(categoryFormat)});");
        }

        if (GetAttribute(chart.Attributes, "value-format", "valueFormat", "y-format", "yFormat", "value-number-format", "valueNumberFormat") is { Length: > 0 } valueFormat) {
            sb.AppendLine($"{chartVariable}.SetValueAxisNumberFormat({CsString(valueFormat)});");
        }

        var legend = GetAttribute(chart.Attributes, "legend", "legend-position", "legendPosition");
        if (!string.IsNullOrWhiteSpace(legend)) {
            var legendValue = legend!;
            var normalized = NormalizeToken(legendValue);
            if (normalized is "false" or "none" or "hidden" or "off") {
                sb.AppendLine($"{chartVariable}.HideLegend();");
            } else if (TryGetLegendPositionIdentifier(legendValue, out var legendPosition)) {
                sb.AppendLine($"{chartVariable}.SetLegend(C.LegendPositionValues.{legendPosition});");
            }
        }

        var labels = GetAttribute(chart.Attributes, "labels", "data-labels", "dataLabels");
        if (!string.IsNullOrWhiteSpace(labels)) {
            if (IsTruthy(labels!)) {
                sb.AppendLine($"{chartVariable}.SetDataLabels(showValue: true, showCategoryName: false, showSeriesName: false, showLegendKey: false, showPercent: false);");
                var labelPosition = GetAttribute(chart.Attributes, "label-position", "labelPosition", "data-label-position", "dataLabelPosition");
                if (TryGetDataLabelPositionIdentifier(labelPosition, out var dataLabelPosition)) {
                    sb.AppendLine($"{chartVariable}.SetDataLabelPosition(C.DataLabelPositionValues.{dataLabelPosition});");
                }

                var labelFormat = GetAttribute(chart.Attributes, "label-format", "labelFormat", "data-label-format", "dataLabelFormat");
                if (!string.IsNullOrWhiteSpace(labelFormat)) {
                    sb.AppendLine($"{chartVariable}.SetDataLabelNumberFormat({CsString(labelFormat!)});");
                }
            } else {
                sb.AppendLine($"{chartVariable}.ClearDataLabels();");
            }
        }

        var gridlines = GetAttribute(chart.Attributes, "gridlines");
        var valueGridlines = GetAttribute(chart.Attributes, "value-gridlines", "valueGridlines", "y-gridlines", "yGridlines") ?? gridlines;
        var categoryGridlines = GetAttribute(chart.Attributes, "category-gridlines", "categoryGridlines", "x-gridlines", "xGridlines");
        if (!string.IsNullOrWhiteSpace(valueGridlines)) {
            sb.AppendLine(IsTruthy(valueGridlines!)
                ? $"{chartVariable}.SetValueAxisGridlines(showMajor: true, showMinor: false);"
                : $"{chartVariable}.ClearValueAxisGridlines();");
        }

        if (!string.IsNullOrWhiteSpace(categoryGridlines)) {
            sb.AppendLine(IsTruthy(categoryGridlines!)
                ? $"{chartVariable}.SetCategoryAxisGridlines(showMajor: true, showMinor: false);"
                : $"{chartVariable}.ClearCategoryAxisGridlines();");
        }
    }

    private static void EmitPowerPointChartData(OfficeMarkupChartBlock chart, string variableName, StringBuilder sb) {
        var headers = chart.Data[0];
        var categories = chart.Data.Skip(1).Select(row => row.Count > 0 ? row[0] : string.Empty).ToList();
        sb.AppendLine($"var {variableName} = new OfficeChartData(");
        sb.AppendLine($"    new[] {{ {string.Join(", ", categories.Select(CsString))} }},");
        sb.AppendLine("    new[] {");
        for (int seriesIndex = 1; seriesIndex < headers.Count; seriesIndex++) {
            var values = chart.Data.Skip(1).Select(row => NumericLiteral(row.Count > seriesIndex ? row[seriesIndex] : "0"));
            var comma = seriesIndex == headers.Count - 1 ? string.Empty : ",";
            sb.AppendLine($"        new OfficeChartSeries({CsString(headers[seriesIndex])}, new[] {{ {string.Join(", ", values)} }}){comma}");
        }

        sb.AppendLine("    });");
    }

    private static void EmitWordChart(OfficeMarkupChartBlock chart, string documentVariable, StringBuilder sb) {
        if (chart.Data.Count < 2 || chart.Data[0].Count < 2) {
            sb.AppendLine($"// Add {chart.ChartType} chart from source {CsString(chart.Source ?? string.Empty)}.");
            return;
        }

        sb.AppendLine($"// Add {chart.ChartType} chart {CsString(chart.Title ?? string.Empty)} from inline data.");
        sb.AppendLine($"// Word chart APIs can consume the same categories and series represented by this AST node.");
    }

    private static void EmitExcelChart(OfficeMarkupChartBlock chart, StringBuilder sb, int chartIndex) {
        var row = 1;
        var column = 6;
        chart.Attributes.TryGetValue("cell", out var cell);
        var (chartSheetExpression, placementCell) = ResolveWorkbookTarget(chart.Sheet, cell ?? string.Empty);
        if (TryParseCellAddress(placementCell, out var parsedRow, out var parsedColumn)) {
            row = parsedRow;
            column = parsedColumn;
        }

        var width = TryParseInt(chart.Placement?.Width, out var parsedWidth) ? parsedWidth : 640;
        var height = TryParseInt(chart.Placement?.Height, out var parsedHeight) ? parsedHeight : 360;
        var chartType = ToExcelChartType(chart.ChartType);
        var chartVariable = $"chart{chartIndex}";
        if (!string.IsNullOrWhiteSpace(chart.Source)) {
            EmitExcelChartFromSource(chart, sb, chartIndex, chartVariable, chartSheetExpression, row, column, width, height, chartType);
            EmitExcelChartSemanticOptions(chart, chartVariable, sb);
            return;
        }

        if (chart.Data.Count < 2 || chart.Data[0].Count < 2) {
            sb.AppendLine($"// Add {chart.ChartType} chart from source {CsString(chart.Source ?? string.Empty)}.");
            return;
        }

        var dataVariable = $"chartData{chartIndex}";
        EmitExcelChartData(chart, dataVariable, sb);
        sb.AppendLine($"var {chartVariable} = {chartSheetExpression}.AddChart({dataVariable}, row: {row}, column: {column}, widthPixels: {width}, heightPixels: {height}, type: ExcelChartType.{chartType}, title: {CsString(chart.Title ?? string.Empty)});");
        EmitExcelChartSemanticOptions(chart, chartVariable, sb);
    }

    private static void EmitExcelChartFromSource(
        OfficeMarkupChartBlock chart,
        StringBuilder sb,
        int chartIndex,
        string chartVariable,
        string chartSheetExpression,
        int row,
        int column,
        int width,
        int height,
        string chartType) {
        var source = chart.Source ?? string.Empty;
        if (TrySplitSheetQualifiedReference(source, out var sourceSheetName, out var localSource)) {
            var sourceSheetVariable = $"chartSourceSheet{chartIndex}";
            sb.AppendLine($"var {sourceSheetVariable} = GetOrAddSheet({CsString(sourceSheetName)});");
            EmitExcelChartDataRangeFromSource(sb, chartIndex, sourceSheetVariable, localSource);
            sb.AppendLine($"var {chartVariable} = {chartSheetExpression}.AddChart(chartDataRange{chartIndex}, row: {row}, column: {column}, widthPixels: {width}, heightPixels: {height}, type: ExcelChartType.{chartType}, title: {CsString(chart.Title ?? string.Empty)});");
            return;
        }

        if (source.IndexOf(':') >= 0) {
            sb.AppendLine($"var {chartVariable} = {chartSheetExpression}.AddChartFromRange({CsString(source)}, row: {row}, column: {column}, widthPixels: {width}, heightPixels: {height}, type: ExcelChartType.{chartType}, title: {CsString(chart.Title ?? string.Empty)});");
        } else {
            sb.AppendLine($"var {chartVariable} = {chartSheetExpression}.AddChartFromTable({CsString(source)}, row: {row}, column: {column}, widthPixels: {width}, heightPixels: {height}, type: ExcelChartType.{chartType}, title: {CsString(chart.Title ?? string.Empty)});");
        }
    }

    private static void EmitExcelChartDataRangeFromSource(StringBuilder sb, int chartIndex, string sourceSheetVariable, string localSource) {
        if (localSource.IndexOf(':') >= 0) {
            sb.AppendLine($"var (chartR1_{chartIndex}, chartC1_{chartIndex}, chartR2_{chartIndex}, chartC2_{chartIndex}) = A1.ParseRange({CsString(localSource)});");
        } else {
            sb.AppendLine($"var chartSourceRange{chartIndex} = {sourceSheetVariable}.GetTableRange({CsString(localSource)}) ?? throw new global::System.InvalidOperationException({CsString("Chart source table was not found.")});");
            sb.AppendLine($"var (chartR1_{chartIndex}, chartC1_{chartIndex}, chartR2_{chartIndex}, chartC2_{chartIndex}) = A1.ParseRange(chartSourceRange{chartIndex});");
        }

        sb.AppendLine($"var chartDataRange{chartIndex} = new ExcelChartDataRange({sourceSheetVariable}.Name, chartR1_{chartIndex}, chartC1_{chartIndex}, chartR2_{chartIndex} - chartR1_{chartIndex}, chartC2_{chartIndex} - chartC1_{chartIndex}, hasHeaderRow: true);");
    }

    private static void EmitExcelChartSemanticOptions(OfficeMarkupChartBlock chart, string chartVariable, StringBuilder sb) {
        if (GetAttribute(chart.Attributes, "category-title", "categoryTitle", "x-title", "xTitle", "x-axis-title", "xAxisTitle") is { Length: > 0 } categoryTitle) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisTitle({CsString(categoryTitle)});");
        }

        if (GetAttribute(chart.Attributes, "value-title", "valueTitle", "y-title", "yTitle", "y-axis-title", "yAxisTitle") is { Length: > 0 } valueTitle) {
            sb.AppendLine($"{chartVariable}.SetValueAxisTitle({CsString(valueTitle)});");
        }

        if (GetAttribute(chart.Attributes, "category-format", "categoryFormat", "x-format", "xFormat", "category-number-format", "categoryNumberFormat") is { Length: > 0 } categoryFormat) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisNumberFormat({CsString(categoryFormat)});");
        }

        if (GetAttribute(chart.Attributes, "value-format", "valueFormat", "y-format", "yFormat", "value-number-format", "valueNumberFormat") is { Length: > 0 } valueFormat) {
            sb.AppendLine($"{chartVariable}.SetValueAxisNumberFormat({CsString(valueFormat)});");
        }

        var legend = GetAttribute(chart.Attributes, "legend", "legend-position", "legendPosition");
        if (!string.IsNullOrWhiteSpace(legend)) {
            var legendValue = legend!;
            var normalized = NormalizeToken(legendValue);
            if (normalized is "false" or "none" or "hidden" or "off") {
                sb.AppendLine($"{chartVariable}.HideLegend();");
            } else if (TryGetLegendPositionIdentifier(legendValue, out var legendPosition)) {
                sb.AppendLine($"{chartVariable}.SetLegend(C.LegendPositionValues.{legendPosition});");
            }
        }

        var labels = GetAttribute(chart.Attributes, "labels", "data-labels", "dataLabels");
        if (!string.IsNullOrWhiteSpace(labels) && IsTruthy(labels!)) {
            var labelPosition = GetAttribute(chart.Attributes, "label-position", "labelPosition", "data-label-position", "dataLabelPosition");
            var labelFormat = GetAttribute(chart.Attributes, "label-format", "labelFormat", "data-label-format", "dataLabelFormat");
            var positionExpression = TryGetDataLabelPositionIdentifier(labelPosition, out var dataLabelPosition)
                ? $"C.DataLabelPositionValues.{dataLabelPosition}"
                : "null";
            var numberFormatExpression = !string.IsNullOrWhiteSpace(labelFormat) ? CsString(labelFormat!) : "null";
            sb.AppendLine($"{chartVariable}.SetDataLabels(showValue: true, showCategoryName: false, showSeriesName: false, showLegendKey: false, showPercent: false, position: {positionExpression}, numberFormat: {numberFormatExpression});");
        }

        var gridlines = GetAttribute(chart.Attributes, "gridlines");
        var valueGridlines = GetAttribute(chart.Attributes, "value-gridlines", "valueGridlines", "y-gridlines", "yGridlines") ?? gridlines;
        var categoryGridlines = GetAttribute(chart.Attributes, "category-gridlines", "categoryGridlines", "x-gridlines", "xGridlines");
        if (!string.IsNullOrWhiteSpace(valueGridlines)) {
            sb.AppendLine($"{chartVariable}.SetValueAxisGridlines(showMajor: {BoolLiteral(IsTruthy(valueGridlines!))}, showMinor: false);");
        }

        if (!string.IsNullOrWhiteSpace(categoryGridlines)) {
            sb.AppendLine($"{chartVariable}.SetCategoryAxisGridlines(showMajor: {BoolLiteral(IsTruthy(categoryGridlines!))}, showMinor: false);");
        }
    }

    private static void EmitExcelChartData(OfficeMarkupChartBlock chart, string variableName, StringBuilder sb) {
        var headers = chart.Data[0];
        var categories = chart.Data.Skip(1).Select(row => row.Count > 0 ? row[0] : string.Empty).ToList();
        sb.AppendLine($"var {variableName} = new ExcelChartData(");
        sb.AppendLine($"    new[] {{ {string.Join(", ", categories.Select(CsString))} }},");
        sb.AppendLine("    new[] {");
        for (int seriesIndex = 1; seriesIndex < headers.Count; seriesIndex++) {
            var values = chart.Data.Skip(1).Select(row => NumericLiteral(row.Count > seriesIndex ? row[seriesIndex] : "0"));
            var comma = seriesIndex == headers.Count - 1 ? string.Empty : ",";
            sb.AppendLine($"        new ExcelChartSeries({CsString(headers[seriesIndex])}, new[] {{ {string.Join(", ", values)} }}){comma}");
        }

        sb.AppendLine("    });");
    }
}
