using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Maps the object-export contract onto shared tabular projection and native semantic slide blocks.</summary>
internal static partial class PowerPointObjectExport {
    internal static PowerPointDeckPlan CreatePlan(IReadOnlyCollection<object?> input, string title, string? subtitle,
        string[]? properties, string? category, string[]? values, OfficeChartType kind, string? chartTitle,
        string? groupBy, bool noSummary, bool noTable, int? maximumRows, CultureInfo culture, bool chartTypeBound) {
        var normalization = new PowerShellObjectNormalizerOptions { Culture = culture };
        using var table = ExcelTabularInputService.ToDataTable(TableInputCollector.RequireRows(input, "InputObject"),
            normalizerOptions: normalization);
        if (table.Rows.Count == 0 || table.Columns.Count == 0)
            throw new ArgumentException("Provide at least one data row with a property.", nameof(input));
        var columns = SelectColumns(table, properties, "Property");
        bool chartRequested = category != null || values != null;
        if (chartRequested && (string.IsNullOrWhiteSpace(category) || values == null || values.Length == 0))
            throw new ArgumentException("Supply both Category and Value to add a chart.");
        if (!chartRequested && (chartTypeBound || chartTitle != null))
            throw new ArgumentException("ChartType and ChartTitle require Category and Value.");
        if (noTable && !chartRequested)
            throw new ArgumentException("NoTable requires Category and Value; a title alone has no exported content.");
        if (!Enum.IsDefined(typeof(OfficeChartType), kind))
            throw new ArgumentException("Choose a supported shared chart kind.", nameof(kind));
        if (groupBy != null) RequireColumn(table, groupBy, "GroupBy");

        var plan = new PowerPointDeckPlan().AddContent(title, Array.Empty<PowerPointLayoutBlock>(),
            PowerPointLayout.Title, subtitle);
        if (!noSummary && !noTable) {
            DataColumn? grouping = groupBy != null ? RequireColumn(table, groupBy, "GroupBy") :
                category != null ? RequireColumn(table, category, "Category") : FindSummaryColumn(table, normalization);
            if (grouping != null) {
                var metrics = table.Rows.Cast<DataRow>()
                    .GroupBy(row => PowerShellObjectNormalizer.NormalizeCellText(row[grouping], normalization), StringComparer.OrdinalIgnoreCase)
                    .Select(group => (PowerPointLayoutBlock)new PowerPointLayoutKpi(
                        group.Count().ToString("N0", culture), group.Key.Length == 0 ? "(empty)" : group.Key)).ToArray();
                plan.AddContent("At a glance", metrics, PowerPointLayout.Kpi);
            }
        }
        if (!noTable) {
            var rows = table.Rows.Cast<DataRow>().Select(row => (IReadOnlyList<string>)columns.Select(column =>
                PowerShellObjectNormalizer.NormalizeCellText(row[column], normalization)).ToArray()).ToArray();
            plan.AddContent(title, new[] { new PowerPointLayoutTable(columns.Select(column => column.ColumnName).ToArray(), rows) {
                MaximumRows = maximumRows, IncludeHeaders = true
            } }, PowerPointLayout.TitleAndContent);
        }
        if (chartRequested) {
            DataColumn categoryColumn = RequireColumn(table, category!, "Category");
            var seriesColumns = SelectColumns(table, values, "Value");
            if ((kind == OfficeChartType.Pie || kind == OfficeChartType.Doughnut) && seriesColumns.Length != 1)
                throw new ArgumentException("Pie and Doughnut charts require exactly one Value series.", nameof(values));
            if (kind == OfficeChartType.Bubble)
                throw new ArgumentException("Bubble charts require per-point sizes, which Category and Value do not supply.", nameof(kind));
            double[]? x = kind == OfficeChartType.Scatter
                ? table.Rows.Cast<DataRow>().Select(row => ChartNumber(row[categoryColumn], categoryColumn.ColumnName, culture)).ToArray()
                : null;
            var labels = table.Rows.Cast<DataRow>().Select((row, index) => x == null
                ? PowerShellObjectNormalizer.NormalizeCellText(row[categoryColumn], normalization)
                : x[index].ToString("R", CultureInfo.InvariantCulture)).ToArray();
            var series = seriesColumns.Select(column => new OfficeSeries(column.ColumnName,
                table.Rows.Cast<DataRow>().Select(row => ChartNumber(row[column], column.ColumnName, culture)), x)).ToArray();
            string heading = chartTitle ?? categoryColumn.ColumnName;
            plan.AddContent(heading, new[] { new PowerPointLayoutChart(kind, new OfficeChartData(labels, series)) {
                Title = heading
            } }, PowerPointLayout.TitleAndContent);
        }
        return plan;
    }

    private static DataColumn[] SelectColumns(DataTable table, string[]? names, string parameter) {
        if (names == null) return table.Columns.Cast<DataColumn>().ToArray();
        if (names.Length == 0) throw new ArgumentException(parameter + " requires at least one column.", parameter);
        var columns = names.Select(name => RequireColumn(table, name, parameter)).ToArray();
        if (columns.Distinct().Count() != columns.Length)
            throw new ArgumentException(parameter + " cannot repeat a column.", parameter);
        return columns;
    }

    private static DataColumn RequireColumn(DataTable table, string name, string parameter) =>
        table.Columns.Cast<DataColumn>().FirstOrDefault(column =>
            !string.IsNullOrWhiteSpace(name) && column.ColumnName.Equals(name, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException("Unknown " + parameter + " column '" + name + "'.", parameter);

    private static DataColumn? FindSummaryColumn(DataTable table, PowerShellObjectNormalizerOptions normalization) {
        foreach (DataColumn column in table.Columns) {
            var values = table.Rows.Cast<DataRow>().Select(row => row[column]).Where(value => value != DBNull.Value).ToArray();
            if (values.Length == 0 || values.Any(value => value is not string)) continue;
            int distinct = values.Select(value => PowerShellObjectNormalizer.NormalizeCellText(value, normalization))
                .Distinct(StringComparer.OrdinalIgnoreCase).Take(7).Count();
            if (distinct >= 2 && distinct <= 6) return column;
        }
        return null;
    }

    /// <summary>Binds chart values without turning blanks, Boolean values or non-finite numbers into data points.</summary>
    private static double ChartNumber(object value, string column, CultureInfo culture) {
        bool numeric = value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal;
        double result;
        if (numeric) result = Convert.ToDouble(value, culture);
        else if (value is string text && double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, culture, out result)) { }
        else throw new ArgumentException("Chart column '" + column + "' requires numeric values.");
        if (double.IsNaN(result) || double.IsInfinity(result))
            throw new ArgumentException("Chart column '" + column + "' requires finite numeric values.");
        return result;
    }
}
