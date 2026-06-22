using System.Data;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Builds a dashboard table and optional chart from a data table.
        /// </summary>
        /// <param name="data">Tabular data to write.</param>
        /// <param name="options">Dashboard layout and styling options.</param>
        /// <returns>Dashboard table range and generated chart metadata.</returns>
        public ExcelDashboardResult AddDashboard(DataTable data, ExcelDashboardOptions? options = null) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            var opts = options ?? new ExcelDashboardOptions();
            if (data.Columns.Count == 0) throw new ArgumentException("Dashboard data must contain at least one column.", nameof(data));
            if (opts.TableRow <= 0) throw new ArgumentOutOfRangeException(nameof(ExcelDashboardOptions.TableRow));
            if (opts.TableColumn <= 0) throw new ArgumentOutOfRangeException(nameof(ExcelDashboardOptions.TableColumn));
            ValidateDashboardLayout(opts);

            WriteDashboardHeader(opts);

            string? tableName = string.IsNullOrWhiteSpace(opts.TableName)
                ? (string.IsNullOrWhiteSpace(data.TableName) ? null : data.TableName)
                : opts.TableName;
            string tableRange = InsertDataTableAsTable(
                data,
                startRow: opts.TableRow,
                startColumn: opts.TableColumn,
                includeHeaders: true,
                tableName: tableName,
                style: opts.TableStyle,
                includeAutoFilter: opts.IncludeAutoFilter);
            string? actualTableName = ResolveTableNameByRange(tableRange) ?? tableName;

            if (opts.AutoFit) {
                AutoFitColumnsFor(Enumerable.Range(opts.TableColumn, data.Columns.Count));
            }

            ExcelChart? chart = null;
            if (opts.AddChart) {
                int chartRow = opts.ChartRow ?? opts.TableRow;
                int chartColumn = opts.ChartColumn ?? opts.TableColumn + data.Columns.Count + 2;
                chart = AddDashboardChart(new ExcelDashboardChartOptions {
                    Preset = opts.ChartPreset,
                    Range = tableRange,
                    Row = chartRow,
                    Column = chartColumn,
                    Title = string.IsNullOrWhiteSpace(opts.ChartTitle) ? opts.Title : opts.ChartTitle,
                    HasHeaders = true
                });
            }

            return new ExcelDashboardResult(tableRange, actualTableName, chart);
        }

        private string? ResolveTableNameByRange(string tableRange) {
            if (string.IsNullOrWhiteSpace(tableRange)) {
                return null;
            }

            return _worksheetPart.TableDefinitionParts
                .Select(part => part.Table)
                .FirstOrDefault(table => string.Equals(table?.Reference?.Value, tableRange, StringComparison.OrdinalIgnoreCase))
                is { } table
                    ? table.Name?.Value ?? table.DisplayName?.Value
                    : null;
        }

        private void WriteDashboardHeader(ExcelDashboardOptions options) {
            if (!string.IsNullOrWhiteSpace(options.Title)) {
                CellValue(1, 1, options.Title!);
                CellBold(1, 1, true);
            }

            if (!string.IsNullOrWhiteSpace(options.Subtitle)) {
                CellValue(2, 1, options.Subtitle!);
            }
        }

        private static void ValidateDashboardLayout(ExcelDashboardOptions options) {
            int headerLastRow = 0;
            if (!string.IsNullOrWhiteSpace(options.Title)) {
                headerLastRow = 1;
            }

            if (!string.IsNullOrWhiteSpace(options.Subtitle)) {
                headerLastRow = 2;
            }

            if (headerLastRow > 0 && options.TableRow <= headerLastRow) {
                throw new ArgumentException("Dashboard table row overlaps the configured title or subtitle rows.", nameof(options));
            }
        }
    }
}
