using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Reports;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Table;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Engine lanes for the workbook shapes the PowerShell module produces: a table from mixed-type object rows, the
/// same table with fitted columns, and the table with a chart over two of its columns. Each lane runs on .NET
/// Framework and .NET, checks the saved workbook and prints a Markdown table.
/// </summary>
/// <remarks>
/// <c>--compare-engine-lanes [--rows 5000,10000,25000] [--iterations 9] [--warmups 3] [--lane autofit,table-save,table-chart]</c>
/// </remarks>
internal static class ExcelEngineLanes {
    static ExcelEngineLanes() => ExcelPackage.License.SetNonCommercialOrganization("OfficeIMO benchmarks");

    private static readonly string[] ColumnNames = {
        "Id", "Name", "Department", "Region", "Score", "IsEnabled", "Created", "Owner", "TicketCount", "Notes"
    };

    // Region and Score sit next to each other so every library charts one contiguous block.
    private const int RegionColumn = 4;
    private const int ScoreColumn = 5;

    private sealed class Lane {
        internal Lane(string name, bool autoFit, bool chart, bool writer = false) {
            Name = name;
            AutoFit = autoFit;
            Chart = chart;
            Writer = writer;
        }

        internal string Name { get; }
        internal bool AutoFit { get; }
        internal bool Chart { get; }

        /// <summary>OfficeIMO writes the rows through its streaming workbook writer, as Export-OfficeExcel does.</summary>
        internal bool Writer { get; }
    }

    private static readonly Lane[] Lanes = {
        new Lane("table-save", autoFit: false, chart: false),
        new Lane("autofit", autoFit: true, chart: false),
        new Lane("autofit-writer", autoFit: true, chart: false, writer: true),
        new Lane("table-chart", autoFit: false, chart: true)
    };

    private static readonly (string Name, Func<DataTable, Lane, byte[]> Write, Func<Lane, bool> Supports)[] Libraries = {
        ("OfficeIMO", WriteOfficeImo, _ => true),
        ("EPPlus", WriteEpPlus, _ => true),
        // ClosedXML does not create charts.
        ("ClosedXML", WriteClosedXml, lane => !lane.Chart)
    };

    internal static int Run(string[] args) {
        int[] rows = ReadList(args, "--rows", new[] { 5000, 10000, 25000 });
        int iterations = ReadList(args, "--iterations", new[] { 9 })[0];
        int warmups = ReadList(args, "--warmups", new[] { 3 })[0];
        string[]? selected = ReadOption(args, "--lane")?.Split(',');
        var lanes = Lanes.Where(lane => selected == null || selected.Contains(lane.Name, StringComparer.OrdinalIgnoreCase)).ToArray();

        string runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;
        Console.WriteLine($"Runtime: {runtime}, {(Environment.Is64BitProcess ? "64" : "32")}-bit; iterations {iterations}, warmups {warmups}; median ms");
        Console.WriteLine();
        Console.WriteLine("| Lane | Rows | " + string.Join(" | ", Libraries.Select(library => library.Name)) + " | OfficeIMO vs best |");
        Console.WriteLine("| --- | ---: | " + string.Join(" | ", Libraries.Select(_ => "---:")) + " | --- |");
        int failures = 0;
        foreach (int rowCount in rows) {
            DataTable table = CreateTable(rowCount);
            foreach (Lane lane in lanes) {
                var medians = new List<double?>();
                foreach (var library in Libraries) {
                    if (!library.Supports(lane)) {
                        medians.Add(null);
                        continue;
                    }

                    string? problem = Validate(library.Write(table, lane), rowCount, lane);
                    if (problem != null) {
                        Console.Error.WriteLine($"{library.Name} {lane.Name} {rowCount}: {problem}");
                        failures++;
                        medians.Add(null);
                        continue;
                    }

                    medians.Add(Measure(() => library.Write(table, lane), iterations, warmups));
                }

                double? officeImo = medians[0];
                double? best = medians.Skip(1).Where(value => value.HasValue).Select(value => value!.Value).DefaultIfEmpty().Min();
                string verdict = officeImo == null || best == null || best == 0 ? "n/a"
                    : officeImo <= best ? Ratio(best.Value / officeImo.Value) + "x faster"
                    : Ratio(officeImo.Value / best.Value) + "x slower";
                Console.WriteLine($"| {lane.Name} | {rowCount} | " + string.Join(" | ", medians.Select(Format)) + $" | {verdict} |");
            }
        }

        return failures == 0 ? 0 : 1;
    }

    private static string Ratio(double value) => value.ToString("0.00", CultureInfo.InvariantCulture);

    private static string Format(double? value) => value == null ? "n/a" : value.Value.ToString("0.0", CultureInfo.InvariantCulture);

    private static double Measure(Func<byte[]> write, int iterations, int warmups) {
        for (int i = 0; i < warmups; i++) write();
        var samples = new double[iterations];
        for (int i = 0; i < iterations; i++) {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            var watch = Stopwatch.StartNew();
            write();
            samples[i] = watch.Elapsed.TotalMilliseconds;
        }

        Array.Sort(samples);
        return samples[samples.Length / 2];
    }

    /// <summary>Mixed-type object rows as PowerShell objects arrive: every column typed as object.</summary>
    private static DataTable CreateTable(int rowCount) {
        var table = new DataTable("Data");
        foreach (string name in ColumnNames) table.Columns.Add(name, typeof(object));
        string[] regions = { "NA", "EU", "APAC", "LATAM" };
        var start = new DateTime(2024, 1, 1);
        for (int i = 1; i <= rowCount; i++) {
            table.Rows.Add(
                i,
                "Server-" + i.ToString("000000", CultureInfo.InvariantCulture),
                "Department-" + (i % 25).ToString(CultureInfo.InvariantCulture),
                regions[i % 4],
                Math.Round((i * 1.137) % 1000, 3),
                i % 3 != 0,
                start.AddMinutes(i),
                "owner" + (i % 250).ToString(CultureInfo.InvariantCulture) + "@example.test",
                i % 17,
                "Benchmark row " + i.ToString(CultureInfo.InvariantCulture));
        }

        return table;
    }

    private static string DataRange(int rowCount) => "A1:J" + (rowCount + 1).ToString(CultureInfo.InvariantCulture);

    private static byte[] WriteOfficeImo(DataTable table, Lane lane) {
        if (lane.Writer) return WriteOfficeImoStreaming(table);
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            var sheet = document.AddSheet("Data");
            sheet.AddTable(table, new ExcelTableOptions {
                Name = "Data",
                Style = ExcelTableStyle.TableStyleMedium9,
                AutoFit = false
            }, row: 1, column: 1);
            if (lane.AutoFit) sheet.AutoFitColumnsFor(Enumerable.Range(1, ColumnNames.Length));
            if (lane.Chart) {
                sheet.AddDashboardChart(new ExcelDashboardChartOptions {
                    Range = "D1:E" + (table.Rows.Count + 1).ToString(CultureInfo.InvariantCulture),
                    Row = 2,
                    Column = 12,
                    Width = OfficeLength.Pixels(640),
                    Height = OfficeLength.Pixels(360),
                    ChartType = OfficeChartType.ColumnClustered,
                    Title = "Score by region"
                });
            }

            document.Save(stream);
        }

        return stream.ToArray();
    }

    private static byte[] WriteOfficeImoStreaming(DataTable table) {
        using var stream = new MemoryStream();
        var options = new ExcelWriterOptions {
            CreateTable = true,
            TableName = "Data",
            TableStyle = ExcelTableStyle.TableStyleMedium9,
            AutoFit = true,
            BoldHeaders = true,
            FreezeHeader = true,
            UseSharedStrings = false,
            UseCellValueNumberFormats = true
        };
        using (var writer = ExcelDocument.CreateWriter(stream, options, leaveOpen: true)) {
            var rows = writer.BeginTable("Data");
            rows.DeclareColumns(ColumnNames);
            foreach (DataRow row in table.Rows) rows.WriteRow(ColumnNames, row.ItemArray);
            rows.Complete();
        }

        return stream.ToArray();
    }

    private static byte[] WriteEpPlus(DataTable table, Lane lane) {
        using var package = new ExcelPackage();
        var sheet = package.Workbook.Worksheets.Add("Data");
        sheet.Cells[1, 1].LoadFromDataTable(table, true, TableStyles.Medium9);
        int lastRow = table.Rows.Count + 1;
        if (lane.AutoFit) sheet.Cells[DataRange(table.Rows.Count)].AutoFitColumns();
        if (lane.Chart) {
            var chart = sheet.Drawings.AddChart("ScoreByRegion", eChartType.ColumnClustered);
            chart.Title.Text = "Score by region";
            chart.SetPosition(1, 0, 11, 0);
            chart.SetSize(640, 360);
            chart.Series.Add(sheet.Cells[2, ScoreColumn, lastRow, ScoreColumn], sheet.Cells[2, RegionColumn, lastRow, RegionColumn]);
        }

        return package.GetAsByteArray();
    }

    private static byte[] WriteClosedXml(DataTable table, Lane lane) {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Data");
        var inserted = sheet.Cell(1, 1).InsertTable(table, "Data");
        inserted.Theme = XLTableTheme.TableStyleMedium9;
        if (lane.AutoFit) sheet.Columns(1, ColumnNames.Length).AdjustToContents();
        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>Checks the saved workbook holds every row, a fitted width per column and the chart when the lane asks.</summary>
    private static string? Validate(byte[] package, int rowCount, Lane lane) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        var sheetEntry = archive.Entries.FirstOrDefault(entry => entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase));
        if (sheetEntry == null) return "no worksheet part";
        string xml;
        using (var reader = new StreamReader(sheetEntry.Open())) xml = reader.ReadToEnd();
        int rows = Regex.Matches(xml, "<(?:x:)?row[ >]").Count;
        if (rows != rowCount + 1) return $"expected {rowCount + 1} rows, found {rows}";
        if (archive.Entries.All(entry => !entry.FullName.StartsWith("xl/tables/", StringComparison.OrdinalIgnoreCase))) return "no table part";
        if (lane.AutoFit) {
            var widths = Regex.Matches(xml, "<(?:x:)?col [^>]*min=\"(\\d+)\"[^>]*max=\"(\\d+)\"[^>]*width=\"([\\d.]+)\"|<(?:x:)?col [^>]*width=\"([\\d.]+)\"[^>]*min=\"(\\d+)\"[^>]*max=\"(\\d+)\"");
            var fitted = new HashSet<int>();
            foreach (Match match in widths) {
                bool first = match.Groups[1].Success;
                int min = int.Parse(first ? match.Groups[1].Value : match.Groups[5].Value, CultureInfo.InvariantCulture);
                int max = int.Parse(first ? match.Groups[2].Value : match.Groups[6].Value, CultureInfo.InvariantCulture);
                double width = double.Parse(first ? match.Groups[3].Value : match.Groups[4].Value, CultureInfo.InvariantCulture);
                if (width > 0) for (int column = min; column <= max; column++) fitted.Add(column);
            }

            if (Enumerable.Range(1, ColumnNames.Length).Any(column => !fitted.Contains(column))) return "not every column has a fitted width";
        }

        if (lane.Chart && archive.Entries.All(entry => entry.FullName.IndexOf("/charts/chart", StringComparison.OrdinalIgnoreCase) < 0)) {
            return "no chart part";
        }

        return null;
    }

    private static string? ReadOption(string[] args, string name) {
        int index = Array.FindIndex(args, arg => string.Equals(arg, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static int[] ReadList(string[] args, string name, int[] fallback) {
        string? value = ReadOption(args, name);
        return value == null ? fallback : value.Split(',').Select(item => int.Parse(item, CultureInfo.InvariantCulture)).ToArray();
    }
}

#if NETFRAMEWORK
internal static class ExcelEngineLanesProgram {
    private static int Main(string[] args) {
        var rest = args.Length > 0 && string.Equals(args[0], "--compare-engine-lanes", StringComparison.OrdinalIgnoreCase)
            ? args.Skip(1).ToArray()
            : args;
        return ExcelEngineLanes.Run(rest);
    }
}
#endif
