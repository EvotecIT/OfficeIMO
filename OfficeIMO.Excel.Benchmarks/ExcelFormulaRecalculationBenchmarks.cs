using System.Globalization;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using OfficeIMO.Benchmarks;
using OfficeOpenXml;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Recalculates every formula in an order-report workbook: per-row arithmetic, IF, VLOOKUP, INDEX/MATCH with
/// IFERROR, text concatenation and sheet totals. The <c>ConditionalTotals</c> shape adds a SUMIFS per row over the
/// whole order column, which makes the pass quadratic in the row count. Each library loads the same
/// workbook once and recalculates it in full per invocation, so parsing, dispatch and reference resolution are
/// measured rather than package loading. Setup checks that every lane computes the same totals.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance", "FormulaRecalculation")]
public class ExcelFormulaRecalculationBenchmarks {
    private static readonly string[] Regions = ["North", "South", "East", "West", "Central", "Online", "Export", "Retail"];
    private int FormulasPerRow => ConditionalTotals ? 7 : 6;

    [Params(2_000)]
    public int Rows { get; set; }

    [Params("RowFormulas", "ConditionalTotals")]
    public string Shape { get; set; } = "RowFormulas";

    private bool ConditionalTotals => Shape == "ConditionalTotals";

    private ExcelDocument? _officeImo;
    private MemoryStream? _officeImoStream;
    private XLWorkbook? _closedXml;
    private ExcelPackage? _epPlus;

    [GlobalSetup]
    public void Setup() {
        byte[] workbook = CreateWorkbook(Rows, ConditionalTotals);
        _officeImoStream = new MemoryStream();
        _officeImoStream.Write(workbook, 0, workbook.Length);
        _officeImoStream.Position = 0;
        _officeImo = ExcelDocument.Load(_officeImoStream);
        _closedXml = new XLWorkbook(new MemoryStream(workbook));

        int lastRow = Rows + 1;
        string detail = $"{Shape}: {Rows} order rows x {FormulasPerRow} formulas plus 4 totals; the totals and two per-row columns compared with expected values";
        Observation expected = Expected(Rows, ConditionalTotals);
        ExcelOutputValidation.Check("OfficeIMO", "Default", detail, () => {
            int calculated = OfficeIMO();
            int formulas = Rows * FormulasPerRow + 4;
            if (calculated != formulas) throw new InvalidDataException($"OfficeIMO calculated {calculated} of {formulas} formulas.");
            ExcelSheet sheet = _officeImo.Sheets.First(s => s.Name == "Orders");
            expected.Check(new Observation(
                OfficeImoNumber(sheet, lastRow + 2, 5),
                OfficeImoNumber(sheet, lastRow + 2, 6),
                OfficeImoNumber(sheet, lastRow + 2, 7),
                OfficeImoNumber(sheet, lastRow + 2, 8),
                SumColumn(row => OfficeImoNumber(sheet, row, 8), lastRow),
                ConditionalTotals ? SumColumn(row => OfficeImoNumber(sheet, row, 10), lastRow) : 0,
                OfficeImoText(sheet, 2, 11)));
        });
        ExcelOutputValidation.Check("ClosedXML", "Default", detail, () => {
            ClosedXML();
            IXLWorksheet sheet = _closedXml.Worksheet("Orders");
            expected.Check(new Observation(
                sheet.Cell(lastRow + 2, 5).GetDouble(),
                sheet.Cell(lastRow + 2, 6).GetDouble(),
                sheet.Cell(lastRow + 2, 7).GetDouble(),
                sheet.Cell(lastRow + 2, 8).GetDouble(),
                SumColumn(row => sheet.Cell(row, 8).GetDouble(), lastRow),
                ConditionalTotals ? SumColumn(row => sheet.Cell(row, 10).GetDouble(), lastRow) : 0,
                sheet.Cell(2, 11).GetString()));
        });
        if (BenchmarkLibrarySelection.IsEnabled("epplus")) {
            _epPlus = new ExcelPackage(new MemoryStream(workbook));
            ExcelOutputValidation.Check("EPPlus", "Default", detail, () => {
                EPPlus();
                ExcelWorksheet sheet = _epPlus.Workbook.Worksheets["Orders"];
                expected.Check(new Observation(
                    Convert.ToDouble(sheet.Cells[lastRow + 2, 5].Value, CultureInfo.InvariantCulture),
                    Convert.ToDouble(sheet.Cells[lastRow + 2, 6].Value, CultureInfo.InvariantCulture),
                    Convert.ToDouble(sheet.Cells[lastRow + 2, 7].Value, CultureInfo.InvariantCulture),
                    Convert.ToDouble(sheet.Cells[lastRow + 2, 8].Value, CultureInfo.InvariantCulture),
                    SumColumn(row => Convert.ToDouble(sheet.Cells[row, 8].Value, CultureInfo.InvariantCulture), lastRow),
                    ConditionalTotals ? SumColumn(row => Convert.ToDouble(sheet.Cells[row, 10].Value, CultureInfo.InvariantCulture), lastRow) : 0,
                    Convert.ToString(sheet.Cells[2, 11].Value, CultureInfo.InvariantCulture) ?? string.Empty));
            });
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        _officeImo?.Dispose();
        _officeImoStream?.Dispose();
        _closedXml?.Dispose();
        _epPlus?.Dispose();
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public int OfficeIMO() => _officeImo!.RecalculateSupportedFormulas();

    [Benchmark]
    [BenchmarkCategory("library:ClosedXML", "case:Default")]
    public int ClosedXML() {
        _closedXml!.RecalculateAllFormulas();
        return 0;
    }

    [Benchmark]
    [BenchmarkCategory("library:EPPlus", "case:Default")]
    public int EPPlus() {
        _epPlus!.Workbook.Calculate();
        return 0;
    }

    /// <summary>Builds the shared workbook with OfficeIMO, without cached formula values.</summary>
    internal static byte[] CreateWorkbook(int rows, bool conditionalTotals) {
        int lastRow = rows + 1;
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create(output)) {
            ExcelSheet orders = document.AddSheet("Orders");
            ExcelSheet rates = document.AddSheet("Rates");
            for (int index = 0; index < Regions.Length; index++) {
                rates.Cell(index + 1, 1).SetValue(Regions[index]);
                rates.Cell(index + 1, 2).SetValue(Rate(index));
            }

            orders.Cell(1, 1).SetValue("Id");
            orders.Cell(1, 2).SetValue("Region");
            orders.Cell(1, 3).SetValue("Quantity");
            orders.Cell(1, 4).SetValue("Price");
            for (int row = 2; row <= lastRow; row++) {
                int order = row - 1;
                orders.Cell(row, 1).SetValue(order);
                orders.Cell(row, 2).SetValue(Regions[order % Regions.Length]);
                orders.Cell(row, 3).SetValue(Quantity(order));
                orders.Cell(row, 4).SetValue(Price(order));
                orders.Cell(row, 5).SetFormula($"C{row}*D{row}");
                orders.Cell(row, 6).SetFormula($"IF(E{row}>500,\"High\",\"Low\")");
                orders.Cell(row, 7).SetFormula($"VLOOKUP(B{row},Rates!$A$1:$B$8,2,FALSE)");
                orders.Cell(row, 8).SetFormula($"ROUND(E{row}*G{row},2)");
                orders.Cell(row, 9).SetFormula($"IFERROR(INDEX(Rates!$B$1:$B$8,MATCH(B{row},Rates!$A$1:$A$8,0)),0)");
                if (conditionalTotals) orders.Cell(row, 10).SetFormula($"SUMIFS($C$2:$C${lastRow},$B$2:$B${lastRow},B{row})");
                orders.Cell(row, 11).SetFormula($"LEFT(B{row},3)&\"-\"&A{row}");
            }

            int totals = lastRow + 2;
            orders.Cell(totals, 5).SetFormula($"SUM(E2:E{lastRow})");
            orders.Cell(totals, 6).SetFormula($"COUNTIF(F2:F{lastRow},\"High\")");
            orders.Cell(totals, 7).SetFormula($"AVERAGE(G2:G{lastRow})");
            orders.Cell(totals, 8).SetFormula($"MAX(H2:H{lastRow})");
            document.Save(output);
        }

        return output.ToArray();
    }

    private static double Rate(int region) => 1 + region * 0.05;
    private static int Quantity(int order) => 1 + order * 7 % 40;
    private static double Price(int order) => 2.5 + order * 13 % 50;

    private static Observation Expected(int rows, bool conditionalTotals) {
        double sumE = 0, high = 0, sumG = 0, maxH = double.MinValue, sumH = 0;
        var quantityByRegion = new double[Regions.Length];
        for (int order = 1; order <= rows; order++) quantityByRegion[order % Regions.Length] += Quantity(order);
        double sumJ = 0;
        for (int order = 1; order <= rows; order++) {
            int region = order % Regions.Length;
            double e = Quantity(order) * Price(order);
            sumE += e;
            if (e > 500) high++;
            sumG += Rate(region);
            double h = Math.Round(e * Rate(region), 2, MidpointRounding.AwayFromZero);
            sumH += h;
            maxH = Math.Max(maxH, h);
            sumJ += quantityByRegion[region];
        }

        return new Observation(sumE, high, sumG / rows, maxH, sumH, conditionalTotals ? sumJ : 0, Regions[1].Substring(0, 3) + "-1");
    }

    private static double SumColumn(Func<int, double> value, int lastRow) {
        double sum = 0;
        for (int row = 2; row <= lastRow; row++) sum += value(row);
        return sum;
    }

    private static double OfficeImoNumber(ExcelSheet sheet, int row, int column) =>
        sheet.TryGetCachedFormulaValue(row, column, out string? value) && value != null
            ? double.Parse(value, CultureInfo.InvariantCulture)
            : throw new InvalidDataException($"OfficeIMO left R{row}C{column} without a cached value.");

    private static string OfficeImoText(ExcelSheet sheet, int row, int column) =>
        sheet.TryGetCachedFormulaValue(row, column, out string? value) && value != null
            ? value
            : throw new InvalidDataException($"OfficeIMO left R{row}C{column} without a cached value.");

    private readonly record struct Observation(
        double TotalAmount, double HighCount, double AverageRate, double MaxWeighted,
        double WeightedSum, double RegionQuantitySum, string FirstLabel) {
        internal void Check(Observation actual) {
            Compare(nameof(TotalAmount), TotalAmount, actual.TotalAmount);
            Compare(nameof(HighCount), HighCount, actual.HighCount);
            Compare(nameof(AverageRate), AverageRate, actual.AverageRate);
            Compare(nameof(MaxWeighted), MaxWeighted, actual.MaxWeighted);
            Compare(nameof(WeightedSum), WeightedSum, actual.WeightedSum);
            Compare(nameof(RegionQuantitySum), RegionQuantitySum, actual.RegionQuantitySum);
            if (!string.Equals(FirstLabel, actual.FirstLabel, StringComparison.Ordinal))
                throw new InvalidDataException($"First label '{actual.FirstLabel}' differs from the expected '{FirstLabel}'.");
        }

        private static void Compare(string name, double expected, double actual) {
            if (double.IsNaN(actual) || double.IsInfinity(actual) || Math.Abs(expected - actual) > 1e-6 * Math.Max(1, Math.Abs(expected)))
                throw new InvalidDataException($"{name} {actual} differs from the expected {expected}.");
        }
    }
}
