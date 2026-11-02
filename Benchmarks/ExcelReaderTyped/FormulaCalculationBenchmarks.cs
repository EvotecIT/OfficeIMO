using BenchmarkDotNet.Attributes;
using System.Globalization;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Calculates a category report and validates every formula cache outside measurement.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("FormulaCalculation")]
    public class FormulaCalculationBenchmarks {
        private ExcelDocument _document = null!;
        private ExcelSheet _orders = null!;
        private readonly List<(int Row, int Column, string Value)> _expected = [];
        private int _calculated;

        [Params("IndexHit", "IndexMiss", "VlookupMiss", "XlookupMiss")]
        public string Scenario { get; set; } = "IndexHit";

        [IterationSetup]
        public void Setup() {
            _expected.Clear();
            _calculated = -1;
            _document = ExcelDocument.Create();
            var products = _document.AddWorksheet("Products");
            for (int p = 1; p <= 200; p++) {
                products.CellValue(p + 1, 1, "P" + p);
                products.CellValue(p + 1, 2, "C" + p % 23);
            }
            _orders = _document.AddWorksheet("Orders");
            int[] counts = new int[23];
            double[] sums = new double[23];
            for (int i = 1; i <= 1000; i++) {
                int row = i + 1, product = (i - 1) % 200 + 1;
                bool missing = Scenario != "IndexHit" && i % 97 == 0;
                _orders.CellValue(row, 1, missing ? "PX" : "P" + product);
                string expression = Scenario switch {
                    "VlookupMiss" => $"VLOOKUP(A{row},Products!$A$2:$B$201,2,FALSE)",
                    "XlookupMiss" => $"XLOOKUP(A{row},Products!$A$2:$A$201,Products!$B$2:$B$201)",
                    _ => $"INDEX(Products!$B$2:$B$201,MATCH(A{row},Products!$A$2:$A$201,0))"
                };
                _orders.CellFormula(row, 2, $"IFERROR({expression},\"Unknown\")");
                double amount = i % 11 + 1;
                _orders.CellValue(row, 3, amount);
                _orders.CellFormula(row, 4, $"C{row}*2");
                _expected.Add((row, 2, missing ? "Unknown" : "C" + product % 23));
                Expect(row, 4, amount * 2);
                if (!missing) {
                    counts[product % 23]++;
                    sums[product % 23] += amount * 2;
                }
            }
            for (int category = 0; category < 23; category++) {
                int row = 1003 + category;
                _orders.CellFormula(row, 2, $"COUNTIFS(B2:B1001,\"C{category}\")");
                _orders.CellFormula(row, 4, $"SUMIFS(D2:D1001,B2:B1001,\"C{category}\")");
                Expect(row, 2, counts[category]);
                Expect(row, 4, sums[category]);
            }
            _orders.CellFormula(1027, 2, "SUM(B1003:B1025)");
            _orders.CellFormula(1027, 4, "SUM(D1003:D1025)");
            Expect(1027, 2, counts.Sum());
            Expect(1027, 4, sums.Sum());
        }

        [Benchmark]
        public int Calculate() => _calculated = _document.Calculate();

        [IterationCleanup]
        public void ValidateAndDispose() {
            try {
                // BDN can invoke cleanup for an overhead iteration without executing the workload.
                if (_calculated < 0) return;
                if (_calculated != _expected.Count)
                    throw new InvalidOperationException($"Calculated {_calculated} of {_expected.Count} formulas.");
                foreach (var item in _expected) {
                    if (!_orders.TryGetCachedFormulaValue(item.Row, item.Column, out string? value) || value != item.Value)
                        throw new InvalidOperationException($"Incorrect formula cache at ({item.Row},{item.Column}): '{value}', expected '{item.Value}'.");
                }
            } finally {
                _document.Dispose();
            }
        }

        private void Expect(int row, int column, double value) =>
            _expected.Add((row, column, value.ToString(CultureInfo.InvariantCulture)));
    }
}
