using BenchmarkDotNet.Attributes;
using System.Globalization;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Compares linear and binary lookup over identical sorted inputs and complete outputs.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("FormulaCalculation")]
    public class FormulaLookupScaleBenchmarks {
        private ExcelDocument _document = null!;
        private ExcelSheet _report = null!;
        private int _calculated;

        [Params(512, 8192)]
        public int KeyCount { get; set; }

        [Params(1, 2)]
        public int SearchMode { get; set; }

        [GlobalSetup]
        public void DescribeInput() => BenchmarkInput.WriteDescription();

        [IterationSetup]
        public void Setup() {
            _calculated = -1;
            _document = ExcelDocument.Create();
            var source = _document.AddWorksheet("Source");
            for (int key = 1; key <= KeyCount; key++) {
                source.CellValue(key, 1, key);
                source.CellValue(key, 2, key * 10);
            }
            _report = _document.AddWorksheet("Report");
            for (int index = 0; index < 256; index++)
                _report.CellFormula(index + 1, 1,
                    $"XLOOKUP({ExpectedKey(index)},Source!A1:A{KeyCount},Source!B1:B{KeyCount},,0,{SearchMode})");
        }

        [Benchmark]
        public int Calculate() => _calculated = _document.Calculate();

        [IterationCleanup]
        public void ValidateAndDispose() {
            try {
                if (_calculated < 0) return;
                if (_calculated != 256) throw new InvalidOperationException($"Calculated {_calculated} of 256 lookups.");
                for (int index = 0; index < 256; index++) {
                    string expected = (ExpectedKey(index) * 10).ToString(CultureInfo.InvariantCulture);
                    if (!_report.TryGetCachedFormulaValue(index + 1, 1, out string? actual) || actual != expected)
                        throw new InvalidOperationException($"Lookup {index} returned '{actual}', expected '{expected}'.");
                }
            } finally { _document.Dispose(); }
        }

        private int ExpectedKey(int index) => index * 37 % KeyCount + 1;
    }
}
