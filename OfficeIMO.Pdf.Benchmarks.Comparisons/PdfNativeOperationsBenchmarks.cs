using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons {
    /// <summary>Native reading and page operations over independently generated, validated 100-page inputs.</summary>
    [MemoryDiagnoser]
    public class PdfNativeOperationsBenchmarks {
        /// <summary>Independent producer used outside the measured operation.</summary>
        [Params("IText", "MigraDoc")]
        public string Producer { get; set; } = "IText";

        /// <summary>Complete operation, including source parsing and output serialization where applicable.</summary>
        [Params("ReadText", "ReadStructured", "Select", "Merge", "Split", "SplitSelections")]
        public string Operation { get; set; } = "ReadText";

        private byte[] _source = null!;
        private byte[][] _sources = null!;
        private int[] _selection = null!;
        private PdfBenchmarkScenario _scenario = null!;
        private readonly List<(PdfBenchmarkScenario Scenario, byte[] Bytes)> _mergeInputs = new();

        /// <summary>Creates inputs and independently validates every expected page, narrative and table row.</summary>
        [GlobalSetup]
        public void Setup() {
            _scenario = new PdfBenchmarkScenario(PdfBenchmarkScale.High, "Native PDF operations", 100, 4, 1);
            _source = Generate(_scenario);
            PdfBenchmarkValidation.ValidateGenerated(_source, _scenario, Producer);
            _selection = Enumerable.Range(0, 25).Select(i => 100 - i * 4).ToArray();
            if (Operation == "Merge") {
                for (int i = 1; i <= 25; i++) {
                    var scenario = _scenario with { PageCount = 4, DocumentNumber = i };
                    byte[] bytes = Generate(scenario);
                    PdfBenchmarkValidation.ValidateGenerated(bytes, scenario, Producer);
                    _mergeInputs.Add((scenario, bytes));
                }
                _sources = _mergeInputs.Select(item => item.Bytes).ToArray();
            }
            object result = Execute();
            if (Operation.StartsWith("Read", StringComparison.Ordinal)) {
                PdfBenchmarkValidation.ValidateRead(PdfBenchmarkValidation.Observe(100, (string)result), _scenario, Operation);
            } else {
                byte[][] outputs = result is byte[] bytes ? new[] { bytes } : (byte[][])result;
                IReadOnlyList<IReadOnlyList<PdfExpectedPage>> expected = Operation switch {
                    "Select" => new[] { (IReadOnlyList<PdfExpectedPage>)_selection.Select(p => PdfBenchmarkValidation.ExpectedPage(_scenario, p)).ToArray() },
                    "Merge" => new[] { (IReadOnlyList<PdfExpectedPage>)_mergeInputs.SelectMany(item => Enumerable.Range(1, 4).Select(p => PdfBenchmarkValidation.ExpectedPage(item.Scenario, p))).ToArray() },
                    _ => Enumerable.Range(1, 100).Select(p => (IReadOnlyList<PdfExpectedPage>)new[] { PdfBenchmarkValidation.ExpectedPage(_scenario, p) }).ToArray()
                };
                PdfManipulationValidation.Validate(outputs, expected, Operation);
            }
        }

        private byte[] Generate(PdfBenchmarkScenario scenario) => Producer == "IText" ? ITextPdfGenerator.Generate(scenario) : MigraDocPdfGenerator.Generate(scenario);

        /// <summary>Runs one operation and returns its observable result.</summary>
        [Benchmark]
        public object Execute() => Operation switch {
            "ReadText" => PdfReadDocument.Open(_source).ExtractText(),
            "ReadStructured" => PdfDocument.Load(_source).Read(new PdfReadOptions { Profile = PdfReadProfile.Structured }).Text,
            "Select" => PdfDocument.Load(_source).Pages.Extract(_selection).ToBytes(),
            "Merge" => PdfDocument.Merge(_sources.Select(source => PdfDocument.Load(source))).ToBytes(),
            "Split" => PdfDocument.Load(_source).Pages.Split().Select(document => document.ToBytes()).ToArray(),
            "SplitSelections" => PdfDocument.Load(_source).Pages.Split(Enumerable.Range(1, 100).Select(page => PdfPageSelection.From(page)).ToArray()).Select(document => document.ToBytes()).ToArray(),
            _ => throw new InvalidOperationException()
        };
    }
}

