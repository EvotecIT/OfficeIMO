using BenchmarkDotNet.Running;
using OfficeIMO.Drawing.Benchmarks.Comparisons;

if (args.Length == 1 && args[0].Equals("--validate", StringComparison.OrdinalIgnoreCase)) {
    ImageComparisonValidation.Validate(Console.Out);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
