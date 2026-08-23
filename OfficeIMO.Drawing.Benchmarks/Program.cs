using BenchmarkDotNet.Running;
using OfficeIMO.Drawing.Benchmarks;

if (args.Length == 1 && args[0].Equals("--validate", StringComparison.OrdinalIgnoreCase)) {
    ImageBenchmarkValidator.Validate(Console.Out);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
