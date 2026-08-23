using BenchmarkDotNet.Running;
using OfficeIMO.Drawing.Benchmarks;

if (args.Length == 1 && args[0].Equals("--validate", StringComparison.OrdinalIgnoreCase)) {
    ImageBenchmarkValidator.Validate(Console.Out);
    return;
}

if (args.Length == 2 && args[0].Equals("--resampling-previews", StringComparison.OrdinalIgnoreCase)) {
    ImageResamplingEvidence.WritePreviews(args[1], Console.Out);
    return;
}

if (args.Length >= 1 && args[0].Equals("--memory-evidence", StringComparison.OrdinalIgnoreCase)) {
    ImagePeakMemoryEvidence.Validate(Console.Out, args.Skip(1).ToArray());
    return;
}

if (args.Length == 4 && args[0].Equals("--memory-worker", StringComparison.OrdinalIgnoreCase)) {
    ImagePeakMemoryEvidence.RunWorker(args[1], args[2], args[3], Console.Out);
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
