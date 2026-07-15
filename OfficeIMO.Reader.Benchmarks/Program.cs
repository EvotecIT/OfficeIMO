using BenchmarkDotNet.Running;
using OfficeIMO.Reader.Benchmarks.Comparison;

if (args.Length > 0 && string.Equals(args[0], "compare", StringComparison.OrdinalIgnoreCase)) {
    return await ReaderComparisonCommand.RunAsync(args.Skip(1).ToArray()).ConfigureAwait(false);
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;