using BenchmarkDotNet.Running;
using OfficeIMO.Reader.Benchmarks.Comparison;

if (args.Length > 0 && string.Equals(args[0], "evidence", StringComparison.OrdinalIgnoreCase)) {
    return await ReaderComparisonCommand.RunAsync(args.Skip(1).ToArray()).ConfigureAwait(false);
}

if (args.Length > 0 && string.Equals(args[0], "office-evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    return ReaderComparisonCommand.RunOfficeProbe(args.Skip(1).ToArray());
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
return 0;
