using OfficeIMO.Excel.Benchmarks;
using BenchmarkDotNet.Running;

if (args.Length >= 2 && string.Equals(args[0], "--snapshot", StringComparison.OrdinalIgnoreCase)) {
    string outputPath = ExcelBenchmarkSnapshotRunner.WriteSnapshot(args[1]);
    Console.WriteLine($"Excel benchmark snapshot written to '{outputPath}'.");
    return;
}

if (args.Length >= 2 && string.Equals(args[0], "--profile-write", StringComparison.OrdinalIgnoreCase)) {
    string outputPath = ExcelWriteProfileRunner.WriteProfile(args[1]);
    Console.WriteLine($"Excel write profile written to '{outputPath}'.");
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
