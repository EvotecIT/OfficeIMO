using OfficeIMO.PowerPoint.Benchmarks;

if (args.Length == 2 && string.Equals(args[0], "--visual-corpus",
        StringComparison.OrdinalIgnoreCase)) {
    return PowerPointRoadmapVisualCorpus.Create(args[1]);
}

if (args.Length > 0 && string.Equals(args[0], "--native-text-layout",
        StringComparison.OrdinalIgnoreCase)) {
    return PowerPointNativeTextLayoutBenchmark.Run(args.Skip(1).ToArray());
}

if (args.Length > 0 && string.Equals(args[0], "--bdn", StringComparison.OrdinalIgnoreCase)) {
    BenchmarkDotNet.Running.BenchmarkSwitcher.FromAssembly(typeof(PowerPointModelCreateBenchmarks).Assembly).Run(args.Skip(1).ToArray());
    return 0;
}

if (args.Length > 0 && string.Equals(args[0], "--flows-ab", StringComparison.OrdinalIgnoreCase)) {
    return PowerPointWithoutSdkFlowBenchmarks.RunInterleaved(args.Length > 1 ? int.Parse(args[1], System.Globalization.CultureInfo.InvariantCulture) : 200);
}

if (args.Length > 0 && string.Equals(args[0], "--probe",
        StringComparison.OrdinalIgnoreCase)) {
    return PowerPointBaselineRunner.RunProbe(args.Skip(1).ToArray());
}

return PowerPointBaselineRunner.RunBaseline(args);
