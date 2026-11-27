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
    // --flows-ab [rounds] [--no-check Flow,Flow]: the flows named after --no-check skip their setup check.
    var uncheckedFlows = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    int rounds = 200;
    for (int index = 1; index < args.Length; index++) {
        if (string.Equals(args[index], "--no-check", StringComparison.OrdinalIgnoreCase) && index + 1 < args.Length) {
            foreach (string flow in args[++index].Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)) uncheckedFlows.Add(flow);
        } else {
            rounds = int.Parse(args[index], System.Globalization.CultureInfo.InvariantCulture);
        }
    }
    return PowerPointWithoutSdkFlowBenchmarks.RunInterleaved(rounds, uncheckedFlows);
}

if (args.Length > 0 && string.Equals(args[0], "--content-flows", StringComparison.OrdinalIgnoreCase)) {
    // --content-flows [rounds]: charts, SmartArt, layout and image export flows.
    return PowerPointContentFlowBenchmarks.RunInterleaved(args.Length > 1 ? int.Parse(args[1], System.Globalization.CultureInfo.InvariantCulture) : 50);
}

if (args.Length > 0 && string.Equals(args[0], "--probe",
        StringComparison.OrdinalIgnoreCase)) {
    return PowerPointBaselineRunner.RunProbe(args.Skip(1).ToArray());
}

return PowerPointBaselineRunner.RunBaseline(args);
