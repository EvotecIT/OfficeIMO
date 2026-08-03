using OfficeIMO.PowerPoint.Benchmarks;

if (args.Length == 2 && string.Equals(args[0], "--visual-corpus",
        StringComparison.OrdinalIgnoreCase)) {
    return PowerPointRoadmapVisualCorpus.Create(args[1]);
}

if (args.Length > 0 && string.Equals(args[0], "--probe",
        StringComparison.OrdinalIgnoreCase)) {
    return PowerPointBaselineRunner.RunProbe(args.Skip(1).ToArray());
}

return PowerPointBaselineRunner.RunBaseline(args);
