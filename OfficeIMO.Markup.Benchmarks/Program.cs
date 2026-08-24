using BenchmarkDotNet.Running;
using OfficeIMO.Markup;
using OfficeIMO.Markup.Benchmarks;

if (args.Length > 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
    Environment.ExitCode = OfficeMarkupEvidenceRunner.RunProbe(args[1..]);
    return;
}

if (args.Any(argument => string.Equals(argument, "--verify-budgets", StringComparison.OrdinalIgnoreCase))) {
    Environment.ExitCode = OfficeMarkupEvidenceRunner.RunEvidence(args, verifyBudgets: true);
    return;
}

if (args.Any(argument => string.Equals(argument, "--evidence", StringComparison.OrdinalIgnoreCase))) {
    Environment.ExitCode = OfficeMarkupEvidenceRunner.RunEvidence(args, verifyBudgets: false);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "probe", StringComparison.OrdinalIgnoreCase)) {
    OfficeMarkupBenchmarkFixture fixture = OfficeMarkupBenchmarkCorpus.Get(args.Length > 1 ? args[1] : "Large");
    int repetitions = args.Length > 2 ? int.Parse(args[2], System.Globalization.CultureInfo.InvariantCulture) : 1;
    OfficeMarkupParseResult? result = null;
    for (int repetition = 0; repetition < repetitions; repetition++) {
        result = OfficeMarkupParser.Parse(fixture.Source, OfficeMarkupBenchmarkValidation.OfficeOptions);
    }
    Console.WriteLine(result!.Document.Blocks.Count);
    return;
}

if (args.Length > 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
    OfficeMarkupBenchmarkValidation.ValidateAll();
    return;
}

BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
