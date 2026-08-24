using BenchmarkDotNet.Running;

namespace OfficeIMO.Provenance.Benchmarks;

internal static class Program {
    private static int Main(string[] args) {
        if (args.Length != 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
            return ProvenanceEvidenceRunner.RunProbe(args[1..]);
        }

        if (args.Length != 0 && string.Equals(args[0], "evidence", StringComparison.OrdinalIgnoreCase)) {
            return ProvenanceEvidenceRunner.Run(args[1..]);
        }

        if (args.Length != 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
            ProvenanceBenchmarkValidation.ValidateAll(writeSummary: true);
            return 0;
        }

        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
        return 0;
    }
}
