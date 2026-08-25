using BenchmarkDotNet.Running;

namespace OfficeIMO.Visio.Benchmarks;

internal static class Program {
    private static int Main(string[] args) {
        if (args.Length != 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
            return VisioEvidenceRunner.RunProbe(args[1..]);
        }

        if (args.Length != 0 && string.Equals(args[0], "evidence", StringComparison.OrdinalIgnoreCase)) {
            return VisioEvidenceRunner.Run(args[1..]);
        }

        if (args.Length != 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
            VisioBenchmarkValidation.ValidateAll(writeSummary: true);
            return 0;
        }

        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
        return 0;
    }
}
