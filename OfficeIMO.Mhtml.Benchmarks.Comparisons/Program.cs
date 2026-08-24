using BenchmarkDotNet.Running;

namespace OfficeIMO.Mhtml.Benchmarks.Comparisons;

internal static class Program {
    private static int Main(string[] args) {
        if (args.Length != 0 && string.Equals(args[0], "--evidence-probe", StringComparison.OrdinalIgnoreCase)) {
            return MhtmlEvidenceRunner.RunProbe(args[1..]);
        }

        if (args.Length != 0 && string.Equals(args[0], "evidence", StringComparison.OrdinalIgnoreCase)) {
            return MhtmlEvidenceRunner.Run(args[1..]);
        }

        if (args.Length != 0 && string.Equals(args[0], "validate", StringComparison.OrdinalIgnoreCase)) {
            foreach (MhtmlComparisonReport report in MhtmlComparisonValidation.ValidateAll()) {
                Console.WriteLine(
                    $"{report.Scale,-6} {report.ResourceCount,3} resources | " +
                    $"{report.DecodedResourceBytes,10:N0} decoded bytes | " +
                    $"OfficeIMO {report.OfficeIMOOutputBytes,10:N0} bytes | " +
                    $"MimeKit {report.MimeKitOutputBytes,10:N0} bytes");
            }
            return 0;
        }

        BenchmarkSwitcher.FromAssembly(typeof(Program).Assembly).Run(args);
        return 0;
    }
}
