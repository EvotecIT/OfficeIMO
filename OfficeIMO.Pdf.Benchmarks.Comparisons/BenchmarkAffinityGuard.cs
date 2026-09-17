using System.Diagnostics;
using System.Globalization;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class BenchmarkAffinityGuard {
    internal const string EnvironmentVariable = "OFFICEIMO_EXPECTED_BENCHMARK_AFFINITY";

    internal static void Validate() {
        string? configured = Environment.GetEnvironmentVariable(EnvironmentVariable);
        if (string.IsNullOrWhiteSpace(configured)) return;

        string value = configured.StartsWith("0x", StringComparison.OrdinalIgnoreCase)
            ? configured[2..]
            : configured;
        if (!ulong.TryParse(value, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out ulong expected)) {
            throw new InvalidOperationException($"Invalid {EnvironmentVariable} value '{configured}'.");
        }
        long signed;
        if (OperatingSystem.IsWindows()) {
            signed = Process.GetCurrentProcess().ProcessorAffinity.ToInt64();
        } else if (OperatingSystem.IsLinux()) {
            signed = Process.GetCurrentProcess().ProcessorAffinity.ToInt64();
        } else {
            throw new PlatformNotSupportedException("Benchmark affinity validation requires Windows or Linux.");
        }
        ulong observed = unchecked((ulong)signed);
        if (observed != expected) {
            throw new InvalidOperationException(
                $"Benchmark worker affinity is 0x{observed:X}, expected 0x{expected:X}.");
        }
    }
}
