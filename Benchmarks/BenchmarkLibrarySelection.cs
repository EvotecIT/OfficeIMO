using System;

namespace OfficeIMO.Benchmarks;

/// <summary>
/// Restricted comparison libraries (see <c>Build/benchmark-libraries.json</c>) run only when the caller enabled them.
/// The benchmark runners list the enabled library ids in <see cref="EnabledLibrariesVariable"/>; a benchmark run
/// without it leaves every restricted library out. Nothing here declares a licence: an enabled library still reads
/// its own key or setting from the caller.
/// </summary>
internal static class BenchmarkLibrarySelection {
    internal const string EnabledLibrariesVariable = "OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES";

    /// <summary>BenchmarkDotNet category prefix that marks a restricted library's cases, as in <c>Library:epplus</c>.</summary>
    internal const string CategoryPrefix = "Library:";

    /// <summary>Whether the restricted library <paramref name="libraryId"/> is enabled for this process.</summary>
    internal static bool IsEnabled(string libraryId) {
        string? value = Environment.GetEnvironmentVariable(EnabledLibrariesVariable);
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        foreach (string enabled in value!.Split(',')) {
            if (string.Equals(enabled.Trim(), libraryId, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    internal static string SkippedMessage(string libraryId) =>
        $"{libraryId} skipped: it is a restricted library in Build/benchmark-libraries.json. Enable it with the runner's -IncludeLibrary or -AcceptLibraryLicense, or list it in {EnabledLibrariesVariable}.";
}
