using System;
using System.Collections.Generic;
using System.IO;
#if !NETFRAMEWORK
using System.Text.Json;
#endif

namespace OfficeIMO.Benchmarks;

/// <summary>
/// Restricted comparison libraries (see <c>Build/benchmark-libraries.json</c>) run only when the caller enabled them.
/// The benchmark runners list the enabled library ids in <see cref="EnabledLibrariesVariable"/>; a benchmark run
/// without it leaves every restricted library out. Nothing here declares a licence: an enabled library still reads
/// its own key or setting from the caller.
/// </summary>
internal static class BenchmarkLibrarySelection {
    internal const string EnabledLibrariesVariable = "OFFICEIMO_BENCHMARK_ENABLED_LIBRARIES";

    /// <summary>
    /// BenchmarkDotNet category prefix that names the library a case measures, as in <c>library:onepdf</c> or
    /// <c>Library:epplus</c>. The prefix is matched without regard to case.
    /// </summary>
    internal const string CategoryPrefix = "Library:";

    private static readonly Lazy<HashSet<string>> OpenLibraries = new(LoadOpenLibraries);

    /// <summary>
    /// Whether the library <paramref name="libraryId"/> may run in this process: open libraries always may, restricted
    /// ones only when the caller listed them.
    /// </summary>
    internal static bool IsEnabled(string libraryId) => IsOpen(libraryId) || IsListed(libraryId);

    /// <summary>
    /// Whether the policy marks <paramref name="libraryId"/> as open. A library the policy does not declare, or a run
    /// that cannot read the policy, counts as restricted.
    /// </summary>
    internal static bool IsOpen(string libraryId) => OpenLibraries.Value.Contains(libraryId.Trim());

    private static bool IsListed(string libraryId) {
        string? value = Environment.GetEnvironmentVariable(EnabledLibrariesVariable);
        if (string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        foreach (string enabled in value!.Split(',')) {
            if (string.Equals(enabled.Trim(), libraryId.Trim(), StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    internal static string SkippedMessage(string libraryId) =>
        $"{libraryId} skipped: it is a restricted library in Build/benchmark-libraries.json. Enable it with the runner's -IncludeLibrary or -AcceptLibraryLicense, or list it in {EnabledLibrariesVariable}.";

    private static HashSet<string> LoadOpenLibraries() {
        var open = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
#if NETFRAMEWORK
        // .NET Framework builds read no policy, so every library-tagged case there stays restricted until listed.
        return open;
#else
        string? path = FindPolicy();
        if (path == null) {
            ReportPolicyProblem("Build/benchmark-libraries.json was not found");
            return open;
        }

        try {
            using JsonDocument policy = JsonDocument.Parse(File.ReadAllBytes(path));
            foreach (JsonElement library in policy.RootElement.GetProperty("libraries").EnumerateArray()) {
                if (library.TryGetProperty("restriction", out JsonElement restriction) &&
                    restriction.ValueKind == JsonValueKind.String && restriction.GetString() == "open" &&
                    library.TryGetProperty("id", out JsonElement id) && id.ValueKind == JsonValueKind.String &&
                    id.GetString() is { Length: > 0 } value) {
                    open.Add(value);
                }
            }

            return open;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or JsonException or
                                                KeyNotFoundException or InvalidOperationException) {
            ReportPolicyProblem($"{path} could not be read ({exception.GetType().Name}: {exception.Message})");
            return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        }
#endif
    }

#if !NETFRAMEWORK
    /// <summary>Overrides where the policy is read from; otherwise it is found above the benchmark binaries.</summary>
    internal const string PolicyPathVariable = "OFFICEIMO_BENCHMARK_LIBRARY_POLICY";

    private static string? FindPolicy() {
        string? configured = Environment.GetEnvironmentVariable(PolicyPathVariable);
        if (!string.IsNullOrWhiteSpace(configured)) {
            return configured;
        }

        foreach (string start in new[] { AppContext.BaseDirectory, Environment.CurrentDirectory }) {
            for (DirectoryInfo? directory = new(start); directory != null; directory = directory.Parent) {
                string candidate = Path.Combine(directory.FullName, "Build", "benchmark-libraries.json");
                if (File.Exists(candidate)) {
                    return candidate;
                }
            }
        }

        return null;
    }

    // Without a readable policy no library counts as open: every library-tagged case runs only when the caller lists it.
    private static void ReportPolicyProblem(string problem) =>
        Console.Error.WriteLine(
            $"Benchmark library policy unavailable: {problem}. Every library-tagged case is treated as restricted and runs only when listed in {EnabledLibrariesVariable}.");
#endif
}
