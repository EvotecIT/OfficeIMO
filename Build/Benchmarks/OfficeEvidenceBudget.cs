using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.Benchmarks;

internal sealed record OfficeEvidenceObservation(
    string Key,
    double ElapsedMicrosecondsPerOperation,
    long AllocatedBytesPerOperation,
    long PeakManagedHeapGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes,
    long? ArtifactBytes);

internal static class OfficeEvidenceBudgetEvaluator {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNameCaseInsensitive = true
    };

    internal static void EnsureWithin(
        string budgetPath,
        string suite,
        IReadOnlyCollection<OfficeEvidenceObservation> observations) {
        string fullPath = Path.GetFullPath(budgetPath);
        OfficeEvidenceBudgetFile file = JsonSerializer.Deserialize<OfficeEvidenceBudgetFile>(
            File.ReadAllText(fullPath), JsonOptions) ?? throw new InvalidDataException("The evidence budget file is empty.");
        if (file.SchemaVersion != 1) throw new InvalidDataException($"Unsupported evidence budget schema {file.SchemaVersion}.");
        string operatingSystem = GetOperatingSystem();
        OfficeEvidenceBudget[] suiteBudgets = file.Budgets
            .Where(item => string.Equals(item.Suite, suite, StringComparison.Ordinal))
            .ToArray();
        if (suiteBudgets.Length == 0) throw new InvalidDataException($"No budgets are defined for suite '{suite}'.");
        string[] requiredKeys = suiteBudgets.Select(item => item.Key).Distinct(StringComparer.Ordinal).ToArray();
        var failures = new List<string>();
        foreach (string key in requiredKeys) {
            OfficeEvidenceBudget? budget = suiteBudgets.SingleOrDefault(item =>
                string.Equals(item.OperatingSystem, operatingSystem, StringComparison.Ordinal) &&
                string.Equals(item.Key, key, StringComparison.Ordinal));
            budget ??= suiteBudgets.SingleOrDefault(item =>
                string.Equals(item.OperatingSystem, "Any", StringComparison.Ordinal) &&
                string.Equals(item.Key, key, StringComparison.Ordinal));
            if (budget == null) {
                failures.Add($"{key}: no {operatingSystem} budget is defined");
                continue;
            }
            OfficeEvidenceObservation[] matching = observations.Where(item => item.Key == key).ToArray();
            if (matching.Length == 0) {
                failures.Add($"{key}: no measurement was produced");
                continue;
            }
            foreach (OfficeEvidenceObservation observation in matching) Evaluate(observation, budget, failures);
        }
        if (failures.Count != 0) {
            throw new InvalidOperationException(
                $"{suite} evidence exceeded its {operatingSystem} budget:{Environment.NewLine}- " +
                string.Join(Environment.NewLine + "- ", failures));
        }
        Console.WriteLine($"{suite} evidence satisfied {requiredKeys.Length} {operatingSystem} budget(s).");
    }

    private static void Evaluate(
        OfficeEvidenceObservation observation,
        OfficeEvidenceBudget budget,
        ICollection<string> failures) {
        AddFailure(failures, observation.Key, "elapsed us/op", observation.ElapsedMicrosecondsPerOperation, budget.MaxElapsedMicrosecondsPerOperation);
        AddFailure(failures, observation.Key, "allocated bytes/op", observation.AllocatedBytesPerOperation, budget.MaxAllocatedBytesPerOperation);
        AddFailure(failures, observation.Key, "managed peak bytes", observation.PeakManagedHeapGrowthBytes, budget.MaxPeakManagedHeapGrowthBytes);
        AddFailure(failures, observation.Key, "process peak bytes", observation.AbsoluteProcessPeakWorkingSetBytes, budget.MaxAbsoluteProcessPeakWorkingSetBytes);
        if (budget.MaxArtifactBytes.HasValue) {
            if (!observation.ArtifactBytes.HasValue) failures.Add($"{observation.Key}: no artifact size was recorded");
            else AddFailure(failures, observation.Key, "artifact bytes", observation.ArtifactBytes.Value, budget.MaxArtifactBytes.Value);
        }
    }

    private static void AddFailure(
        ICollection<string> failures,
        string key,
        string metric,
        double actual,
        double maximum) {
        if (maximum <= 0) throw new InvalidDataException($"{key}: {metric} budget must be positive.");
        if (actual > maximum) failures.Add($"{key}: {metric} {actual:F2} > {maximum:F2}");
    }

    private static string GetOperatingSystem() {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return "Windows";
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux)) return "Linux";
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX)) return "macOS";
        throw new PlatformNotSupportedException("Evidence budgets are defined for Windows, Linux, and macOS.");
    }
}

internal sealed class OfficeEvidenceBudgetFile {
    public int SchemaVersion { get; set; }
    public List<OfficeEvidenceBudget> Budgets { get; set; } = new();
}

internal sealed class OfficeEvidenceBudget {
    public string Suite { get; set; } = string.Empty;
    public string OperatingSystem { get; set; } = string.Empty;
    public string Key { get; set; } = string.Empty;
    public double MaxElapsedMicrosecondsPerOperation { get; set; }
    public long MaxAllocatedBytesPerOperation { get; set; }
    public long MaxPeakManagedHeapGrowthBytes { get; set; }
    public long MaxAbsoluteProcessPeakWorkingSetBytes { get; set; }
    public long? MaxArtifactBytes { get; set; }
}
