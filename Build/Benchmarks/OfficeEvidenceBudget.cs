using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.Benchmarks;

internal sealed record OfficeEvidenceObservation(
    string Key,
    long WorkloadBytes,
    double ElapsedMicrosecondsPerOperation,
    long AllocatedBytesPerOperation,
    long PeakManagedHeapGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes,
    long? ArtifactBytes);

internal sealed record OfficeEvidenceRequirement(
    string Key,
    long ExpectedWorkloadBytes,
    string WorkloadLabel);

internal static class OfficeEvidenceBudgetEvaluator {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNameCaseInsensitive = true
    };

    internal static void EnsureWithin(
        string budgetPath,
        string suite,
        IReadOnlyCollection<OfficeEvidenceRequirement> requirements,
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
        if (requirements.Count == 0) throw new InvalidDataException($"No workload requirements are defined for suite '{suite}'.");
        var failures = new List<string>();
        OfficeEvidenceRequirement[] duplicateRequirements = requirements
            .GroupBy(item => item.Key, StringComparer.Ordinal)
            .Where(group => group.Count() != 1)
            .Select(group => group.First())
            .ToArray();
        foreach (OfficeEvidenceRequirement duplicate in duplicateRequirements) {
            failures.Add($"{duplicate.Key}: duplicate workload requirement");
        }
        foreach (OfficeEvidenceRequirement requirement in requirements) {
            if (string.IsNullOrWhiteSpace(requirement.Key)) failures.Add("A workload requirement has no key");
            if (requirement.ExpectedWorkloadBytes <= 0) failures.Add($"{requirement.Key}: required workload bytes must be positive");
            if (string.IsNullOrWhiteSpace(requirement.WorkloadLabel)) failures.Add($"{requirement.Key}: workload label is missing");
        }
        HashSet<string> requiredKeys = requirements.Select(item => item.Key).ToHashSet(StringComparer.Ordinal);
        foreach (string extraKey in suiteBudgets.Select(item => item.Key).Distinct(StringComparer.Ordinal).Where(key => !requiredKeys.Contains(key))) {
            failures.Add($"{extraKey}: budget has no matching workload requirement");
        }
        foreach (OfficeEvidenceRequirement requirement in requirements) {
            OfficeEvidenceBudget[] operatingSystemBudgets = suiteBudgets.Where(item =>
                string.Equals(item.OperatingSystem, operatingSystem, StringComparison.Ordinal) &&
                string.Equals(item.Key, requirement.Key, StringComparison.Ordinal)).ToArray();
            OfficeEvidenceBudget[] portableBudgets = suiteBudgets.Where(item =>
                string.Equals(item.OperatingSystem, "Any", StringComparison.Ordinal) &&
                string.Equals(item.Key, requirement.Key, StringComparison.Ordinal)).ToArray();
            if (operatingSystemBudgets.Length > 1 || portableBudgets.Length > 1) {
                failures.Add($"{requirement.Key}: duplicate {operatingSystem}/Any budget definition");
                continue;
            }
            OfficeEvidenceBudget? budget = operatingSystemBudgets.SingleOrDefault() ?? portableBudgets.SingleOrDefault();
            if (budget == null) {
                failures.Add($"{requirement.Key}: no {operatingSystem} or Any budget is defined");
                continue;
            }
            OfficeEvidenceObservation[] matching = observations.Where(item => item.Key == requirement.Key).ToArray();
            if (matching.Length == 0) {
                failures.Add($"{requirement.Key}: no measurement was produced");
                continue;
            }
            foreach (OfficeEvidenceObservation observation in matching) {
                if (observation.WorkloadBytes != requirement.ExpectedWorkloadBytes) {
                    failures.Add(
                        $"{requirement.Key}: {requirement.WorkloadLabel} {observation.WorkloadBytes} != " +
                        $"required {requirement.ExpectedWorkloadBytes}");
                }
                Evaluate(observation, budget, failures);
            }
        }
        if (failures.Count != 0) {
            throw new InvalidOperationException(
                $"{suite} evidence exceeded its {operatingSystem} budget:{Environment.NewLine}- " +
                string.Join(Environment.NewLine + "- ", failures));
        }
        Console.WriteLine($"{suite} evidence satisfied {requirements.Count} required {operatingSystem} budget(s).");
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
