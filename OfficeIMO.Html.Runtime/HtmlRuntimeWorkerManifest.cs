using System.Reflection;
using System.Text.Json;

namespace OfficeIMO.Html.Runtime;

internal static class HtmlRuntimeWorkerManifest {
    private const int CurrentProtocolVersion = 1;
    private const int MaximumManifestCharacters = 64 * 1024;
    private const string ManifestFileName = "OfficeIMO.Html.Runtime.Worker.manifest.json";
    private const string WorkerAssemblyName = "OfficeIMO.Html.Runtime.Worker";

    internal static HtmlRuntimeProviderDescriptor Load(string workerPath) {
        AssemblyName worker = AssemblyName.GetAssemblyName(workerPath);
        Version runtimeVersion = typeof(HtmlRuntimeWorkerManifest).Assembly.GetName().Version
            ?? throw new NotSupportedException("The OfficeIMO runtime assembly has no compatibility version.");
        if (!string.Equals(worker.Name, WorkerAssemblyName, StringComparison.Ordinal)
            || worker.Version != runtimeVersion)
            throw new NotSupportedException($"The deployed runtime worker must be {WorkerAssemblyName} version {runtimeVersion}.");

        string manifestPath = Path.Combine(Path.GetDirectoryName(workerPath)!, ManifestFileName);
        if (!File.Exists(manifestPath))
            throw new FileNotFoundException("The runtime worker compatibility manifest is not deployed.", manifestPath);
        string json = File.ReadAllText(manifestPath);
        if (json.Length > MaximumManifestCharacters)
            throw new NotSupportedException("The runtime worker compatibility manifest is too large.");
        using JsonDocument document = JsonDocument.Parse(json, new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Disallow, MaxDepth = 8 });
        JsonElement root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object
            || RequiredInt(root, "protocolVersion") != CurrentProtocolVersion
            || !string.Equals(RequiredString(root, "providerId"), "officeimo.trusted-process", StringComparison.Ordinal))
            throw new NotSupportedException("The deployed runtime worker uses an incompatible protocol manifest.");

        HtmlRuntimeProfile[] profiles = RequiredArray(root, "profiles").Select(value => {
            string name = RequiredString(value);
            return Enum.TryParse(name, ignoreCase: false, out HtmlRuntimeProfile profile) && Enum.IsDefined(profile)
                ? profile
                : throw new NotSupportedException($"The runtime worker manifest advertises unknown profile '{name}'.");
        }).ToArray();
        string[] capabilities = RequiredArray(root, "capabilities").Select(RequiredString).ToArray();
        int maximumContexts = RequiredInt(root, "maximumContexts");
        int maximumPagesPerContext = RequiredInt(root, "maximumPagesPerContext");
        if (!SameSet(profiles, HtmlRuntimeProviderDescriptor.ProcessWorkerProfiles)
            || !SameSet(capabilities, HtmlRuntimeProviderDescriptor.ProcessWorkerCapabilities, StringComparer.Ordinal)
            || maximumContexts != int.MaxValue || maximumPagesPerContext != 1)
            throw new NotSupportedException("The runtime worker manifest does not match this protocol's supported profiles, capabilities, or limits.");
        return HtmlRuntimeProviderDescriptor.CreateProcessWorker(
            worker.Version?.ToString() ?? "unknown",
            HtmlRuntimeProviderDescriptor.ProcessWorkerProfiles,
            HtmlRuntimeProviderDescriptor.ProcessWorkerCapabilities,
            maximumContexts,
            maximumPagesPerContext);
    }

    private static int RequiredInt(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out JsonElement value) && value.TryGetInt32(out int number)
            ? number
            : throw new NotSupportedException($"The runtime worker manifest requires integer property '{name}'.");

    private static string RequiredString(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out JsonElement value) ? RequiredString(value)
            : throw new NotSupportedException($"The runtime worker manifest requires string property '{name}'.");

    private static string RequiredString(JsonElement value) => value.ValueKind == JsonValueKind.String
        ? value.GetString()!
        : throw new NotSupportedException("The runtime worker manifest contains a non-string array value.");

    private static IEnumerable<JsonElement> RequiredArray(JsonElement owner, string name) =>
        owner.TryGetProperty(name, out JsonElement value) && value.ValueKind == JsonValueKind.Array
            ? value.EnumerateArray()
            : throw new NotSupportedException($"The runtime worker manifest requires array property '{name}'.");

    private static bool SameSet<T>(IEnumerable<T> actual, IEnumerable<T> expected, IEqualityComparer<T>? comparer = null) {
        comparer ??= EqualityComparer<T>.Default;
        return new HashSet<T>(actual, comparer).SetEquals(expected)
            && actual.Count() == expected.Count();
    }
}
