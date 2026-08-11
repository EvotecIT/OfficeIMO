using System.Reflection;
using System.Text.Json;

namespace OfficeIMO.Word.Benchmarks;

internal static class WordLibraryFeatureCatalog {
    private const string ResourceSuffix = "word-library-capabilities.json";
    private static readonly string[] AllowedSupport = ["yes", "low-level", "partial", "adjacent-package", "no-built-in"];

    internal static void WriteMarkdown(TextWriter writer) {
        WordLibraryCatalog catalog = Load();
        writer.WriteLine("| Capability | " + string.Join(" | ", catalog.Libraries.Select(library => library.Name)) + " |");
        writer.WriteLine("| --- | " + string.Join(" | ", catalog.Libraries.Select(_ => "---")) + " |");
        foreach (WordLibraryCapability capability in catalog.Capabilities) {
            writer.WriteLine(
                "| " + Escape(capability.Name) + " | " +
                string.Join(" | ", catalog.Libraries.Select(library =>
                    Escape(capability.Support[library.Id]))) + " |");
        }
        writer.WriteLine();
        writer.WriteLine("Support values: yes, low-level, partial, adjacent-package, no-built-in.");
    }

    private static WordLibraryCatalog Load() {
        Assembly assembly = typeof(WordLibraryFeatureCatalog).Assembly;
        string resourceName = assembly.GetManifestResourceNames()
            .Single(name => name.EndsWith(ResourceSuffix, StringComparison.Ordinal));
        using Stream stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidDataException("The embedded Word capability catalog could not be opened.");
        WordLibraryCatalog catalog = JsonSerializer.Deserialize<WordLibraryCatalog>(
            stream,
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidDataException("The Word capability catalog is empty.");
        Validate(catalog);
        return catalog;
    }

    private static void Validate(WordLibraryCatalog catalog) {
        if (catalog.Libraries.Count < 2 || catalog.Capabilities.Count == 0) {
            throw new InvalidDataException("The Word capability catalog must contain libraries and capabilities.");
        }
        string[] ids = catalog.Libraries.Select(library => library.Id).ToArray();
        if (ids.Any(string.IsNullOrWhiteSpace) || ids.Distinct(StringComparer.Ordinal).Count() != ids.Length) {
            throw new InvalidDataException("Word capability library identifiers must be non-empty and unique.");
        }
        foreach (WordLibraryCapability capability in catalog.Capabilities) {
            foreach (string id in ids) {
                if (!capability.Support.TryGetValue(id, out string? value)) {
                    throw new InvalidDataException("Capability '" + capability.Name + "' has no value for '" + id + "'.");
                }
                if (!AllowedSupport.Contains(value, StringComparer.Ordinal)) {
                    throw new InvalidDataException("Capability '" + capability.Name + "' has unknown support value '" + value + "'.");
                }
            }
        }
    }

    private static string Escape(string value) => value.Replace("|", "\\|", StringComparison.Ordinal);

    private sealed class WordLibraryCatalog {
        public List<WordLibraryDescriptor> Libraries { get; set; } = [];
        public List<WordLibraryCapability> Capabilities { get; set; } = [];
    }

    private sealed class WordLibraryDescriptor {
        public string Id { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }

    private sealed class WordLibraryCapability {
        public string Name { get; set; } = string.Empty;
        public Dictionary<string, string> Support { get; set; } = new(StringComparer.Ordinal);
    }
}
