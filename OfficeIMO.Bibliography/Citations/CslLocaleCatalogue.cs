using System.Collections.Concurrent;
using System.Text.Json;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal static class CslLocaleCatalogue {
    private const string Prefix = "OfficeIMO.Bibliography.Citations.Locales.";
    private static readonly ConcurrentDictionary<string, Lazy<XElement>> Cache = new ConcurrentDictionary<string, Lazy<XElement>>(StringComparer.OrdinalIgnoreCase);
    private static readonly Lazy<IReadOnlyDictionary<string, string>> PrimaryDialects = new Lazy<IReadOnlyDictionary<string, string>>(ReadDialects);

    internal static string PrimaryDialect(string language) => PrimaryDialects.Value.TryGetValue(language.Split('-')[0], out string? dialect) ? dialect : "en-US";

    internal static XElement? Get(string language) {
        string resource = Prefix + "locales-" + language + ".xml";
        if (!typeof(CslLocaleCatalogue).Assembly.GetManifestResourceNames().Contains(resource, StringComparer.Ordinal)) return null;
        return Cache.GetOrAdd(language, _ => new Lazy<XElement>(() => ReadLocale(resource))).Value;
    }

    private static XElement ReadLocale(string resource) {
        using Stream stream = typeof(CslLocaleCatalogue).Assembly.GetManifestResourceStream(resource)!;
        using var reader = new StreamReader(stream);
        return CslStyle.ReadXml(reader.ReadToEnd(), 4 * 1024 * 1024, 64, default);
    }

    private static IReadOnlyDictionary<string, string> ReadDialects() {
        using Stream stream = typeof(CslLocaleCatalogue).Assembly.GetManifestResourceStream(Prefix + "locales.json")!;
        using JsonDocument json = JsonDocument.Parse(stream);
        return json.RootElement.GetProperty("primary-dialects").EnumerateObject().ToDictionary(property => property.Name, property => property.Value.GetString()!, StringComparer.OrdinalIgnoreCase);
    }
}
