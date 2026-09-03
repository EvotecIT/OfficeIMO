using System.Globalization;

namespace OfficeIMO.Studio.Infrastructure.Localization;

internal sealed record StudioCultureDescriptor(string Name, string EnglishName, string NativeName, bool HasReviewedTranslation);

/// <summary>Defines the culture-pack contract without advertising unfinished translations as available.</summary>
internal static class StudioCultureCatalog {
    internal const string DefaultCulture = "en";
    internal const string PseudoCulture = "en-XA";

    internal static IReadOnlyList<StudioCultureDescriptor> Available { get; } = [
        Create("en", reviewed: true),
        new StudioCultureDescriptor(PseudoCulture, "Pseudolocalized English", "Pseudolocalized English", true)
    ];

    internal static IReadOnlyList<StudioCultureDescriptor> Planned { get; } = [
        Create("en", reviewed: true), Create("pl", reviewed: false), Create("de", reviewed: false),
        Create("fr", reviewed: false), Create("it", reviewed: false), Create("es", reviewed: false),
        Create("pt", reviewed: false), Create("nl", reviewed: false), Create("cs", reviewed: false),
        Create("sk", reviewed: false), Create("uk", reviewed: false), Create("ro", reviewed: false),
        Create("hu", reviewed: false), Create("sv", reviewed: false), Create("da", reviewed: false),
        Create("fi", reviewed: false), Create("nb", reviewed: false), Create("ar", reviewed: false),
        Create("he", reviewed: false), Create("ja", reviewed: false), Create("ko", reviewed: false),
        Create("zh-Hans", reviewed: false), Create("zh-Hant", reviewed: false)
    ];

    internal static string NormalizeOrDefault(string? cultureName) {
        if (string.IsNullOrWhiteSpace(cultureName)) return DefaultCulture;
        try {
            string normalized = CultureInfo.GetCultureInfo(cultureName.Trim()).Name;
            return Available.Concat(Planned).Any(culture =>
                string.Equals(culture.Name, normalized, StringComparison.OrdinalIgnoreCase))
                ? normalized
                : DefaultCulture;
        } catch (CultureNotFoundException) {
            return DefaultCulture;
        }
    }

    private static StudioCultureDescriptor Create(string name, bool reviewed) {
        CultureInfo culture = CultureInfo.GetCultureInfo(name);
        return new StudioCultureDescriptor(culture.Name, culture.EnglishName, culture.NativeName, reviewed);
    }
}
