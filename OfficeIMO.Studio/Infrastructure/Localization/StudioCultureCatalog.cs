using System.Globalization;

namespace OfficeIMO.Studio.Infrastructure.Localization;

internal sealed record StudioCultureDescriptor(string Name, string EnglishName, string NativeName, bool HasReviewedTranslation);

/// <summary>Defines the culture-pack contract without advertising unfinished translations as available.</summary>
internal static class StudioCultureCatalog {
    internal const string DefaultCulture = "en";
    internal const string PseudoCulture = "en-XA";

    // The languages the picker offers, each with complete StudioStrings.<name>.resx resources and the matching Apple
    // String Catalog translations. Build/Studio/Import-StudioTranslations.ps1 adds a language here with its strings.
    private static readonly string[] ShippedCultures = ["en", "pl", "de", "fr", "es", "it", "pt-BR"];

    // Culture names reserved for future packs. They are accepted in preferences but not offered until shipped.
    private static readonly string[] ReservedCultures = [
        "it", "es", "pt", "nl", "cs", "sk", "uk", "ro", "hu", "sv", "da", "fi", "nb", "ar", "he", "ja", "ko", "zh-Hans", "zh-Hant"
    ];

    internal static IReadOnlyList<StudioCultureDescriptor> Available { get; } = [
        .. ShippedCultures.Select(static name => Create(name, reviewed: true)),
        new StudioCultureDescriptor(PseudoCulture, "Pseudolocalized English", "Pseudolocalized English", true)
    ];

    internal static IReadOnlyList<StudioCultureDescriptor> Planned { get; } =
        [.. ReservedCultures.Except(ShippedCultures, StringComparer.OrdinalIgnoreCase).Select(static name => Create(name, reviewed: false))];

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
