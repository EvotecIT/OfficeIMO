using System.Globalization;
using System.Resources;

namespace OfficeIMO.Studio.Infrastructure.Localization;

/// <summary>Resolves neutral English resources, culture satellites, and the expansion pseudolocale.</summary>
internal sealed class StudioLocalizer : IStudioLocalizer {
    private static readonly ResourceManager Resources = new(
        "OfficeIMO.Studio.Localization.StudioStrings",
        typeof(StudioLocalizer).Assembly);

    internal StudioLocalizer(CultureInfo culture) {
        Culture = culture ?? throw new ArgumentNullException(nameof(culture));
    }

    public CultureInfo Culture { get; }

    public string Get(string key) {
        ArgumentException.ThrowIfNullOrWhiteSpace(key);
        string value = Resources.GetString(key, Culture) ?? Resources.GetString(key, CultureInfo.GetCultureInfo("en")) ?? $"⟦{key}⟧";
        return IsPseudoCulture ? StudioPseudoLocalizer.Transform(value) : value;
    }

    public string GetOrDefault(string key, string fallback) {
        ArgumentException.ThrowIfNullOrWhiteSpace(key);
        ArgumentNullException.ThrowIfNull(fallback);
        string value = Resources.GetString(key, Culture) ?? Resources.GetString(key, CultureInfo.GetCultureInfo("en")) ?? fallback;
        return IsPseudoCulture ? StudioPseudoLocalizer.Transform(value) : value;
    }

    public string Format(string key, params object?[] arguments) =>
        string.Format(Culture, Get(key), arguments);

    public string FormatOrDefault(string key, string fallback, params object?[] arguments) =>
        string.Format(Culture, GetOrDefault(key, fallback), arguments);

    private bool IsPseudoCulture => string.Equals(Culture.Name, StudioCultureCatalog.PseudoCulture, StringComparison.OrdinalIgnoreCase);
}
