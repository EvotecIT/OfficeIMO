using System.Globalization;

namespace OfficeIMO.Studio.Infrastructure.Localization;

/// <summary>Applies the selected UI culture consistently to resources and presentation formatting.</summary>
internal sealed class StudioCultureService {
    internal CultureInfo Apply(string? cultureName) {
        string normalized = StudioCultureCatalog.NormalizeOrDefault(cultureName);
        CultureInfo culture = CultureInfo.GetCultureInfo(normalized);
        CultureInfo.CurrentCulture = culture;
        CultureInfo.CurrentUICulture = culture;
        CultureInfo.DefaultThreadCurrentCulture = culture;
        CultureInfo.DefaultThreadCurrentUICulture = culture;
        return culture;
    }
}
