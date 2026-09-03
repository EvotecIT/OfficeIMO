using System.Globalization;

namespace OfficeIMO.Studio.Infrastructure.Localization;

/// <summary>Provides the XAML localization bridge configured by the application composition root.</summary>
internal static class StudioLocalization {
    private static IStudioLocalizer _current = new StudioLocalizer(CultureInfo.GetCultureInfo(StudioCultureCatalog.DefaultCulture));

    internal static IStudioLocalizer Current => _current;

    internal static void Configure(IStudioLocalizer localizer) =>
        _current = localizer ?? throw new ArgumentNullException(nameof(localizer));
}
