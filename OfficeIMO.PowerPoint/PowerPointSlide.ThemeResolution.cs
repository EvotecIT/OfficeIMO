using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        // Resolve in the consuming slide context, including inherited shapes/backgrounds.
        internal A.ColorScheme? ResolvedColorScheme =>
            SlidePart.ThemeOverridePart?.ThemeOverride?.ColorScheme
            ?? SlidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.ColorScheme
            ?? SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;

        internal OfficeTheme AuthoringTheme => OwnerDocument?.ReadLayoutTheme(
            Math.Max(0, GetLayoutInfo()?.MasterIndex ?? 0), ResolvedColorScheme,
            SlidePart.ThemeOverridePart?.ThemeOverride?.FontScheme
            ?? SlidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.FontScheme
            ?? SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FontScheme).Theme
            ?? OfficeTheme.Default;

        internal A.FormatScheme? ResolvedFormatScheme =>
            SlidePart.ThemeOverridePart?.ThemeOverride?.FormatScheme
            ?? SlidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.FormatScheme
            ?? SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FormatScheme;

        internal OpenXmlElement? ResolvedColorMap => ResolveExportColorMap(SlidePart);

        internal static OpenXmlElement? ResolveExportColorMap(DocumentFormat.OpenXml.Packaging.SlidePart? slidePart) {
            OpenXmlElement? master = slidePart?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.ColorMap;
            var slideOverride = slidePart?.Slide?.ColorMapOverride;
            if (slideOverride?.GetFirstChild<A.MasterColorMapping>() != null) return master;
            if (slideOverride?.GetFirstChild<A.OverrideColorMapping>() is { } slideMap) return slideMap;
            var layoutOverride = slidePart?.SlideLayoutPart?.SlideLayout?.ColorMapOverride;
            return layoutOverride?.GetFirstChild<A.OverrideColorMapping>() ?? master;
        }
    }
}
