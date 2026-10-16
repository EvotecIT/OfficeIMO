using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        // Resolve in the consuming slide context, including inherited shapes/backgrounds.
        internal A.ColorScheme? ResolvedColorScheme =>
            SlidePart.ThemeOverridePart?.ThemeOverride?.ColorScheme
            ?? SlidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.ColorScheme
            ?? SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;

        internal A.FormatScheme? ResolvedFormatScheme =>
            SlidePart.ThemeOverridePart?.ThemeOverride?.FormatScheme
            ?? SlidePart.SlideLayoutPart?.ThemeOverridePart?.ThemeOverride?.FormatScheme
            ?? SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FormatScheme;

        internal OpenXmlElement? ResolvedColorMap {
            get {
                OpenXmlElement? master = SlidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.ColorMap;
                var slideOverride = SlidePart.Slide?.ColorMapOverride;
                if (slideOverride?.GetFirstChild<A.MasterColorMapping>() != null) return master;
                if (slideOverride?.GetFirstChild<A.OverrideColorMapping>() is { } slideMap) return slideMap;
                var layoutOverride = SlidePart.SlideLayoutPart?.SlideLayout?.ColorMapOverride;
                return layoutOverride?.GetFirstChild<A.OverrideColorMapping>() ?? master;
            }
        }
    }
}
