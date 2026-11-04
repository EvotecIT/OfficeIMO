using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>Encodes explicit DrawingML colors without resolving native theme slots to RGB.</summary>
internal static class PowerPointDrawingColor {
    internal static OpenXmlElement Create(OfficeColor color, OfficeTheme theme) {
        if (color.ThemeSlot is string slot) {
            var scheme = new A.SchemeColor();
            scheme.SetAttribute(new OpenXmlAttribute("val", string.Empty, slot));
            return scheme;
        }
        color = color.Resolve(theme);
        var rgb = new A.RgbColorModelHex { Val = color.ToRgbHex() };
        if (color.A < 255) rgb.Append(new A.Alpha { Val = (int)System.Math.Round(color.A * 100000D / 255D) });
        return rgb;
    }

    internal static A.SolidFill Fill(OfficeColor? color, OfficeTheme theme) =>
        new A.SolidFill(Create(color ?? throw new System.ArgumentNullException(nameof(color)), theme));

    /// <summary>Finds the current consuming theme without persisting wrapper state in the native XML.</summary>
    internal static OfficeTheme ThemeFor(OpenXmlPartContainer? owner) {
        if (owner is not OpenXmlPart part) return OfficeTheme.Default;
        PowerPointDocument? document = part.OpenXmlPackage.Features.Get<PowerPointDocument>();
        if (document == null) return OfficeTheme.Default;
        SlidePart? slide = part as SlidePart ?? (part as NotesSlidePart)?.SlidePart;
        if (slide != null) {
            PowerPointSlide? consuming = document.Slides.FirstOrDefault(candidate => ReferenceEquals(candidate.SlidePart, slide));
            if (consuming != null) return consuming.AuthoringTheme;
        }
        SlideLayoutPart? layout = part as SlideLayoutPart ?? slide?.SlideLayoutPart;
        SlideMasterPart? master = part as SlideMasterPart ?? layout?.SlideMasterPart;
        int masterIndex = master == null ? 0 : document.OpenXmlDocument.PresentationPart!.SlideMasterParts
            .TakeWhile(candidate => !ReferenceEquals(candidate, master)).Count();
        return document.ReadLayoutTheme(masterIndex,
            layout?.ThemeOverridePart?.ThemeOverride?.ColorScheme,
            layout?.ThemeOverridePart?.ThemeOverride?.FontScheme).Theme;
    }

    // These getters expose local formatting. Inherited fills and transformed scheme colors
    // have no exact OfficeColor representation and remain absent rather than being flattened.
    internal static OfficeColor? Read(OpenXmlElement? container) {
        A.SchemeColor? scheme = container?.GetFirstChild<A.SchemeColor>();
        if (scheme != null) {
            if (scheme.HasChildren) return null;
            return scheme.Val?.InnerText switch {
                "accent1" => OfficeColor.Accent1, "accent2" => OfficeColor.Accent2,
                "accent3" => OfficeColor.Accent3, "accent4" => OfficeColor.Accent4,
                "accent5" => OfficeColor.Accent5, "accent6" => OfficeColor.Accent6,
                "dk1" => OfficeColor.Text, "dk2" => OfficeColor.TextMuted,
                "lt1" => OfficeColor.Background, _ => null
            };
        }
        return OfficeOpenXmlThemeColorResolver.ResolveColor(container, null);
    }
}
