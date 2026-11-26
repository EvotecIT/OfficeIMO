using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// The SDK forms of the color helpers, for code that still builds SDK elements; they go once that code works on
/// <see cref="OfficeXmlNode"/>.
/// </summary>
internal static partial class PowerPointDrawingColor {
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

    /// <summary>
    /// The theme content in <paramref name="part"/> consumes: its slide's authoring theme, or its layout's and master's.
    /// A part of the package is read through the SDK, which this opens.
    /// </summary>
    internal static OfficeTheme ThemeFor(OpenXml.PowerPointPartRef part, PowerPointSlide? slide) {
        if (slide != null) return slide.AuthoringTheme;
        if (part.IsNone) return OfficeTheme.Default;
        object? sdk = part.Current.SdkPart;
        if (sdk == null && part.Document is { } document) {
            _ = document.OpenXmlDocument;
            sdk = part.Current.SdkPart;
        }
        return ThemeFor(sdk as OpenXmlPartContainer);
    }

    internal static OfficeColor? Read(OpenXmlElement? container) => Read(OfficeXmlNode.FromSdk(container));
}
