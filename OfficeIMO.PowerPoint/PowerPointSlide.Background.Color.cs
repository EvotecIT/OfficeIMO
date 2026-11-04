using System;
using OfficeIMO.Drawing;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointSlide {
    /// <summary>Sets a solid background, preserving native theme references and literal transparency.</summary>
    public void SetBackgroundColor(OfficeColor color) {
        A.SchemeColorValues? scheme = color.Role switch {
            OfficeColorRole.Accent1 => A.SchemeColorValues.Accent1,
            OfficeColorRole.Accent2 => A.SchemeColorValues.Accent2,
            OfficeColorRole.Accent3 => A.SchemeColorValues.Accent3,
            OfficeColorRole.Accent4 => A.SchemeColorValues.Accent4,
            OfficeColorRole.Accent5 => A.SchemeColorValues.Accent5,
            OfficeColorRole.Accent6 => A.SchemeColorValues.Accent6,
            OfficeColorRole.Text => A.SchemeColorValues.Dark1,
            OfficeColorRole.TextMuted => A.SchemeColorValues.Dark2,
            OfficeColorRole.Background => A.SchemeColorValues.Light1,
            _ => null
        };
        // Remove the former background image through the native relationship cleanup.
        ClearBackgroundImage();
        if (scheme.HasValue) {
            BackgroundColor = "000000";
            var fill = SlideRoot.CommonSlideData!.Background!.BackgroundProperties!.GetFirstChild<A.SolidFill>()!;
            fill.RemoveAllChildren();
            fill.Append(new A.SchemeColor { Val = scheme.Value });
        } else {
            OfficeColor literal = color.IsThemeRole
                ? color.Resolve(OwnerDocument?.LayoutTheme.Theme ?? OfficeTheme.Default)
                : color;
            BackgroundColor = literal.ToRgbHex();
            if (literal.A < 255) {
                var rgb = SlideRoot.CommonSlideData!.Background!.BackgroundProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!;
                rgb.Append(new A.Alpha { Val = (int)Math.Round(literal.A * 100000D / 255D) });
            }
        }
    }

    /// <summary>Removes the slide's background override so the layout or master can supply its background.</summary>
    public void ClearBackground() {
        ClearBackgroundImage();
        CommonSlideData? common = SlideRoot.CommonSlideData;
        if (common != null) common.Background = null;
    }
}
