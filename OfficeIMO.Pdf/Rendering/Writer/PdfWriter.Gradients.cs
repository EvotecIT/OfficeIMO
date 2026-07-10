using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string EnsureAxialShading(
        System.Collections.Generic.IList<PageShading> shadings,
        OfficeLinearGradient gradient,
        double x0,
        double y0,
        double x1,
        double y1) {
        for (int index = 0; index < shadings.Count; index++) {
            PageShading existing = shadings[index];
            if (existing.MatchesAxial(x0, y0, x1, y1, gradient.Stops)) return existing.Name;
        }

        string name = "SH" + (shadings.Count + 1).ToString(CultureInfo.InvariantCulture);
        shadings.Add(new PageShading {
            Name = name,
            Stops = new System.Collections.Generic.List<OfficeGradientStop>(gradient.Stops),
            X0 = x0,
            Y0 = y0,
            X1 = x1,
            Y1 = y1
        });
        return name;
    }

    private static string EnsureRadialShading(
        System.Collections.Generic.IList<PageShading> shadings,
        OfficeRadialGradient gradient) {
        double x0 = gradient.StartX;
        double y0 = 1D - gradient.StartY;
        double x1 = gradient.EndX;
        double y1 = 1D - gradient.EndY;
        for (int index = 0; index < shadings.Count; index++) {
            PageShading existing = shadings[index];
            if (existing.MatchesRadial(x0, y0, gradient.StartRadius, x1, y1, gradient.EndRadius, gradient.Stops)) return existing.Name;
        }

        string name = "SH" + (shadings.Count + 1).ToString(CultureInfo.InvariantCulture);
        shadings.Add(new PageShading {
            Name = name,
            IsRadial = true,
            Stops = new System.Collections.Generic.List<OfficeGradientStop>(gradient.Stops),
            X0 = x0,
            Y0 = y0,
            R0 = gradient.StartRadius,
            X1 = x1,
            Y1 = y1,
            R1 = gradient.EndRadius
        });
        return name;
    }
}
