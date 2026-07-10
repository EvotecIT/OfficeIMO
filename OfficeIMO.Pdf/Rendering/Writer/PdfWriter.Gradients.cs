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
        bool elliptical = !gradient.EndRadiusX.Equals(gradient.EndRadiusY);
        double x0 = elliptical ? (gradient.StartX - gradient.EndX) / gradient.EndRadiusX : gradient.StartX;
        double y0 = elliptical ? (gradient.EndY - gradient.StartY) / gradient.EndRadiusY : 1D - gradient.StartY;
        double r0 = elliptical ? gradient.StartRadiusX / gradient.EndRadiusX : gradient.StartRadius;
        double x1 = elliptical ? 0D : gradient.EndX;
        double y1 = elliptical ? 0D : 1D - gradient.EndY;
        double r1 = elliptical ? 1D : gradient.EndRadius;
        for (int index = 0; index < shadings.Count; index++) {
            PageShading existing = shadings[index];
            if (existing.MatchesRadial(x0, y0, r0, x1, y1, r1, gradient.Stops)) return existing.Name;
        }

        string name = "SH" + (shadings.Count + 1).ToString(CultureInfo.InvariantCulture);
        shadings.Add(new PageShading {
            Name = name,
            IsRadial = true,
            Stops = new System.Collections.Generic.List<OfficeGradientStop>(gradient.Stops),
            X0 = x0,
            Y0 = y0,
            R0 = r0,
            X1 = x1,
            Y1 = y1,
            R1 = r1
        });
        return name;
    }

    private static void ApplyRadialGradientTransform(
        ContentStreamBuilder content,
        OfficeIMO.Drawing.OfficeShape shape,
        double x,
        double y) {
        OfficeRadialGradient gradient = shape.FillRadialGradient!;
        if (gradient.EndRadiusX.Equals(gradient.EndRadiusY)) {
            content.TransformMatrix(shape.Width, 0D, 0D, shape.Height, x, y);
            return;
        }

        content.TransformMatrix(
            shape.Width * gradient.EndRadiusX,
            0D,
            0D,
            shape.Height * gradient.EndRadiusY,
            x + (shape.Width * gradient.EndX),
            y + (shape.Height * (1D - gradient.EndY)));
    }
}
