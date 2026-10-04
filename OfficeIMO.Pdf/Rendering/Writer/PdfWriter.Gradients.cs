using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string? EnsureShapeFillGradient(
        System.Collections.Generic.IList<PageShading> shadings,
        OfficeShape shape, double xShape, double bottomY, bool localCoordinates) {
        if (shape.Kind == OfficeShapeKind.Line) return null;
        if (shape.FillRadialGradient != null) {
            string radialName = EnsureRadialShading(shadings, shape.FillRadialGradient);
            RegisterGradientAlphaBounds(shadings, radialName, shape, xShape, bottomY, localCoordinates);
            return radialName;
        }
        OfficeLinearGradient? gradient = shape.FillGradient;
        if (gradient == null) return null;

        double originX = localCoordinates ? 0D : xShape;
        double originY = localCoordinates ? 0D : bottomY;
        OfficeLinearGradient projected = gradient.TransformCoordinates(
            new OfficeTransform(shape.Width, 0D, 0D, localCoordinates ? shape.Height : -shape.Height,
                originX, localCoordinates ? 0D : originY + shape.Height));
        string name = EnsureAxialShading(shadings, gradient,
            projected.StartX, projected.StartY, projected.EndX, projected.EndY);
        RegisterGradientAlphaBounds(shadings, name, shape, xShape, bottomY, localCoordinates);
        return name;
    }

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
        double y,
        bool localCoordinates = false) {
        var transform = RadialShadingTransform(shape, x, y, localCoordinates);
        content.TransformMatrix(transform.M11, transform.M12, transform.M21, transform.M22, transform.OffsetX, transform.OffsetY);
    }

    private static OfficeTransform RadialShadingTransform(OfficeShape shape, double x, double y, bool localCoordinates) {
        OfficeRadialGradient gradient = shape.FillRadialGradient!;
        var coordinates = gradient.EndRadiusX.Equals(gradient.EndRadiusY)
            ? new OfficeTransform(1D, 0D, 0D, -1D, 0D, 1D)
            : new OfficeTransform(gradient.EndRadiusX, 0D, 0D, -gradient.EndRadiusY, gradient.EndX, gradient.EndY);
        return coordinates.Then(gradient.CoordinateTransform).Then(new OfficeTransform(
            shape.Width, 0D, 0D, localCoordinates ? shape.Height : -shape.Height,
            x, localCoordinates ? 0D : y + shape.Height));
    }
}
