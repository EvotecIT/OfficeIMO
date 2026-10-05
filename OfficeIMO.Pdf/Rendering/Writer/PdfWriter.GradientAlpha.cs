using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static string GradientAlphaStateName(string shadingName) => "GA_" + shadingName;

    private static bool HasGradientAlpha(OfficeShape shape) =>
        HasGradientAlpha(shape.FillRadialGradient?.Stops ?? shape.FillGradient?.Stops);

    private static bool HasGradientAlpha(IReadOnlyList<OfficeGradientStop>? stops) {
        if (stops == null) return false;
        foreach (OfficeGradientStop stop in stops) if (stop.Color.A != 255) return true;
        return false;
    }

    /// <summary>Accumulates the painted region in the shading's own coordinate system.</summary>
    private static void RegisterGradientAlphaBounds(IList<PageShading> shadings, string name,
        OfficeShape shape, double x, double bottomY, bool localCoordinates) {
        if (!HasGradientAlpha(shape) && shape.FillRadialGradient?.OutsideColor == null) return;
        GetGradientPaintBounds(shape, x, bottomY, localCoordinates, out double left, out double top, out double right, out double bottom);
        foreach (PageShading shading in shadings) {
            if (shading.Name != name) continue;
            shading.AlphaLeft = Math.Min(shading.AlphaLeft, left);
            shading.AlphaBottom = Math.Min(shading.AlphaBottom, top);
            shading.AlphaRight = Math.Max(shading.AlphaRight, right);
            shading.AlphaTop = Math.Max(shading.AlphaTop, bottom);
            return;
        }
    }

    private static void GetGradientPaintBounds(OfficeShape shape, double x, double bottomY, bool localCoordinates,
        out double left, out double top, out double right, out double bottom) {
        left = 0D; top = 0D; right = shape.Width; bottom = shape.Height;
        foreach (var contour in OfficeStrokeGeometry.FlattenShape(shape, 1D)) {
            foreach (OfficePoint point in contour.Points) {
                left = Math.Min(left, point.X); right = Math.Max(right, point.X);
                top = Math.Min(top, point.Y); bottom = Math.Max(bottom, point.Y);
            }
        }
        // Include curve approximation and antialiasing margins; these bounds do not
        // allocate a bitmap or change the clipping path used to paint the gradient.
        left -= 1D; top -= 1D; right += 1D; bottom += 1D;
        if (shape.FillRadialGradient != null) {
            var inverse = RadialShadingTransform(shape, 0D, 0D, localCoordinates: true).Invert();
            var corners = new[] { new OfficePoint(left, top), new OfficePoint(right, top), new OfficePoint(right, bottom), new OfficePoint(left, bottom) };
            left = top = double.PositiveInfinity; right = bottom = double.NegativeInfinity;
            foreach (var corner in corners) {
                var point = inverse.TransformPoint(corner);
                left = Math.Min(left, point.X); right = Math.Max(right, point.X);
                top = Math.Min(top, point.Y); bottom = Math.Max(bottom, point.Y);
            }
        } else if (!localCoordinates) {
            left += x; right += x;
            double lower = bottomY + shape.Height - bottom;
            bottom = bottomY + shape.Height - top;
            top = lower;
        }
    }

    /// <summary>Preserves stop alpha with a native luminosity mask, retaining vector shading.</summary>
    private static void AddGradientAlphaResources(IList<byte[]> objects, PageShading shading,
        List<(string Name, int Id)> graphicsStates) {
        if (!HasGradientAlpha(shading.Stops)) return;
        string maskShading = shading.IsRadial
            ? PdfVisualResourceDictionaryBuilder.BuildRadialShadingObject(shading.X0, shading.Y0, shading.R0,
                shading.X1, shading.Y1, shading.R1, shading.Stops, alphaOnly: true)
            : PdfVisualResourceDictionaryBuilder.BuildAxialShadingObject(shading.X0, shading.Y0, shading.X1, shading.Y1, shading.Stops, alphaOnly: true);
        int maskShadingId = AddObject(objects, maskShading);
        string entries = "/Type /XObject /Subtype /Form /FormType 1 /BBox [" +
            F(shading.AlphaLeft) + " " + F(shading.AlphaBottom) + " " +
            F(shading.AlphaRight) + " " + F(shading.AlphaTop) + "]" +
            " /Group << /S /Transparency /CS /DeviceGray /I true >>" +
            " /Resources << /Shading << /A " + maskShadingId + " 0 R >> >>";
        var maskContent = new StringBuilder();
        var content = new ContentStreamBuilder(maskContent);
        if (shading.OutsideColor is OfficeColor outside) {
            content.FillGray(outside.A / 255D)
                .Rectangle(shading.AlphaLeft, shading.AlphaBottom, shading.AlphaRight - shading.AlphaLeft, shading.AlphaTop - shading.AlphaBottom)
                .FillPath();
        }
        content.Shading("A");
        int maskFormId = AddFlateStreamObject(objects, Encoding.ASCII.GetBytes(maskContent.ToString()), entries);
        int stateId = AddObject(objects, "<< /Type /ExtGState /SMask << /S /Luminosity /G " + maskFormId + " 0 R /BC [0] >> >>\n");
        graphicsStates.Add(("/" + GradientAlphaStateName(shading.Name), stateId));
    }
}
