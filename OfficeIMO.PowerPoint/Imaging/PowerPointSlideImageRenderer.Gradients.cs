using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static bool HasResolvableShapeFillGradient(PowerPointShape source,
            A.ColorScheme? colorScheme) =>
            TryCreateShapeFillGradient(source, colorScheme,
                out _, out _, out _);

        private static bool TryApplyShapeFillGradient(OfficeShape target,
            PowerPointShape source, A.ColorScheme? colorScheme) {
            if (!TryCreateShapeFillGradient(source, colorScheme,
                    out OfficeLinearGradient? linear,
                    out OfficeRadialGradient? radial,
                    out OfficeColor fallback)) {
                return false;
            }

            target.FillColor = fallback;
            target.FillGradient = linear;
            target.FillRadialGradient = radial;
            return true;
        }

        private static bool TryCreateShapeFillGradient(PowerPointShape source,
            A.ColorScheme? colorScheme,
            out OfficeLinearGradient? linear,
            out OfficeRadialGradient? radial,
            out OfficeColor fallback) {
            linear = null;
            radial = null;
            fallback = default;
            A.GradientFill? gradient = GetOpenXmlShapeProperties(source)?
                .GetFirstChild<A.GradientFill>();
            if (gradient == null
                || !TryResolveGradientStops(gradient, colorScheme,
                    out IReadOnlyList<OfficeGradientStop>? stops)) {
                return false;
            }

            fallback = stops![0].Color;
            A.LinearGradientFill? linearFill = gradient.GetFirstChild<A.LinearGradientFill>();
            if (linearFill != null) {
                double angle = (linearFill.Angle?.Value ?? 0) / 60000D;
                linear = OfficeLinearGradient.FromAngle(stops, angle);
                return true;
            }

            if (gradient.GetFirstChild<A.PathGradientFill>() != null) {
                radial = new OfficeRadialGradient(
                    0.5D, 0.5D, 0D,
                    0.5D, 0.5D, 0.5D,
                    stops);
                return true;
            }

            return false;
        }

        private static bool TryResolveGradientStops(A.GradientFill gradient,
            A.ColorScheme? colorScheme,
            out IReadOnlyList<OfficeGradientStop>? result) {
            result = null;
            A.GradientStop[] sourceStops = gradient.GetFirstChild<A.GradientStopList>()?
                .Elements<A.GradientStop>()
                .OrderBy(stop => stop.Position?.Value ?? 0)
                .ToArray() ?? Array.Empty<A.GradientStop>();
            if (sourceStops.Length < 2) return false;

            var stops = new List<OfficeGradientStop>(sourceStops.Length + 2);
            foreach (A.GradientStop sourceStop in sourceStops) {
                OfficeColor? color = OfficeOpenXmlThemeColorResolver.ResolveColor(
                    sourceStop, colorScheme);
                if (!color.HasValue) return false;
                double offset = Math.Max(0D, Math.Min(1D,
                    (sourceStop.Position?.Value ?? 0) / 100000D));
                stops.Add(new OfficeGradientStop(offset, color.Value));
            }

            if (stops[0].Offset > 0D) {
                stops.Insert(0, new OfficeGradientStop(0D, stops[0].Color));
            }
            if (stops[stops.Count - 1].Offset < 1D) {
                stops.Add(new OfficeGradientStop(1D, stops[stops.Count - 1].Color));
            }
            result = stops;
            return true;
        }
    }
}
