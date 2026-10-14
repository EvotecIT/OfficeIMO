using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    private static void ApplyGradientFill(OdgShape source, OfficeShape target, OdfConversionReport report, OfficeTransform transform) {
        if (!source.HasGradientFill) return;
        string feature = "shape:" + source.Name + ":fill-gradient";
        if (!IsClosedDrawGeometry(target)) {
            report.Add("shape:" + source.Name + ":inactive-fill", OdfConversionMappingStatus.Converted,
                message: "Draw open paths do not paint their declared gradient. The definition and binding remain preserved.");
            return;
        }
        try {
            OdfGradient definition = source.ResolveFillGradient();
            OdfGradientPattern pattern = definition.Pattern;
            ApplyNativeGradientFill(target, pattern, source.GradientStepCount);
            report.Add(feature, OdfConversionMappingStatus.Approximated,
                message: NativeGradientMessage(pattern));
            if ((target.FillGradient != null || target.FillRadialGradient != null) &&
                (transform.M12 != 0 || transform.M21 != 0 || transform.M11 != transform.M22 || transform.M11 <= 0))
                report.Add(feature + "-transform", OdfConversionMappingStatus.Unsupported,
                    message: "The declared local gradient transforms with the shape. Native Draw refitting under rotation, reflection, shear or nonuniform scale is not qualified for identical appearance.");
        } catch (Exception exception) when (exception is ArgumentException || exception is FormatException || exception is InvalidDataException || exception is NotSupportedException || exception is OverflowException) {
            report.Add(feature, OdfConversionMappingStatus.Unsupported, message: exception.Message);
        }
    }

    /// <summary>Projects the shared native two-color profile onto shape or background-local geometry.</summary>
    private static void ApplyNativeGradientFill(OfficeShape target, OdfGradientPattern pattern, int? stepCount) {
        if (stepCount > 0) throw new NotSupportedException("Fixed gradient bands are preserved but not projected.");
        if (pattern.Style is not (OdfGradientStyle.Linear or OdfGradientStyle.Axial or OdfGradientStyle.Radial))
            throw new NotSupportedException("Ellipsoid, square and rectangular gradient fields are preserved but not projected.");
        var bounds = target.Kind == OfficeShapeKind.Path ? OfficePathGeometry.Bounds(target.PathCommands) : (0D, 0D, target.Width, target.Height);
        double width = bounds.Item3 - bounds.Item1, height = bounds.Item4 - bounds.Item2;
        if (!(width > 0) || !(height > 0) || double.IsInfinity(width) || double.IsInfinity(height))
            throw new NotSupportedException("Gradient geometry requires finite positive paint bounds.");
        OfficeColor start = GradientColor(pattern.StartColor, pattern.StartIntensity), end = GradientColor(pattern.EndColor, pattern.EndIntensity);
        if (pattern.Style == OdfGradientStyle.Axial && pattern.Border == 1)
            throw new NotSupportedException("A 100% axial border leaves a singular center seam outside this projection profile.");
        if (pattern.Border == 1) target.FillColor = start;
        else if (pattern.Style == OdfGradientStyle.Radial) {
            if (pattern.CenterX < 0 || pattern.CenterX > 1 || pattern.CenterY < 0 || pattern.CenterY > 1)
                throw new NotSupportedException("Off-area native radial centers remain preserved but are not qualified for projection.");
            double centerX = (bounds.Item1 + width * pattern.CenterX) / target.Width;
            double centerY = (bounds.Item2 + height * pattern.CenterY) / target.Height;
            double scale = Math.Max(width, height), radius = scale * Math.Sqrt(Math.Pow(width / scale, 2) + Math.Pow(height / scale, 2)) / 2;
            // Native radial interpolation runs from the end color at the center to the start color at the border.
            target.FillRadialGradient = new OfficeRadialGradient(centerX, centerY, 0, 0, centerX, centerY,
                radius / target.Width, radius / target.Height,
                new[] { new OfficeGradientStop(0, end), new OfficeGradientStop(1 - pattern.Border, start), new OfficeGradientStop(1, start) });
        } else {
            double angle = (pattern.AngleDegrees % 360) * Math.PI / 180;
            double dx = Math.Sin(angle), dy = Math.Cos(angle), span = width * Math.Abs(dx) + height * Math.Abs(dy);
            double centerX = (bounds.Item1 + bounds.Item3) / 2, centerY = (bounds.Item2 + bounds.Item4) / 2;
            IReadOnlyList<OfficeGradientStop> stops = pattern.Style == OdfGradientStyle.Axial ? new[] {
                new OfficeGradientStop(0, end), new OfficeGradientStop(pattern.Border / 2, end), new OfficeGradientStop(0.5, start),
                new OfficeGradientStop(1 - pattern.Border / 2, end), new OfficeGradientStop(1, end)
            } : new[] { new OfficeGradientStop(0, start), new OfficeGradientStop(pattern.Border, start), new OfficeGradientStop(1, end) };
            // Preserve the physical color-field normal under the nonuniform normalization of the shape canvas.
            target.FillGradient = OfficeLinearGradient.CreateImported(centerX - dx * span / 2, centerY - dy * span / 2,
                centerX + dx * span / 2, centerY + dy * span / 2, stops)
                .TransformCoordinates(OfficeTransform.Scale(1 / target.Width, 1 / target.Height));
        }
    }

    private static string NativeGradientMessage(OdfGradientPattern pattern) => "Native two-color " + pattern.Style.ToString().ToLowerInvariant() +
        " fill uses geometry bounds, angle, border and color intensities. Automatic native banding is represented by continuous interpolation.";
    private static bool IsClosedDrawGeometry(OfficeShape shape) => shape.Kind != OfficeShapeKind.Line &&
        (shape.Kind != OfficeShapeKind.Path || OfficePathContour.Split(shape.PathCommands).All(contour => contour.IsClosed));
    private static OfficeColor GradientColor(OdfColor color, double intensity) {
        OfficeColor rgb = OfficeColor.Parse(color.ToString());
        return OfficeColor.FromRgb((byte)Math.Round(rgb.R * intensity), (byte)Math.Round(rgb.G * intensity), (byte)Math.Round(rgb.B * intensity));
    }
}
