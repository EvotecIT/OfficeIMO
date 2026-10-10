using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    private void ProjectFill(PublisherEscherShape source, OfficeShape shape, double width, double height) {
        OfficeArtShapeStyle style = source.Style;
        if (style.FillEnabled == false || style.FillType.GetValueOrDefault() == 0) return;
        string location = PublisherEscherReader.ShapeLocation(source.Id);
        if (style.FillType is not (4U or 7U)) {
            _context.Add("PUB_FILL_APPROXIMATED", $"Native fill type {style.FillType} was approximated by its primary color.",
                OfficeConversionLossKind.Approximation, location);
            return;
        }
        if (!OfficeArtLinearGradientProjector.TryProject(style, width, height, reference => Resolve(reference, source.Id),
                out OfficeArtLinearGradientFill? fill, out string? failure)) {
            string code = style.FillUsesCustomRectangle == true || style.FillAlignedWithShape == false
                ? "PUB_GRADIENT_MAPPING_UNSUPPORTED" : "PUB_GRADIENT_INVALID";
            _context.Add(code, failure + " Its primary color was retained.", OfficeConversionLossKind.Approximation, location);
            return;
        }
        shape.FillGradient = fill!.Gradient;
        // A common shape opacity keeps native transparency exact when the
        // foreground/background ratio is representable by RGBA stop alpha.
        shape.FillOpacity = fill.Opacity;
        if (style.FillShadeType.GetValueOrDefault(0x40000003) != 0)
            _context.Add("PUB_GRADIENT_INTERPOLATION_APPROXIMATED", "Native gradient color or position correction was approximated with encoded sRGB interpolation.",
                OfficeConversionLossKind.Approximation, location);
        bool unqualifiedBackgroundOpacity = style.FillGradientStops.Count > 0
            && (style.FillBackOpacity ?? 1) != (style.FillOpacity ?? 1);
        if (fill.OpacityApproximated || unqualifiedBackgroundOpacity) {
            string message = fill.OpacityApproximated ? "Some native gradient stop alpha values required RGBA quantization. " : "";
            if (unqualifiedBackgroundOpacity) message += "A multi-color native gradient uses foreground opacity; separate background opacity requires native qualification.";
            _context.Add("PUB_GRADIENT_OPACITY_APPROXIMATED", message.Trim(), OfficeConversionLossKind.Approximation, location);
        }
        if (style.FillRotatesWithShape != true && HasDirectionalTransform(source))
            _context.Add("PUB_GRADIENT_TRANSFORM_APPROXIMATED", "The gradient follows its rotated or reflected frame. A native fill that does not follow the shape requires additional projection.",
                OfficeConversionLossKind.Approximation, location);
    }

    private static bool HasDirectionalTransform(PublisherEscherShape source) =>
        source.Transform.RotationDegrees.GetValueOrDefault() % 360 != 0
        || source.Transform.FlipHorizontal || source.Transform.FlipVertical
        || source.GroupTransform.M12 != 0 || source.GroupTransform.M21 != 0
        || source.GroupTransform.M11 < 0 || source.GroupTransform.M22 < 0;
}
