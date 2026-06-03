using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDoc {
    internal static void ValidateClipPathForImage(OfficeClipPath? clipPath, double width, double height) {
        ValidateClipPathInside(clipPath, width, height, nameof(clipPath), "Clip paths must fit inside the image width and height.");
    }

    internal static void ValidateImageStyleForBox(PdfImageStyle style, double width, double height, string clipPathParamName) {
        Guard.NotNull(style, nameof(style));
        Guard.LeftCenterRightAlign(style.Align, nameof(style.Align), "Image");
        ValidateImageFit(style.Fit, nameof(style.Fit));
        ValidateClipPathInside(style.ClipPath, width, height, clipPathParamName, "Clip paths must fit inside the image width and height.");
        if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
            throw new System.ArgumentException("Image spacing before must be a non-negative finite value.", nameof(style));
        }

        if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
            throw new System.ArgumentException("Image spacing after must be a non-negative finite value.", nameof(style));
        }
    }

    internal static void ValidateDrawingStyle(PdfDrawingStyle style, string objectName) {
        Guard.NotNull(style, nameof(style));
        Guard.LeftCenterRightAlign(style.Align, nameof(style.Align), objectName);
        if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
            throw new System.ArgumentException(objectName + " spacing before must be a non-negative finite value.", nameof(style));
        }

        if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
            throw new System.ArgumentException(objectName + " spacing after must be a non-negative finite value.", nameof(style));
        }

        if (style.Decorative && !string.IsNullOrWhiteSpace(style.AlternativeText)) {
            throw new System.ArgumentException(objectName + " style cannot be both decorative and alternate-text bearing.", nameof(style));
        }
    }

    internal static void ValidateImageFit(OfficeImageFit fit, string paramName) {
        if (fit != OfficeImageFit.Stretch && fit != OfficeImageFit.Contain && fit != OfficeImageFit.Cover) {
            throw new System.ArgumentOutOfRangeException(paramName, "Unsupported image fit mode.");
        }
    }

    internal static void ValidateImageFitDimensions(OfficeImageInfo imageInfo, OfficeImageFit fit, string paramName) {
        if (fit == OfficeImageFit.Stretch) {
            return;
        }

        if (imageInfo.Width <= 0 || imageInfo.Height <= 0) {
            throw new System.ArgumentException("Contain and cover image fitting require image dimensions.", paramName);
        }
    }

    private static void ValidatePointInsideShape(OfficePoint point, OfficeShape shape) {
        Guard.NonNegative(point.X, nameof(shape.Points));
        Guard.NonNegative(point.Y, nameof(shape.Points));
        if (point.X > shape.Width || point.Y > shape.Height) {
            throw new System.ArgumentOutOfRangeException(nameof(shape), "Shape points must fit inside the shape width and height.");
        }
    }

    private static void ValidateShapeClipPath(OfficeShape shape) {
        var clipPath = shape.ClipPath;
        ValidateClipPathInside(clipPath, shape.Width, shape.Height, nameof(shape), "Clip paths must fit inside the shape width and height.");
    }

    private static void ValidateClipPathInside(OfficeClipPath? clipPath, double width, double height, string paramName, string fitMessage) {
        if (clipPath == null) {
            return;
        }

        Guard.Positive(clipPath.Width, paramName);
        Guard.Positive(clipPath.Height, paramName);
        if (clipPath.Width > width || clipPath.Height > height) {
            throw new System.ArgumentOutOfRangeException(paramName, fitMessage);
        }

        if (clipPath.Kind == OfficeClipPathKind.RoundedRectangle) {
            Guard.NonNegative(clipPath.CornerRadius, paramName);
            if (clipPath.CornerRadius > System.Math.Min(clipPath.Width, clipPath.Height) / 2D) {
                throw new System.ArgumentOutOfRangeException(paramName, "Clip path corner radius cannot exceed half of the clip path width or height.");
            }
        } else if (clipPath.Kind == OfficeClipPathKind.Path) {
            if (clipPath.Commands.Count == 0 || clipPath.Commands[0].Kind != OfficePathCommandKind.MoveTo) {
                throw new System.ArgumentException("Clip paths require commands starting with MoveTo.", paramName);
            }

            bool hasDraw = false;
            for (int i = 0; i < clipPath.Commands.Count; i++) {
                var command = clipPath.Commands[i];
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                        ValidatePointInsideClip(command.Point, clipPath);
                        break;
                    case OfficePathCommandKind.LineTo:
                        ValidatePointInsideClip(command.Point, clipPath);
                        hasDraw = true;
                        break;
                    case OfficePathCommandKind.CubicBezierTo:
                        ValidatePointInsideClip(command.ControlPoint1, clipPath);
                        ValidatePointInsideClip(command.ControlPoint2, clipPath);
                        ValidatePointInsideClip(command.Point, clipPath);
                        hasDraw = true;
                        break;
                    case OfficePathCommandKind.Close:
                        break;
                    default:
                        throw new System.ArgumentOutOfRangeException(paramName, "Unsupported clip path command kind.");
                }
            }

            if (!hasDraw) {
                throw new System.ArgumentException("Clip paths require at least one drawing command.", paramName);
            }
        } else if (clipPath.Kind != OfficeClipPathKind.Rectangle) {
            throw new System.ArgumentOutOfRangeException(paramName, "Unsupported clip path kind.");
        }
    }

    private static void ValidatePointInsideClip(OfficePoint point, OfficeClipPath clipPath) {
        Guard.NonNegative(point.X, nameof(clipPath.Commands));
        Guard.NonNegative(point.Y, nameof(clipPath.Commands));
        if (point.X > clipPath.Width || point.Y > clipPath.Height) {
            throw new System.ArgumentOutOfRangeException(nameof(clipPath), "Clip path commands must fit inside the clip path width and height.");
        }
    }

    private static void ValidateOpacity(double? opacity, string paramName) {
        if (!opacity.HasValue) {
            return;
        }

        double value = opacity.Value;
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
            throw new System.ArgumentOutOfRangeException(paramName, "Opacity must be a finite number between 0 and 1.");
        }
    }
}
