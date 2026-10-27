using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static partial class PowerPointSlideImageRenderer {
    private static void AddPicture(OfficeDrawing drawing, PowerPointPicture picture,
        List<OfficeImageExportDiagnostic> diagnostics, PowerPointShapeBoundsMapping mapping,
        PowerPointEmbeddedImageReadBudget imageBudget) {
        if (!TryGetBounds(picture, drawing, diagnostics, mapping, out double left, out double top, out double width, out double height)) {
            return;
        }

        byte[] bytes;
        try {
            int readLimit = imageBudget.GetNextReadLimit();
            bytes = picture.GetImageBytes(readLimit);
            imageBudget.Consume(bytes.Length);
        } catch (Exception exception) when (exception is InvalidOperationException or InvalidDataException) {
            AddUnsupportedShapeDiagnostic(diagnostics, picture, "Skipped a PowerPoint picture because its embedded image bytes could not be read.");
            return;
        }

        if (bytes.Length == 0) {
            AddUnsupportedShapeDiagnostic(diagnostics, picture, "Skipped a PowerPoint picture because its embedded image bytes are empty.");
            return;
        }

        PowerPointPictureCrop crop = picture.GetCrop();
        OfficeImageProjection projection = CreateImageProjection(
            left,
            top,
            width,
            height,
            crop.Left,
            crop.Top,
            crop.Right,
            crop.Bottom,
            picture.Rotation ?? 0D,
            picture.HorizontalFlip == true,
            picture.VerticalFlip == true);
        TryAddImage(drawing, bytes, picture.ContentType, projection, DescribeShape(picture), diagnostics, picture);
    }

    private static OfficeImageProjection CreateImageProjection(
        double left,
        double top,
        double width,
        double height,
        double cropLeft,
        double cropTop,
        double cropRight,
        double cropBottom,
        double rotationDegrees,
        bool flipHorizontal,
        bool flipVertical) {
        return new OfficeImageProjection(
            new OfficeImagePlacement(left, top, width, height),
            OfficeImageSourceCrop.FromClampedFractions(cropLeft, cropTop, cropRight, cropBottom),
            rotationDegrees,
            left + (width / 2D),
            top + (height / 2D),
            flipHorizontal,
            flipVertical);
    }

    private static void TryAddImage(
        OfficeDrawing drawing,
        byte[] bytes,
        string? contentType,
        OfficeImageProjection projection,
        string source,
        List<OfficeImageExportDiagnostic> diagnostics,
        PowerPointShape? shape) {
        try {
            drawing.AddImage(bytes, contentType, projection, source);
        } catch (ArgumentOutOfRangeException) {
            if (shape == null) {
                AddDiagnostic(
                    diagnostics,
                    PowerPointImageExportDiagnosticCodes.ImageBoundsUnsupported,
                    "Skipped a PowerPoint image because its projected bounds are outside the slide drawing canvas.",
                    OfficeConversionLossKind.Omission);
            } else {
                AddUnsupportedShapeDiagnostic(diagnostics, shape, "Skipped a PowerPoint image because its projected bounds are outside the slide drawing canvas.");
            }
        }
    }
}
