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
    private static void AddResolvedBackground(PowerPointSlide slide, OfficeDrawing drawing,
        List<OfficeImageExportDiagnostic> diagnostics, PowerPointEmbeddedImageReadBudget imageBudget) {
        PowerPointSlideBackground background;
        try {
            int readLimit = imageBudget.GetBackgroundReadLimit();
            background = slide.GetBackground(readLimit);
            if (background.Kind == PowerPointSlideBackgroundKind.Image) {
                imageBudget.Consume(background.ImageByteLength);
            }
        } catch (Exception exception) when (exception is InvalidOperationException or InvalidDataException) {
            AddDiagnostic(
                diagnostics,
                imageBudget.IsExhausted ? OfficeDiagnosticCodes.EmbeddedImageLimit : PowerPointImageExportDiagnosticCodes.InvalidSlideBackgroundImage,
                "Used the configured fallback background because the PowerPoint slide background image exceeded the export limit.",
                OfficeConversionLossKind.Omission);
            return;
        }
        switch (background.Kind) {
            case PowerPointSlideBackgroundKind.None:
                return;
            case PowerPointSlideBackgroundKind.SolidColor:
                if (TryParseOfficeColor(background.Color, out OfficeColor color)) {
                    AddBackgroundRectangle(drawing, color, fillGradient: null);
                } else {
                    AddDiagnostic(
                        diagnostics,
                        PowerPointImageExportDiagnosticCodes.InvalidSlideBackgroundColor,
                        "Used the configured fallback background because the PowerPoint slide background color could not be parsed.",
                        OfficeConversionLossKind.Omission);
                }
                return;
            case PowerPointSlideBackgroundKind.LinearGradient:
                if (TryCreateGradient(background, out OfficeLinearGradient? gradient)) {
                    AddBackgroundRectangle(drawing, gradient!.Stops[0].Color, gradient);
                } else {
                    AddDiagnostic(
                        diagnostics,
                        PowerPointImageExportDiagnosticCodes.InvalidSlideBackgroundGradient,
                        "Used the configured fallback background because the PowerPoint slide background gradient colors could not be parsed.",
                        OfficeConversionLossKind.Omission);
                }
                return;
            case PowerPointSlideBackgroundKind.Image:
                AddBackgroundImage(drawing, background, diagnostics);
                return;
            case PowerPointSlideBackgroundKind.Unsupported:
                AddDiagnostic(
                    diagnostics,
                    PowerPointImageExportDiagnosticCodes.UnsupportedSlideBackground,
                    "Used the configured fallback background because " + (background.UnsupportedReason ?? "the PowerPoint slide background is not supported."),
                    OfficeConversionLossKind.Omission);
                return;
        }
    }

    private static void AddBackgroundImage(OfficeDrawing drawing, PowerPointSlideBackground background, List<OfficeImageExportDiagnostic> diagnostics) {
        byte[]? bytes = background.ImageBytes;
        if (bytes == null || bytes.Length == 0) {
            AddDiagnostic(
                diagnostics,
                PowerPointImageExportDiagnosticCodes.InvalidSlideBackgroundImage,
                "Used the configured fallback background because the PowerPoint slide background image bytes could not be read.",
                OfficeConversionLossKind.Omission);
            return;
        }

        OfficeImageProjection projection = CreateImageProjection(
            0D,
            0D,
            drawing.Width,
            drawing.Height,
            background.ImageCropLeft,
            background.ImageCropTop,
            background.ImageCropRight,
            background.ImageCropBottom,
            rotationDegrees: 0D,
            flipHorizontal: false,
            flipVertical: false);
        TryAddImage(drawing, bytes, background.ImageContentType, projection, "PowerPoint slide background image", diagnostics, null);
    }

    private static void AddBackgroundRectangle(OfficeDrawing drawing, OfficeColor fillColor, OfficeLinearGradient? fillGradient) {
        OfficeShape shape = OfficeShape.Rectangle(drawing.Width, drawing.Height);
        shape.FillColor = fillColor;
        shape.FillGradient = fillGradient?.Clone();
        shape.StrokeWidth = 0D;
        drawing.AddShape(shape, 0D, 0D);
    }

    private static bool TryCreateGradient(PowerPointSlideBackground background, out OfficeLinearGradient? gradient) {
        gradient = null;
        if (!TryParseOfficeColor(background.GradientStartColor, out OfficeColor start) ||
            !TryParseOfficeColor(background.GradientEndColor, out OfficeColor end)) {
            return false;
        }

        gradient = OfficeLinearGradient.FromAngle(start, end, background.GradientAngleDegrees ?? 0D);
        return true;
    }

    private static (double Width, double Height) GetSlideSizePoints(PowerPointSlide slide) {
        long widthEmus = 12192000L;
        long heightEmus = 6858000L;
        if (slide.SlidePart.OpenXmlPackage is DocumentFormat.OpenXml.Packaging.PresentationDocument presentationDocument) {
            DocumentFormat.OpenXml.Presentation.SlideSize? size = presentationDocument.PresentationPart?.Presentation?.SlideSize;
            widthEmus = size?.Cx?.Value ?? widthEmus;
            heightEmus = size?.Cy?.Value ?? heightEmus;
        }

        return (PowerPointUnits.ToPoints(widthEmus), PowerPointUnits.ToPoints(heightEmus));
    }
}
