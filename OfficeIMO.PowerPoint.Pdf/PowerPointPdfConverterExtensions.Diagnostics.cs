using System;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void AddLayoutWarning(
        PowerPointPdfSaveOptions options,
        int slideNumber,
        string code,
        string message,
        PdfCore.PdfLayoutDiagnosticKind kind,
        string source,
        string diagnosticMessage,
        double x,
        double y,
        double width,
        double height) {
        AddWarning(
            options,
            slideNumber,
            code,
            message,
            new PdfCore.PdfLayoutDiagnostic(kind, source, diagnosticMessage, x, y, width, height));
    }

    private static void AddPowerPointListLayoutDiagnostics(PowerPointPdfSaveOptions options, int slideNumber, PptCore.PowerPointTextBox textBox, double x, double y, double width, double height) {
        foreach (PptCore.PowerPointParagraph paragraph in textBox.Paragraphs) {
            if (!HasListMarker(paragraph)) {
                continue;
            }

            if (!paragraph.LeftMarginPoints.HasValue && !paragraph.IndentPoints.HasValue) {
                continue;
            }

            AddLayoutWarning(
                options,
                slideNumber,
                "list-indent-simplified",
                "Rendered a PowerPoint list using PDF text prefixes because explicit PowerPoint list indentation is not yet mapped to PDF hanging-indent layout.",
                PdfCore.PdfLayoutDiagnosticKind.SimplifiedContent,
                "PowerPointList",
                "Explicit PowerPoint list indentation was simplified to a PDF text prefix.",
                x,
                y,
                width,
                height);
            return;
        }
    }

    private static void AddPowerPointPictureAspectRatioDiagnostic(
        PowerPointPdfSaveOptions options,
        int slideNumber,
        byte[] imageBytes,
        PptCore.PowerPointPictureCrop crop,
        OfficeImageFit fit,
        double x,
        double y,
        double width,
        double height) {
        if (!options.WarnOnPictureAspectRatioDistortion ||
            fit != OfficeImageFit.Stretch ||
            crop.HasCrop ||
            width <= 0D ||
            height <= 0D ||
            !OfficeImageReader.TryIdentify(imageBytes, fileName: null, out OfficeImageInfo imageInfo) ||
            imageInfo.Width <= 0 ||
            imageInfo.Height <= 0) {
            return;
        }

        if (!OfficeImagePlacement.ExceedsAspectRatioDistortion(imageInfo.Width, imageInfo.Height, width, height, 1.02D)) {
            return;
        }

        AddLayoutWarning(
            options,
            slideNumber,
            "picture-aspect-distortion",
            "Rendered an uncropped PowerPoint picture with Stretch fit into a frame whose aspect ratio differs from the source image. Set PowerPointPdfSaveOptions.PictureFit to Contain or Cover to preserve aspect ratio.",
            PdfCore.PdfLayoutDiagnosticKind.SimplifiedContent,
            "PowerPointPicture",
            "The mapped PDF picture frame can distort the source image aspect ratio.",
            x,
            y,
            width,
            height);
    }

}
