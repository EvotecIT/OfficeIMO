using System;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static void ReportOmittedSpeakerNotes(PptCore.PowerPointDocument presentation, PowerPointPdfExportOptions options) {
        if (options.PageLayout != PowerPointPdfPageLayout.Slides && options.IncludeSpeakerNotes) return;
        for (int index = 0; index < presentation.Slides.Count; index++) {
            options.CancellationToken.ThrowIfCancellationRequested();
            PptCore.PowerPointSlide slide = presentation.Slides[index];
            if (!options.IncludesSlide(slide)
                || !slide.Notes.TryGetText(out string notes) || string.IsNullOrWhiteSpace(notes)) continue;
            AddWarning(options, index + 1, "speaker-notes-omitted",
                "Speaker notes are omitted by the selected PDF page layout or IncludeSpeakerNotes setting.",
                lossKind: OfficeConversionLossKind.Omission);
        }
    }

    private static void AddWarning(PowerPointPdfExportOptions options, int slideNumber, string code, string message,
        PdfCore.PdfLayoutDiagnostic? layoutDiagnostic = null, OfficeConversionLossKind? lossKind = null) {
        var warning = lossKind.HasValue
            ? new PowerPointPdfExportWarning(slideNumber, code, message, layoutDiagnostic, lossKind.Value)
            : new PowerPointPdfExportWarning(slideNumber, code, message, layoutDiagnostic);
        options.Warnings.Add(warning);
        options.Report.Add(warning.ToConversionWarning());
    }

    private static void AddLayoutWarning(
        PowerPointPdfExportOptions options,
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

    private static void AddPowerPointPictureAspectRatioDiagnostic(
        PowerPointPdfExportOptions options,
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
            "Rendered an uncropped PowerPoint picture with Stretch fit into a frame whose aspect ratio differs from the source image. Set PowerPointPdfExportOptions.PictureFit to Contain or Cover to preserve aspect ratio.",
            PdfCore.PdfLayoutDiagnosticKind.SimplifiedContent,
            "PowerPointPicture",
            "The mapped PDF picture frame can distort the source image aspect ratio.",
            x,
            y,
            width,
            height);
    }

}
