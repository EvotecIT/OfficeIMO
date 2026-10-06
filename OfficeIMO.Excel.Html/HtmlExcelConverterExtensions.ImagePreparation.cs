using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

public static partial class HtmlExcelConverterExtensions {
    private static bool IsImportableExcelImageType(string contentType) =>
        ExcelSheet.IsSupportedImageContentType(contentType)
        || string.Equals(contentType, "image/webp", StringComparison.OrdinalIgnoreCase)
        || string.Equals(contentType, "image/avif", StringComparison.OrdinalIgnoreCase);

    private static bool TryPrepareExcelImage(
        HtmlImageDataUri dataUri, HtmlToExcelResult result, HtmlImportBudget budget,
        string? source, out byte[] bytes, out string contentType,
        out HtmlImportBudgetReservation reservation) {
        bytes = Array.Empty<byte>();
        contentType = dataUri.MediaType;
        reservation = null!;
        if (!IsSupportedExcelImage(dataUri, result, source)) return false;
        if (!budget.TryReserveImageWithShape(dataUri, out reservation, out string detail)) {
            AddImageLimitDiagnostic(result, source, detail);
            return false;
        }
        bool ready = false;
        try {
            if (!dataUri.TryDecodeBytes(out bytes)) {
                AddImageDecodeDiagnostic(result, source, "The embedded image payload could not be decoded.");
                return false;
            }
            ready = TryNormalizeReservedExcelImage(ref bytes, ref contentType, result, budget, source, ref reservation);
            return ready;
        } finally {
            if (!ready) reservation?.Dispose();
        }
    }

    private static bool TryPrepareExcelImage(
        byte[] imageBytes, string imageContentType, HtmlToExcelResult result, HtmlImportBudget budget,
        string? source, out byte[] bytes, out string contentType,
        out HtmlImportBudgetReservation reservation) {
        bytes = imageBytes;
        contentType = imageContentType;
        reservation = null!;
        if (!IsImportableExcelImageType(contentType)) {
            AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceTypeUnsupported,
                "A layout-region image used an unsupported Excel import image type.",
                lossKind: OfficeConversionLossKind.Omission, source: source, detail: "mediaType=" + contentType);
            return false;
        }
        if (!budget.TryReserveImageWithShape(bytes.LongLength, out reservation, out string detail)) {
            AddImageLimitDiagnostic(result, source, detail);
            return false;
        }
        bool ready = false;
        try {
            ready = TryNormalizeReservedExcelImage(ref bytes, ref contentType, result, budget, source, ref reservation);
            return ready;
        } finally {
            if (!ready) reservation?.Dispose();
        }
    }

    private static bool TryNormalizeReservedExcelImage(
        ref byte[] bytes, ref string contentType, HtmlToExcelResult result, HtmlImportBudget budget,
        string? source, ref HtmlImportBudgetReservation reservation) {
        if (ExcelSheet.IsSupportedImageContentType(contentType)) return true;
        if (!OfficeImageReader.TryIdentifyByContent(bytes, null, out OfficeImageInfo info)
            || info.Format is not (OfficeImageFormat.Webp or OfficeImageFormat.Avif)
            || !string.Equals(info.MimeType, contentType, StringComparison.OrdinalIgnoreCase)) {
            AddImageDecodeDiagnostic(result, source, "The declared raster type did not match identifiable image content.");
            return false;
        }
        long pixels = (long)info.Width * info.Height;
        if (pixels > budget.Limits.MaxDecodedImagePixels) {
            AddImageLimitDiagnostic(result, source, "MaxDecodedImagePixels; requested=" + pixels
                + "; maximum=" + budget.Limits.MaxDecodedImagePixels);
            return false;
        }
        var decodeOptions = new OfficeRasterDecodeOptions {
            MaximumEncodedBytes = (int)Math.Min(128L * 1024 * 1024, budget.Limits.MaxImageBytes),
            MaximumDecodedPixels = Math.Min(50_000_000L, budget.Limits.MaxDecodedImagePixels),
            FrameLossPolicy = OfficeRasterFrameLossPolicy.RejectMultipleFrames
        };
        if (!OfficeImagePngConverter.TryConvertToPng(bytes, decodeOptions, out byte[] png, out OfficeRasterDecodeInfo decoded)) {
            AddImageDecodeDiagnostic(result, source, decoded.Diagnostic ?? "The raster image could not be decoded as a static frame.");
            return false;
        }
        // Bound both the decoded input payload and the retained native PNG.
        reservation.Dispose();
        if (!budget.TryReserveImageWithShape(Math.Max(bytes.LongLength, png.LongLength), out reservation, out string detail)) {
            AddImageLimitDiagnostic(result, source, detail);
            return false;
        }
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ContentApproximated,
            "A static raster image was normalized to PNG for native Excel storage.",
            lossKind: OfficeConversionLossKind.Approximation, source: source,
            detail: "sourceMediaType=" + contentType + "; nativeMediaType=image/png");
        bytes = png;
        contentType = "image/png";
        return true;
    }

    private static void AddImageLimitDiagnostic(HtmlToExcelResult result, string? source, string detail) =>
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
            "An embedded worksheet image was omitted because an image or drawing limit was reached.",
            lossKind: OfficeConversionLossKind.Omission, source: source, detail: detail);

    private static void AddImageDecodeDiagnostic(HtmlToExcelResult result, string? source, string detail) =>
        AddImportDiagnostic(result, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
            "An embedded worksheet image could not be decoded as a supported static image.",
            lossKind: OfficeConversionLossKind.Omission, source: source, detail: detail);
}
