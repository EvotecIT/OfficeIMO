using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private readonly Dictionary<(IElement Element, byte[] Source, bool ApplyOrientation),
        (byte[]? Bytes, string ContentType, OfficeImageInfo? Info, bool Rejected)> _imageOrientationResults = new();

    /// <summary>
    /// Retains a source element's completed normalization across flex/grid geometry passes.
    /// Resource bytes are immutable session snapshots; distinct elements and orientation policies
    /// remain distinct work, and neither a failed decode nor a rejected admission is refunded.
    /// </summary>
    private void NormalizeImageOrientation(
        IElement element, bool applyOrientation, string source,
        ref byte[]? bytes, ref string contentType, ref OfficeImageInfo? imageInfo, out bool rejected) {
        rejected = false;
        if (bytes == null) return;
        bool budgeted = _options.ImageNormalizationBudget != null
            && OfficeImageOrientationNormalizer.TryRead(bytes, out OfficeImageOrientation orientation)
            && orientation != OfficeImageOrientation.Normal;
        var key = (element, bytes, applyOrientation);
        if (budgeted && _imageOrientationResults.TryGetValue(key, out var previous)) {
            bytes = previous.Bytes;
            contentType = previous.ContentType;
            imageInfo = previous.Info;
            rejected = previous.Rejected;
            return;
        }
        if (budgeted && !TryAdmitImageOrientationNormalization(bytes, imageInfo, source)) {
            bytes = null;
            rejected = true;
        }
        if (bytes != null && OfficeImageOrientationNormalizer.TryNormalizeToPng(
            bytes, applyOrientation, out byte[] png, out OfficeImageInfo? pngInfo)) {
            bytes = png;
            contentType = "image/png";
            imageInfo = pngInfo;
        }
        if (budgeted) _imageOrientationResults.Add(key, (bytes, contentType, imageInfo, rejected));
    }

    private bool TryAdmitImageOrientationNormalization(byte[] bytes, OfficeImageInfo? imageInfo, string source) {
        HtmlImportBudget budget = _options.ImageNormalizationBudget!;
        long pixels = imageInfo == null ? 0L : (long)imageInfo.Width * imageInfo.Height;
        string detail;
        if (pixels <= 0L || pixels > budget.Limits.MaxDecodedImagePixels) {
            detail = "MaxDecodedImagePixels; requested=" + pixels + "; maximum=" + budget.Limits.MaxDecodedImagePixels;
        } else if (budget.TryBeginImageDecodeWork(bytes.LongLength, out detail)) {
            // Orientation normalization must not flatten a multipage native source before
            // the adapter can enforce its static-image policy. Inspection is also admitted work.
            var decodeOptions = new OfficeRasterDecodeOptions {
                MaximumEncodedBytes = (int)Math.Min(budget.Limits.MaxImageBytes, 128L * 1024L * 1024L),
                MaximumDecodedPixels = Math.Min(budget.Limits.MaxDecodedImagePixels, 50_000_000L),
                FrameLossPolicy = OfficeRasterFrameLossPolicy.RejectMultipleFrames
            };
            if (OfficeRasterContainerInspector.TryInspect(bytes, decodeOptions, out OfficeRasterContainerInfo? container)
                && container != null && container.Frames.Count == 1) return true;
            _diagnostics.Add(ComponentName, HtmlConversionDiagnosticCodes.ResourceDecodeFailed,
                "An embedded image could not be normalized as a supported static image.",
                HtmlDiagnosticSeverity.Warning, source, "Invalid or multiframe orientation source",
                OfficeConversionLossKind.Omission);
            return false;
        }
        _diagnostics.Add(ComponentName, HtmlConversionDiagnosticCodes.TargetLimitExceeded,
            "An embedded image was omitted because its normalization limit was reached.",
            HtmlDiagnosticSeverity.Warning, source, detail, OfficeConversionLossKind.Omission);
        return false;
    }

}
