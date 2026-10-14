using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    /// <summary>Tests an already content-validated raster against the shared drawing/PDF format and transcode bounds.</summary>
    internal static bool IsWithinRasterProjectionProfile(OfficeImageInfo image) =>
        OfficeImagePdfCompatibility.IsSupportedFormat(image.Format) &&
        (image.Format is OfficeImageFormat.Png or OfficeImageFormat.Jpeg ||
            OfficeImagePdfCompatibility.TryValidateTranscodeDimensions(image,
                OfficeImagePdfCompatibility.DefaultMaximumTranscodePixels, out _));

    private static bool TryReadProjectedImage(byte[]? bytes, CancellationToken token, out OfficeImageInfo image, out int svgLosses) {
        image = null!; svgLosses = 0;
        if (bytes == null || !OfficeImageReader.TryValidateContent(bytes, null, token, out image)) return false;
        return image.Format == OfficeImageFormat.Svg
            ? OfficeSvgDrawingReader.TryRead(bytes, new OfficeSvgDrawingReaderOptions { CancellationToken = token }, out _, out svgLosses)
            : IsWithinRasterProjectionProfile(image);
    }

    private static OfficeImageProjection? ProjectImageCrop(OdgShape shape, OfficeImageInfo image, double width, double height, OdfConversionReport report) {
        string feature = "shape:" + shape.Name + ":image-crop";
        try {
            OdfInsets? crop = shape.Crop;
            OfficeImageProjection? projection = OdfImageCropCodec.Project(crop, image, width, height);
            if (crop.HasValue) report.Add(feature, OdfConversionMappingStatus.Converted,
                message: projection.HasValue ? "Intrinsic crop offsets retain source pixels and empty margins while scaling into the frame." : "The crop window has no intersection with image pixels; the frame remains empty.");
            return projection;
        } catch (Exception exception) when (exception is FormatException || exception is ArgumentException || exception is InvalidDataException || exception is NotSupportedException || exception is OverflowException) {
            report.Add(feature, OdfConversionMappingStatus.Unsupported, message: exception.Message + " The image uses an uncropped projection fallback.");
            return new OfficeImageProjection(new OfficeImagePlacement(0, 0, width, height));
        }
    }

    private static OfficeImageProjection? ProjectImageMirror(OdgShape shape, OfficeImageProjection? projection, double width, double height, OdfConversionReport report) {
        string feature = "shape:" + shape.Name + ":image-mirror";
        try {
            OdfImageMirror? mirror = shape.Mirror;
            if (!mirror.HasValue || mirror == OdfImageMirror.None) return projection;
            if (mirror != OdfImageMirror.Horizontal)
                throw new NotSupportedException("Vertical and page-dependent native image mirroring are outside this projection profile.");
            report.Add(feature, OdfConversionMappingStatus.Converted, message: "Horizontal mirroring reflects the cropped source and empty margins inside the original frame.");
            if (!projection.HasValue) return null;
            OfficeImageProjection crop = projection.Value;
            return new OfficeImageProjection(crop.Placement, crop.SourceCrop, rotationCenterX: width / 2D, rotationCenterY: height / 2D, flipHorizontal: true);
        } catch (Exception exception) when (exception is InvalidDataException || exception is NotSupportedException) {
            report.Add(feature, OdfConversionMappingStatus.Unsupported, message: exception.Message + " The image uses an unmirrored projection fallback.");
            return projection;
        }
    }
}
