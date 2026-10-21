using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    /// <summary>Maps DrawingML source-image settings onto the shared PDF image placement contract.</summary>
    private static PdfCore.PdfImageStyle CreateNativeImageStyle(WordImage? image = null,
        WordToPdfOptions? options = null, string source = "Word image") {
        var style = new PdfCore.PdfImageStyle { ScaleDownToFit = true, Fit = OfficeImageFit.Stretch };
        if (image == null) return style;
        // DrawingML percentages use 100000 for 100%; they are not EMU lengths.
        double left = (image.CropLeft ?? 0) / 100000D;
        double top = (image.CropTop ?? 0) / 100000D;
        double right = (image.CropRight ?? 0) / 100000D;
        double bottom = (image.CropBottom ?? 0) / 100000D;
        if (left >= 0D && top >= 0D && right >= 0D && bottom >= 0D && left + right < 1D && top + bottom < 1D) {
            if (left != 0D || top != 0D || right != 0D || bottom != 0D)
                style.SourceCrop = new PdfCore.PdfImageSourceCrop(left, top, right, bottom);
        } else if (options != null) {
            AddNativeExportWarning(options, "NativeImageCropApproximated", source,
                "Image crop outsets or a collapsed source rectangle are outside the supported crop profile; the source image is fitted to its authored frame.");
        }
        // Word uses a top-left coordinate system; PDF image paint uses bottom-left axes.
        style.RotationAngle = -(image.Rotation ?? 0);
        style.HorizontalFlip = image.HorizontalFlip ?? false;
        style.VerticalFlip = image.VerticalFlip ?? false;
        style.AlternativeText = string.IsNullOrWhiteSpace(image.Description) ? null : image.Description;
        return style;
    }
}
