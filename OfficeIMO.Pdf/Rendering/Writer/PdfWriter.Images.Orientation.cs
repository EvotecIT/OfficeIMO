using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Maps stored image coordinates into its EXIF-oriented unit square, in PDF bottom-left coordinates.</summary>
    internal static OfficeTransform GetImageOrientationTransform(OfficeImageOrientation orientation) => orientation switch {
        OfficeImageOrientation.MirrorHorizontal => new OfficeTransform(-1, 0, 0, 1, 1, 0),
        OfficeImageOrientation.Rotate180 => new OfficeTransform(-1, 0, 0, -1, 1, 1),
        OfficeImageOrientation.MirrorVertical => new OfficeTransform(1, 0, 0, -1, 0, 1),
        OfficeImageOrientation.Transpose => new OfficeTransform(0, -1, -1, 0, 1, 1),
        OfficeImageOrientation.Rotate90Clockwise => new OfficeTransform(0, -1, 1, 0, 0, 1),
        OfficeImageOrientation.Transverse => new OfficeTransform(0, 1, 1, 0, 0, 0),
        OfficeImageOrientation.Rotate90CounterClockwise => new OfficeTransform(0, 1, -1, 0, 1, 0),
        _ => OfficeTransform.Identity
    };

    /// <summary>Places encoded samples without decoding, after fitting in the oriented display coordinate system.</summary>
    internal static OfficeTransform GetImagePlacementTransform(PdfImageStream image, OfficeImagePlacement placement) {
        OfficeTransform frame = new OfficeImageProjection(placement).CreateUnitSquareTransform();
        return image.Orientation == OfficeImageOrientation.Normal
            ? frame : GetImageOrientationTransform(image.Orientation).Then(frame);
    }
}
