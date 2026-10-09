using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.OpenDocument;

public static partial class PowerPointOpenDocumentConversionExtensions {
    // Read authored source rectangles before the PowerPoint ratio getters clamp negative or full-image offsets.
    private static Dictionary<uint, A.SourceRectangle> GetPowerPointImageCrops(PresentationPart? presentation, P.SlideId? slideId) {
        var crops = new Dictionary<uint, A.SourceRectangle>();
        if (presentation == null || slideId?.RelationshipId?.Value is not string id ||
            presentation.GetPartById(id) is not SlidePart part || part.Slide == null) return crops;
        foreach (P.Picture picture in part.Slide.Descendants<P.Picture>()) {
            if (picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id?.Value is uint pictureId &&
                picture.BlipFill?.SourceRectangle is A.SourceRectangle crop) crops[pictureId] = crop;
        }
        return crops;
    }

    private static int ApplyPowerPointCrop(A.SourceRectangle source, byte[] imageBytes, string fileName, OdpImage target) {
        if (!TryReadPowerPointCropEdge(source, "l", out double left) || !TryReadPowerPointCropEdge(source, "t", out double top) ||
            !TryReadPowerPointCropEdge(source, "r", out double right) || !TryReadPowerPointCropEdge(source, "b", out double bottom)) return 1;
        if (left == 0 && top == 0 && right == 0 && bottom == 0) return 0;
        if (!OfficeImageReader.TryIdentifyByContent(imageBytes, fileName, out OfficeImageInfo image)) return 1;
        try {
            target.Crop = OdfImageCropCodec.FromSourceFractions(image, left, top, right, bottom);
            return 0;
        } catch (NotSupportedException) {
            return 1;
        }
    }

    private static bool TryReadPowerPointCropEdge(A.SourceRectangle source, string edge, out double value) {
        var attribute = source.GetAttributes().FirstOrDefault(attribute =>
            attribute.LocalName == edge && string.IsNullOrEmpty(attribute.NamespaceUri));
        value = 0;
        if (attribute.LocalName != edge) return true;
        if (!int.TryParse(attribute.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int stored)) return false;
        value = stored / 100000D;
        return true;
    }

    private static int ApplyOdpCrop(OdpImage source, OfficeImageInfo image, PowerPointPicture target) {
        try {
            if (!source.Crop.HasValue) return 0;
            OfficeImageSourceCrop crop = OdfImageCropCodec.ToSourceFractions(source.Crop.Value, image);
            // PowerPoint stores thousandths of a percent. Do not write an empty source window after rounding.
            if (Math.Round(crop.Left * 100000D) + Math.Round(crop.Right * 100000D) >= 100000D ||
                Math.Round(crop.Top * 100000D) + Math.Round(crop.Bottom * 100000D) >= 100000D) return 1;
            target.Crop(crop.Left * 100D, crop.Top * 100D, crop.Right * 100D, crop.Bottom * 100D);
            return 0;
        } catch (Exception exception) when (exception is InvalidDataException or NotSupportedException) {
            return 1;
        }
    }
}
