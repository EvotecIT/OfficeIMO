using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Keeps stored samples unchanged while exposing their display dimensions to every layout consumer.</summary>
    private static PreparedImage PrepareImageWithOrientation(byte[] data, OfficeImageInfo storedInfo,
        PdfWriter.PdfImageStream stream, OfficeImageOrientation orientation,
        bool wasTranscoded = false, OfficeImageFormat? sourceFormat = null) {
        stream.Orientation = orientation;
        OfficeImageInfo layoutInfo = orientation >= OfficeImageOrientation.Transpose
            ? new OfficeImageInfo(storedInfo.Format, storedInfo.Height, storedInfo.Width, storedInfo.DpiY, storedInfo.DpiX) {
                JpegComponentCount = storedInfo.JpegComponentCount,
                TiffPhotometricInterpretation = storedInfo.TiffPhotometricInterpretation
            }
            : storedInfo;
        return new PreparedImage(data, layoutInfo, sourceFormat ?? storedInfo.Format, wasTranscoded, stream);
    }
}
