using PdfCore = OfficeIMO.Pdf;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Pdf;

public static partial class PowerPointPdfConverterExtensions {
    private static int GetImageReadLimit(PowerPointPdfExportOptions options) {
        long remaining = options.MaximumTotalEmbeddedImageBytes - options.EmbeddedImageBytesRead;
        if (remaining <= 0) throw new InvalidOperationException("The aggregate embedded-image byte limit was exceeded.");
        return (int)Math.Min(options.MaximumEmbeddedImageBytes, Math.Min(int.MaxValue, remaining));
    }
}
