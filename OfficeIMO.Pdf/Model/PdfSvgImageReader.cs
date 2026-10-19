using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Validates SVG image resources consistently at admission and serialization.</summary>
internal static class PdfSvgImageReader {
    internal static OfficeDrawing Read(byte[] bytes, CancellationToken cancellationToken = default) {
        var options = new OfficeSvgDrawingReaderOptions { CancellationToken = cancellationToken };
        if (!OfficeSvgDrawingReader.TryRead(bytes, options, out OfficeDrawing? vector, out int unsupported) ||
            vector == null || unsupported != 0) {
            throw new NotSupportedException("OfficeIMO.Pdf cannot import this SVG image completely. Supply an image produced by a renderer supporting its SVG features.");
        }
        return vector;
    }
}
