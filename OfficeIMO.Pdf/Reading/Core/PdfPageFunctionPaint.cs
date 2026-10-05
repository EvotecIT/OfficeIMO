using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfPageFunctionPaint {
    internal PdfPageFunctionPaint(PdfFunctionShading resource, OfficeTransform transform) {
        Resource = resource; InverseTransform = transform.Invert();
    }
    internal PdfFunctionShading Resource { get; }
    internal OfficeTransform InverseTransform { get; }
}
