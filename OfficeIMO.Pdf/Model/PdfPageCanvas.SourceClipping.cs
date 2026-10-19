using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfPageCanvas {
    /// <summary>Retains a source document's rectangular paint boundary without treating that boundary as conversion loss.</summary>
    internal PdfPageCanvas SourceClip(double x, double y, double width, double height, Action<PdfPageCanvas> build) =>
        SourceClip(x, y, OfficeClipPath.Rectangle(width, height), build);

    /// <summary>Retains a source paint clip; additional target clips and text-frame losses remain subject to diagnostics.</summary>
    internal PdfPageCanvas SourceClip(double x, double y, OfficeClipPath clipPath, Action<PdfPageCanvas> build) {
        Clip(x, y, clipPath, build);
        ((PdfCanvasClipItem)_items[_items.Count - 1]).PreservesSourceClip = true;
        return this;
    }
}
