namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    // Operation-level SVG text policy, set through the shared image-export options.
    internal bool SvgTextAsPaths { get; set; }
}
