using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private IReadOnlyDictionary<string, PdfCanvasSourceContent>? _drawingSourceStructure;
        private int _drawingSourcePaintDepth;
        private PdfCanvasSourceContent? _activeDrawingSourceContent;

        // Metadata selects semantic ownership; dispatch and native effects still
        // execute exactly once, in the drawing's original paint order.
        private void DrawSourceStructuredElement(OfficeDrawingElement element, Action paint) {
            if (_drawingSourceStructure == null || element is OfficeDrawingLink) {
                paint(); return;
            }
            PdfCanvasSourceContent? source = null;
            if (element.SourceElementIds != null) {
                foreach (string id in element.SourceElementIds) {
                    if (_drawingSourceStructure.TryGetValue(id, out var candidate)) {
                        if (source != null && !ReferenceEquals(source, candidate))
                            throw new NotSupportedException("Source structure assigns overlapping paint to different semantic owners.");
                        source = candidate;
                    }
                }
            }
            if (_drawingSourcePaintDepth != 0) {
                if (source != null && !ReferenceEquals(source, _activeDrawingSourceContent))
                    throw new NotSupportedException("Source structure assigns overlapping paint to different semantic owners.");
                paint(); return;
            }
            if (source == null && element is OfficeDrawingGroup or OfficeDrawingEffectGroup) {
                paint(); return;
            }
            PageStructElement? previous = _canvasStructureParentElement;
            try {
                if (source != null && !source.Artifact) {
                    foreach (var step in source.Path) {
                        _canvasStructureParentElement = ResolveCanvasStructure(step.Role, step.Options) ?? _canvasStructureParentElement;
                    }
                    int? id = RegisterTextStructureElement("Span", _canvasStructureParentElement, logicalOrder: source.LogicalOrder);
                    if (id.HasValue) sb.Append("/Span << /MCID ").Append(id.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)).Append(" >> BDC\n");
                    else sb.Append("/Span BMC\n");
                } else sb.Append("/Artifact BMC\n");
                _drawingSourcePaintDepth++;
                var previousSource = _activeDrawingSourceContent;
                _activeDrawingSourceContent = source;
                try { paint(); }
                finally { _drawingSourcePaintDepth--; _activeDrawingSourceContent = previousSource; sb.Append("EMC\n"); }
            } finally { _canvasStructureParentElement = previous; }
        }
    }
}
