using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

// Owned adapters supply semantics independently of paint order. No source format
// types or inferred roles enter the PDF writer.
internal sealed class PdfCanvasStructureStep {
    internal PdfCanvasStructureStep(PdfCanvasStructureRole role, PdfCanvasStructureOptions options) {
        Role = role; Options = options.Clone();
    }
    internal PdfCanvasStructureRole Role { get; }
    internal PdfCanvasStructureOptions Options { get; }
}

internal sealed class PdfCanvasSourceContent {
    internal PdfCanvasSourceContent(IReadOnlyList<PdfCanvasStructureStep> path, long order, bool artifact = false) {
        Path = path; LogicalOrder = order; Artifact = artifact;
    }
    internal IReadOnlyList<PdfCanvasStructureStep> Path { get; }
    internal long LogicalOrder { get; }
    internal bool Artifact { get; }
}

public sealed partial class PdfPageCanvas {
    /// <summary>Retains an authored semantic container independently of emitted paint or text.</summary>
    internal void SourceStructureContainer(IReadOnlyList<PdfCanvasStructureStep> path) {
        IReadOnlyList<PdfCanvasItem> items = Array.Empty<PdfCanvasItem>();
        for (int i = path.Count - 1; i >= 0; i--) {
            items = new PdfCanvasItem[] { new PdfCanvasStructureItem(path[i].Role, path[i].Options.Clone(), items) };
        }
        foreach (var item in items) _items.Add(item);
    }
    /// <summary>Preserves drawing paint order while attaching adapter-supplied structure to retained source identities.</summary>
    internal void SourceStructuredDrawing(OfficeDrawing drawing, double width, double height,
        IReadOnlyDictionary<string, PdfCanvasSourceContent> sourceStructure) {
        Drawing(drawing, 0, 0, width, height);
        ((PdfCanvasDrawingItem)_items[_items.Count - 1]).SourceStructure = sourceStructure;
    }
}
