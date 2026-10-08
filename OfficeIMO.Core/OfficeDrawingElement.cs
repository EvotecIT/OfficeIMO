namespace OfficeIMO.Drawing;

/// <summary>
/// Base type for ordered elements inside an <see cref="OfficeDrawing"/> canvas.
/// </summary>
public abstract class OfficeDrawingElement {
    // Detached source identities survive flattening and cloning without changing paint.
    // They are adapter metadata, not a public drawing or accessibility contract.
    internal System.Collections.Generic.IReadOnlyList<string>? SourceElementIds { get; set; }

    internal void RetainSourceElementId(string id) {
        var ids = SourceElementIds ?? System.Array.Empty<string>();
        for (int i = 0; i < ids.Count; i++) if (ids[i] == id) return;
        var combined = new string[ids.Count + 1];
        for (int i = 0; i < ids.Count; i++) combined[i] = ids[i];
        combined[ids.Count] = id;
        SourceElementIds = combined;
    }
    internal abstract OfficeDrawingElement CloneElement();
}
