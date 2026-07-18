using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

/// <summary>A reusable Drawing scene plus structured diagnostics for one OneNote page.</summary>
public sealed class OneNotePageVisualSnapshot {
    internal OneNotePageVisualSnapshot(OfficeDrawing drawing, IReadOnlyList<OfficeImageExportDiagnostic> diagnostics) {
        Drawing = drawing;
        Diagnostics = diagnostics;
    }

    /// <summary>Dependency-free vector scene for the page.</summary>
    public OfficeDrawing Drawing { get; }

    /// <summary>Loss or fallback diagnostics produced while mapping page content.</summary>
    public IReadOnlyList<OfficeImageExportDiagnostic> Diagnostics { get; }
}
