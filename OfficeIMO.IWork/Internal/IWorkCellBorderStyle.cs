using OfficeIMO.Drawing;

namespace OfficeIMO.IWork.Internal;

/// <summary>Selected or inherited cell-edge declarations; rejected edges suppress fallback independently.</summary>
internal sealed class IWorkCellBorderStyle {
    // Top, right, bottom, left match the native cell-style property order.
    private readonly IWorkBorderDeclaration?[] _sides = new IWorkBorderDeclaration?[4];
    internal IWorkBorderDeclaration? this[int edge] => _sides[edge];
    internal bool HasDeclarations => _sides.Any(side => side != null);
    internal bool FullyReconstructed => _sides.All(side => side?.IsSupported != false);
    internal OfficeBorderBox Borders => new(_sides[3]?.Side, _sides[0]?.Side, _sides[1]?.Side, _sides[2]?.Side);

    internal void Read(IWorkArchiveRecord owner, IWorkWireMessage properties, IWorkSourceReferenceIssueCollector references) {
        for (int edge = 0; edge < 4; edge++) {
            int field = properties.HasField(10 + edge) ? 10 + edge : 4 + edge;
            if (!properties.HasField(field)) continue;
            _sides[edge] = IWorkStrokeReader.Read(owner, properties, field, "11/" + field, references, field < 10);
        }
    }

    internal void Invalidate() {
        for (int edge = 0; edge < 4; edge++) _sides[edge] = IWorkBorderDeclaration.Invalid;
    }
}
