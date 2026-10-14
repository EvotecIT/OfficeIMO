using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    private bool[]? FindMissingEmSpaces(IReadOnlyList<int> glyphs, IReadOnlyList<int> scalars) {
        if (!OfficeMissingSpaceMetrics.ContainsUnmappedEmSpace(glyphs, scalars)) return null;
        return OfficeMissingSpaceMetrics.FindMissingEmSpaces(glyphs, scalars,
            scalar => MapGlyph(scalar) != 0 || HasPaintedNotdefMapping(scalar));
    }

    private int TextAdvanceWidth(
        ushort glyph,
        bool synthesizedSpace,
        OfficeTrueTypeVariations.WorkBudget? workBudget,
        CancellationToken cancellationToken) => synthesizedSpace
            ? _unitsPerEm
            : AdvanceWidth(glyph, workBudget, cancellationToken);
}
