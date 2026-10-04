using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeTrueTypeFont {
    /// <summary>Reads the native horizontal advance for an already resolved glyph index.</summary>
    internal double FixedGlyphAdvance(int glyphId, double fontSize) {
        if (glyphId < 0 || glyphId >= _numGlyphs) throw new ArgumentOutOfRangeException(nameof(glyphId));
        return AdvanceWidth((ushort)glyphId) * ScaleFor(fontSize);
    }
    /// <summary>Projects an explicitly positioned glyph without Unicode remapping or shaping.</summary>
    internal List<List<OfficePoint>> FixedGlyphContours(int glyphId, double fontSize, double x, double baseline,
        int maximumPoints, CancellationToken cancellationToken) {
        if (glyphId < 0 || glyphId >= _numGlyphs) throw new ArgumentOutOfRangeException(nameof(glyphId));
        if (maximumPoints <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPoints));
        double scale = ScaleFor(fontSize);
        int count = 0;
        return ReadGlyphContours((ushort)glyphId, new FontTransform(scale, 0, 0, -scale, x, baseline), 0,
            _variations?.CreateWorkBudget(), maximumPoints, ref count, cancellationToken, attachmentPoints: null);
    }
}
