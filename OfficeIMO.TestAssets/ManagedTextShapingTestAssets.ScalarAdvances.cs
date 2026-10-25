using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    // Distinct cmap glyphs and advances make source-form/paint-form sizing errors observable.
    internal static byte[] CreateFontWithScalarAdvances(IReadOnlyDictionary<int, ushort> advances,
        IReadOnlyCollection<int>? blankScalars = null) {
        if (advances == null || advances.Count == 0) throw new ArgumentException("At least one advance is required.", nameof(advances));
        int[] scalars = advances.Keys.OrderBy(static scalar => scalar).ToArray();
        var glyphAdvances = new Dictionary<int, ushort>();
        for (int index = 0; index < scalars.Length; index++) glyphAdvances[index + 1] = advances[scalars[index]];
        int[]? blankGlyphIds = blankScalars == null ? null : scalars.Select((scalar, index) => (scalar, index))
            .Where(item => blankScalars.Contains(item.scalar)).Select(item => item.index + 1).ToArray();
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(scalars), glyphCount: scalars.Length + 1,
            glyphAdvances: glyphAdvances, blankGlyphIds: blankGlyphIds);
    }
}
